#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
textura_triagem.py — filtros de domínio, metatexto e ruído de OCR
================================================================
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

DOMINIO_OMISSAO = "musicologia"
DOMINIOS_VALIDOS = {
    "musicologia", "mir_visao", "matematica", "ciencias_naturais", "outro",
}

RE_METATEXTO = re.compile(
    r"(?i)\b(ieee|arxiv|isbn|doi|proceedings|symposium|trans\.|pp\.|vol\.|"
    r"journal of|this content downloaded from|author\(s\):|filomat)\b"
)
RE_INICIAIS = re.compile(r"(?:[A-Z]\.){2,}")
RE_CABECALHO = re.compile(r"^\d+(?:\.\d+)+\s")  # 2.1.2 ... (exige subnível)
# Legenda só com numeração no INÍCIO do segmento
RE_LEGENDA = re.compile(
    r"(?i)^(fig(?:ure)?|table|example|ex)\.?\s*\d"
)
RE_VOGAL = re.compile(r"[aeiouáéíóúàèìòùâêîôûäëïöü]", re.I)


def carregar_dominios(caminho: Path | None) -> list[tuple[re.Pattern, str]]:
    if caminho is None or not caminho.is_file():
        return []
    regras = []
    for linha in caminho.read_text(encoding="utf-8").splitlines():
        linha = linha.split("#", 1)[0].strip()
        if not linha or "\t" not in linha:
            continue
        pad, dom = linha.split("\t", 1)
        dom = dom.strip().lower()
        if dom not in DOMINIOS_VALIDOS:
            continue
        regras.append((re.compile(pad.strip(), re.I), dom))
    return regras


def classificar_dominio(caminho_ficheiro: str,
                        regras: list[tuple[re.Pattern, str]]) -> str | None:
    nome = str(caminho_ficheiro or "")
    for rx, dom in regras:
        if rx.search(nome):
            return dom
    return None


def e_metatexto(contexto: str) -> bool:
    ctx = str(contexto or "").strip()
    if not ctx:
        return False
    # início do segmento (após espaços)
    if RE_LEGENDA.match(ctx) or RE_CABECALHO.match(ctx):
        return True
    if RE_METATEXTO.search(ctx):
        return True
    if len(RE_INICIAIS.findall(ctx)) >= 2:
        return True
    return False


def e_ruido_ocr(tokens, idx_no: int, idx_termo: int) -> bool:
    a, b = sorted((idx_no, idx_termo))
    for k in range(a + 1, b):
        w = tokens[k][0]
        if w.isdigit():
            return True
        if len(w) > 3 and not RE_VOGAL.search(w):
            return True
    return False


def aplicar_triagem(df: pd.DataFrame, *,
                    regras_dominio: list[tuple[re.Pattern, str]] | None = None,
                    incluir_dominios: set[str] | None = None,
                    col_caminho: str = "caminho_ficheiro") -> tuple[
                        pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    incluir_dominios = incluir_dominios or {DOMINIO_OMISSAO}
    regras_dominio = regras_dominio or []
    col_c = col_caminho if col_caminho in df.columns else "caminho"

    out = df.copy()
    if "motivo_exclusao" not in out.columns:
        out["motivo_exclusao"] = ""
    if "dominio" not in out.columns:
        out["dominio"] = ""

    for i, row in out.iterrows():
        caminho = row.get(col_c, "")
        dom = classificar_dominio(caminho, regras_dominio)
        if dom is None:
            out.at[i, "dominio"] = "por_rever"
        else:
            out.at[i, "dominio"] = dom
            if dom not in incluir_dominios:
                out.at[i, "nuclear"] = False
                if not out.at[i, "motivo_exclusao"]:
                    out.at[i, "motivo_exclusao"] = f"dominio:{dom}"

        ctx = str(row.get("contexto", ""))
        if e_metatexto(ctx):
            out.at[i, "nuclear"] = False
            if not out.at[i, "motivo_exclusao"]:
                out.at[i, "motivo_exclusao"] = "metatexto"

    # Dominios_por_rever ao nível do ficheiro, com nucleares que contribui
    if col_c in out.columns:
        tmp = out.copy()
        tmp["_nuc"] = tmp["nuclear"].astype(bool)
        por_rever = (tmp[tmp["dominio"] == "por_rever"]
                     .groupby(col_c, as_index=False)
                     .agg(n_hits=(col_c, "size"),
                          n_nucleares=("_nuc", "sum"))
                     .rename(columns={col_c: "caminho"})
                     .sort_values(["n_nucleares", "n_hits"],
                                  ascending=False))
    else:
        por_rever = pd.DataFrame(columns=["caminho", "n_hits", "n_nucleares"])

    excl = out.loc[~out["nuclear"].astype(bool)].copy()
    return out, excl, por_rever
