#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
textura_triagem.py — filtros de domínio, metatexto e ruído de OCR
================================================================

Extensões (R95, aditivas — não removem o comportamento anterior):

  • metatexto editorial (JSTOR / copyright / terms of use)
  • falsos amigos morfológicos (ex.: *continuo*, *continuum*, verbos *continue*)
  • domínio omissão configurável (``dominio_omissao``) em vez de ``por_rever``
  • checklist de revisão para a fase 2 (`checklist_revisao`)
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Optional

import pandas as pd

from textura.lexico import (
    DOMINIOS_VALIDOS as DOMINIOS_VALIDOS,
    FALSOS_AMIGOS_FORMAS as FALSOS_AMIGOS_FORMAS,
    RELACOES_NAO_NUCLEARES as RELACOES_NAO_NUCLEARES,
    # Re-export: same frozenset object as ``textura.lexico`` (identity tests).
    RELACOES_NUCLEARES as RELACOES_NUCLEARES,
)

DOMINIO_OMISSAO = "musicologia"

RE_METATEXTO = re.compile(
    r"(?i)\b(ieee|arxiv|isbn|doi|proceedings|symposium|trans\.|pp\.|vol\.|"
    r"journal of|this content downloaded from|author\(s\):|filomat|"
    r"jstor\.org/terms|all rights reserved|tandfonline\.com|terms of use|"
    r"copyright\s*©)\b"
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


def _truthy_nuclear(val) -> bool:
    if isinstance(val, bool):
        return val
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return False
    return str(val).strip().lower() in {"1", "true", "verdadeiro", "yes", "sim", "t", "y"}


def _empty(val) -> bool:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return True
    s = str(val).strip()
    return s == "" or s.lower() in {"nan", "none", "nat"}


def motivo_falso_amigo(matched_form: str) -> Optional[str]:
    """Se a forma for falso amigo conhecido, devolve ``motivo_exclusao``."""
    form = str(matched_form or "").strip().lower()
    return FALSOS_AMIGOS_FORMAS.get(form)


def aplicar_falsos_amigos(df: pd.DataFrame) -> pd.DataFrame:
    """Marca ``nuclear=False`` para falsos amigos (aditivo; não toca o resto)."""
    out = df.copy()
    if "matched_form" not in out.columns:
        return out
    if "motivo_exclusao" not in out.columns:
        out["motivo_exclusao"] = ""
    if "nuclear" not in out.columns:
        return out
    for i, row in out.iterrows():
        if not _truthy_nuclear(row.get("nuclear")):
            continue
        motivo = motivo_falso_amigo(row.get("matched_form"))
        if not motivo:
            continue
        out.at[i, "nuclear"] = False
        if _empty(row.get("motivo_exclusao")):
            out.at[i, "motivo_exclusao"] = motivo
    return out


def checklist_revisao(df: pd.DataFrame) -> dict[str, Any]:
    """Checklist pré-análise (avisos; não bloqueia por si)."""
    if df is None or df.empty:
        return {
            "ok": False, "erros": ["concordância vazia"], "avisos": [],
            "n": 0, "n_nuclear": 0, "score": 0,
        }
    erros: list[str] = []
    avisos: list[str] = []
    n = len(df)
    nuc_mask = df["nuclear"].map(_truthy_nuclear) if "nuclear" in df.columns else pd.Series([False] * n)
    n_nuc = int(nuc_mask.sum())

    if "relacao_sintactica" in df.columns:
        bad_true = df.loc[
            nuc_mask & df["relacao_sintactica"].astype(str).isin(RELACOES_NAO_NUCLEARES)
        ]
        if len(bad_true):
            erros.append(
                f"{len(bad_true)} linhas nuclear=TRUE com relação não nuclear"
            )
        amigos = df.loc[nuc_mask].copy()
        if "matched_form" in amigos.columns:
            ff = amigos["matched_form"].astype(str).str.lower().map(
                lambda f: f in FALSOS_AMIGOS_FORMAS
            )
            if int(ff.sum()):
                avisos.append(
                    f"{int(ff.sum())} nucleares ainda são falsos amigos "
                    f"(continuo/continuum/continue…)"
                )

    if "polaridade" in df.columns and n_nuc:
        vazias = int(df.loc[nuc_mask, "polaridade"].map(_empty).sum())
        if vazias:
            avisos.append(f"{vazias} nucleares sem polaridade")

    if "eixo" in df.columns and n_nuc:
        vazias = int(df.loc[nuc_mask, "eixo"].map(_empty).sum())
        if vazias:
            avisos.append(f"{vazias} nucleares sem eixo")

    if "dominio" in df.columns and n_nuc:
        por = int(
            df.loc[nuc_mask, "dominio"].astype(str).str.strip().eq("por_rever").sum()
        )
        if por:
            avisos.append(f"{por} nucleares com dominio=por_rever")

    if "revisto_por_humano" in df.columns:
        marc = int((~df["revisto_por_humano"].map(_empty)).sum())
        if marc == 0:
            avisos.append("nenhuma linha com revisto_por_humano (GUIA fase 1)")
        elif marc < n:
            avisos.append(f"revisão parcial: {marc}/{n} linhas marcadas")

    # score 0–100 (informativo)
    score = 100
    score -= 25 * len(erros)
    score -= 8 * len(avisos)
    if n_nuc == 0:
        erros.append("0 linhas nucleares — análise ficará vazia")
        score -= 30
    score = max(0, min(100, score))

    return {
        "ok": not erros,
        "erros": erros,
        "avisos": avisos,
        "n": n,
        "n_nuclear": n_nuc,
        "score": score,
    }


def aplicar_triagem(df: pd.DataFrame, *,
                    regras_dominio: list[tuple[re.Pattern, str]] | None = None,
                    incluir_dominios: set[str] | None = None,
                    col_caminho: str = "caminho_ficheiro",
                    dominio_omissao: str | None = None,
                    aplicar_amigos: bool = True) -> tuple[
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
            # R95: omissão opcional (ex. musicologia) em vez de por_rever
            if dominio_omissao and dominio_omissao in DOMINIOS_VALIDOS:
                out.at[i, "dominio"] = dominio_omissao
                if dominio_omissao not in incluir_dominios:
                    out.at[i, "nuclear"] = False
                    if _empty(out.at[i, "motivo_exclusao"]):
                        out.at[i, "motivo_exclusao"] = f"dominio:{dominio_omissao}"
            else:
                out.at[i, "dominio"] = "por_rever"
        else:
            out.at[i, "dominio"] = dom
            if dom not in incluir_dominios:
                out.at[i, "nuclear"] = False
                if _empty(out.at[i, "motivo_exclusao"]):
                    out.at[i, "motivo_exclusao"] = f"dominio:{dom}"

        ctx = str(row.get("contexto", ""))
        if e_metatexto(ctx):
            out.at[i, "nuclear"] = False
            if _empty(out.at[i, "motivo_exclusao"]):
                out.at[i, "motivo_exclusao"] = "metatexto"

    if aplicar_amigos:
        out = aplicar_falsos_amigos(out)

    # Dominios_por_rever ao nível do ficheiro, com nucleares que contribui
    if col_c in out.columns:
        tmp = out.copy()
        tmp["_nuc"] = tmp["nuclear"].map(_truthy_nuclear)
        por_rever = (tmp[tmp["dominio"] == "por_rever"]
                     .groupby(col_c, as_index=False)
                     .agg(n_hits=(col_c, "size"),
                          n_nucleares=("_nuc", "sum"))
                     .rename(columns={col_c: "caminho"})
                     .sort_values(["n_nucleares", "n_hits"],
                                  ascending=False))
    else:
        por_rever = pd.DataFrame(columns=["caminho", "n_hits", "n_nucleares"])

    excl = out.loc[~out["nuclear"].map(_truthy_nuclear)].copy()
    return out, excl, por_rever
