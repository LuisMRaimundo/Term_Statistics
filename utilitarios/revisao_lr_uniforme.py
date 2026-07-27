#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Revisão fase 1 (LR) de UNIFORME_near.xlsx — conforme GUIA_REVISAO_FASE1.md."""

from __future__ import annotations

import re
import shutil
from pathlib import Path

import pandas as pd

SRC = Path(
    r"C:\Users\lmr20\Desktop\Tesaurus e Dicionários"
    r"\CLASSES TEXTURAIS\UNIFORM\UNIFORME_near.xlsx"
)
OUT = SRC.with_name("UNIFORME_near_revisto_LR.xlsx")
PRIOR = Path(__file__).resolve().parents[1] / "UNIFORME_near_revisto_LR.xlsx"

NUCLEARES = {
    "atributiva",
    "predicativa",
    "predicativa_secundaria",
    "nominal_composto",
    "nominal_genitiva",
    "adverbial",
}
NAO_NUCLEARES = {
    "incidental",
    "adverbial_verbal",
    "adverbial_de_grau",
    "coordenada",
    "indeterminada",
}
REVISOR = "LR"


def _truthy(v) -> bool:
    if isinstance(v, bool):
        return v
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    return str(v).strip().lower() in {"1", "true", "verdadeiro", "yes", "sim", "t", "y"}


def _empty(v) -> bool:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return True
    return str(v).strip() == "" or str(v).strip().lower() == "nan"


def main() -> int:
    if not SRC.exists():
        raise SystemExit(f"Fonte ausente: {SRC}")

    xl = pd.ExcelFile(SRC)
    sheets = {n: pd.read_excel(xl, sheet_name=n) for n in xl.sheet_names}
    df = sheets["8_Concordancia"].copy()
    for col in ("nota_revisao", "revisto_por_humano", "motivo_exclusao", "polaridade", "eixo"):
        if col in df.columns:
            df[col] = df[col].astype(object)
    cfg = sheets.get("Config_lexico", pd.DataFrame())

    pol_map = {}
    eixo_map = {}
    if not cfg.empty and "etiqueta" in cfg.columns:
        for _, r in cfg.iterrows():
            et = str(r["etiqueta"]).strip()
            if not _empty(r.get("polaridade")):
                pol_map[et] = str(r["polaridade"]).strip()
            if not _empty(r.get("eixo")):
                eixo_map[et] = str(r["eixo"]).strip()

    # Carry-over from previous LR revision (same hit_key only)
    prior_map = {}
    if PRIOR.exists():
        pr = pd.read_excel(PRIOR, sheet_name="8_Concordancia")
        for _, r in pr.iterrows():
            prior_map[str(r["hit_key"])] = r

    stats = {
        "coerencia_relacao": 0,
        "duplicado_mesmo_tipo": 0,
        "polaridade": 0,
        "eixo": 0,
        "carry_prior": 0,
        "marcadas_LR": 0,
    }

    # --- C: nuclear ↔ relação ---
    for i, row in df.iterrows():
        rel = str(row.get("relacao_sintactica") or "").strip()
        nuc = _truthy(row.get("nuclear"))
        if rel in NAO_NUCLEARES and nuc:
            df.at[i, "nuclear"] = False
            if _empty(row.get("motivo_exclusao")):
                df.at[i, "motivo_exclusao"] = rel
            df.at[i, "nota_revisao"] = (
                (str(row["nota_revisao"]) + " | " if not _empty(row.get("nota_revisao")) else "")
                + "LR: relação não nuclear → exclusão"
            )
            stats["coerencia_relacao"] += 1
        elif rel in NAO_NUCLEARES and not nuc and _empty(row.get("motivo_exclusao")):
            df.at[i, "motivo_exclusao"] = rel
            stats["coerencia_relacao"] += 1

    # --- B: same-passagem duplicates of the SAME termo_tipo (keep best form) ---
    # Different terms in one passagem stay (guide: continuous + homogeneous = OK).
    gids = df.loc[
        df["grupo_passagem_id"].notna() & (df["nuclear"].map(_truthy)),
        "grupo_passagem_id",
    ].unique()
    for gid in gids:
        block = df[(df["grupo_passagem_id"] == gid) & (df["nuclear"].map(_truthy))]
        if len(block) < 2:
            continue
        for tipo, sub in block.groupby(block["termo_tipo"].astype(str)):
            if len(sub) < 2:
                continue
            # Prefer longer matched_form, then smaller distancia, then hit_key.
            cand = list(sub.index)

            def score(idx):
                r = df.loc[idx]
                return (
                    -len(str(r["matched_form"])),
                    float(r["distancia"]) if pd.notna(r["distancia"]) else 99.0,
                    str(r["hit_key"]),
                )

            cand.sort(key=score)
            keep = cand[0]
            for idx in cand[1:]:
                df.at[idx, "nuclear"] = False
                df.at[idx, "motivo_exclusao"] = "janela_sobreposta"
                prev = (
                    ""
                    if _empty(df.at[idx, "nota_revisao"])
                    else str(df.at[idx, "nota_revisao"]) + " | "
                )
                df.at[idx, "nota_revisao"] = (
                    prev
                    + f"LR: duplicado mesmo tipo '{tipo}' em {gid}; "
                    + f"mantida forma '{df.at[keep, 'matched_form']}'"
                )
                stats["duplicado_mesmo_tipo"] += 1

    # --- D: polaridade / eixo from config / base / class default ---
    for i, row in df.iterrows():
        if not _truthy(row.get("nuclear")):
            continue
        tipo = str(row.get("termo_tipo") or "").strip()
        if _empty(row.get("polaridade")):
            fill = None
            if not _empty(row.get("polaridade_base")):
                fill = str(row["polaridade_base"]).strip()
            elif tipo in pol_map:
                fill = pol_map[tipo]
            else:
                fill = "estabilidade"  # classe UNIFORME
            df.at[i, "polaridade"] = fill
            stats["polaridade"] += 1
        if _empty(row.get("eixo")):
            if tipo in eixo_map:
                df.at[i, "eixo"] = eixo_map[tipo]
                stats["eixo"] += 1
            elif _empty(row.get("eixo")):
                df.at[i, "eixo"] = "ambos"
                stats["eixo"] += 1

    # --- prior LR carry-over (stronger human motives on shared keys) ---
    strong = {
        "termo_nao_relacionado_directamente_com_textura",
        "fora_da_classe_uniforme",
        "metatexto",
        "ruido_ocr",
        "citacao_repetida",
    }
    for i, row in df.iterrows():
        hk = str(row["hit_key"])
        if hk not in prior_map:
            continue
        pr = prior_map[hk]
        motivo_p = "" if _empty(pr.get("motivo_exclusao")) else str(pr["motivo_exclusao"]).strip()
        if motivo_p in strong:
            if _truthy(row.get("nuclear")) or _empty(row.get("motivo_exclusao")):
                df.at[i, "nuclear"] = False
                df.at[i, "motivo_exclusao"] = motivo_p
                prev = "" if _empty(df.at[i, "nota_revisao"]) else str(df.at[i, "nota_revisao"]) + " | "
                df.at[i, "nota_revisao"] = prev + "LR: motivo transportado da revisão anterior"
                stats["carry_prior"] += 1

    # --- F: mark review ---
    for i in df.index:
        df.at[i, "revisto_por_humano"] = REVISOR
        stats["marcadas_LR"] += 1

    # Rebuild 9_Excluidas from concordance
    excl = df.loc[~df["nuclear"].map(_truthy)].copy()
    sheets["8_Concordancia"] = df
    sheets["8_Concordancia_Hits"] = df.copy()
    sheets["9_Excluidas"] = excl

    # Update Dominios_por_rever (still por_rever if unchanged)
    if "Dominios_por_rever" in sheets and "dominio" in df.columns:
        dom = (
            df.groupby("caminho_ficheiro", dropna=False)
            .agg(n_hits=("hit_key", "count"), n_nucleares=("nuclear", lambda s: int(s.map(_truthy).sum())))
            .reset_index()
            .rename(columns={"caminho_ficheiro": "caminho"})
        )
        # only files still needing domain
        still = df[df["dominio"].astype(str).str.strip().isin({"por_rever", "nan", ""})]
        paths = set(still["caminho_ficheiro"].dropna().astype(str))
        sheets["Dominios_por_rever"] = dom[dom["caminho"].astype(str).isin(paths)].reset_index(drop=True)

    bak = SRC.with_suffix(SRC.suffix + ".bak-antes-LR")
    if not bak.exists():
        shutil.copy2(SRC, bak)

    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False)

    # also refresh project-local copy
    local = Path(__file__).resolve().parents[1] / "UNIFORME_near_revisto_LR.xlsx"
    shutil.copy2(OUT, local)

    n_nuc = int(df["nuclear"].map(_truthy).sum())
    n_tot = len(df)
    print(f"Escrito: {OUT}")
    print(f"Backup fonte: {bak.name}")
    print(f"Cópia projecto: {local}")
    print(f"Linhas: {n_tot} | nucleares TRUE: {n_nuc} | excluídas: {n_tot - n_nuc}")
    print("Edições:", stats)
    print("motivo_exclusao (FALSE):")
    print(df.loc[~df["nuclear"].map(_truthy), "motivo_exclusao"].value_counts(dropna=False).head(12).to_string())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
