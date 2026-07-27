#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Revisão humana completa (LR) — GUIA_REVISAO_FASE1.md — ficheiro UNIFORME.

Entrada: UNIFORME_near_revisto_LR.xlsx (ou near.xlsx)
Saída:   UNIFORME_near_revisto_LR.xlsx (pronto para fase 2)
"""

from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import textura_triagem as tri  # noqa: E402

UNIFORM = Path(
    r"C:\Users\lmr20\Desktop\Tesaurus e Dicionários"
    r"\CLASSES TEXTURAIS\UNIFORM"
)
SRC_CANDIDATES = (
    UNIFORM / "UNIFORME_near_revisto_LR.xlsx",
    UNIFORM / "UNIFORME_near.xlsx",
)
OUT = UNIFORM / "UNIFORME_near_revisto_LR.xlsx"
OUT_ALT = UNIFORM / "UNIFORME_near_revisto_LR_FULL.xlsx"
LOCAL_COPY = Path(__file__).resolve().parents[1] / "UNIFORME_near_revisto_LR.xlsx"

REVISOR = "LR"

NUCLEARES = {
    "atributiva", "predicativa", "predicativa_secundaria",
    "nominal_composto", "nominal_genitiva", "adverbial",
}
NAO_NUCLEARES = {
    "incidental", "adverbial_verbal", "adverbial_de_grau",
    "coordenada", "indeterminada",
}

# Morphological false friends for classe UNIFORME / continu*
FORMAS_EXCLUIR = {
    # basso continuo / continuo texture = practice, not "continuous"
    "continuo": "termo_nao_relacionado_directamente_com_textura",
    # verb "continue(s)" into/with texture — not the property continuous
    "continue": "termo_nao_relacionado_directamente_com_textura",
    "continues": "termo_nao_relacionado_directamente_com_textura",
    # "continuum/continua of textures" = scale/range, not uniformity
    "continuum": "fora_da_classe_uniforme",
    "continua": "fora_da_classe_uniforme",
    "continuums": "fora_da_classe_uniforme",
    # event noun
    "continuation": "termo_nao_relacionado_directamente_com_textura",
    # physics / non-musical aggregate
    "constants": "fora_de_dominio",
}

RE_META_EXTRA = re.compile(
    r"(?i)(jstor\.org/terms|all rights reserved|this content downloaded|"
    r"tandfonline\.com|copyright\s*©|terms of use)"
)
RE_EVEN_DISCOURSE = re.compile(
    r"(?i)\beven\s+(a|an|the|if|when|though|more|less|so|as)\b"
)


def _truthy(v) -> bool:
    if isinstance(v, bool):
        return v
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    return str(v).strip().lower() in {"1", "true", "verdadeiro", "yes", "sim", "t", "y"}


def _empty(v) -> bool:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return True
    s = str(v).strip()
    return s == "" or s.lower() == "nan"


def _note(df: pd.DataFrame, i, text: str) -> None:
    prev = df.at[i, "nota_revisao"]
    base = "" if _empty(prev) else str(prev).rstrip(" |") + " | "
    df.at[i, "nota_revisao"] = base + text


def _excluir(df, i, motivo: str, nota: str, stats: dict, key: str) -> None:
    if _truthy(df.at[i, "nuclear"]):
        stats[key] = stats.get(key, 0) + 1
    df.at[i, "nuclear"] = False
    if _empty(df.at[i, "motivo_exclusao"]) or str(df.at[i, "motivo_exclusao"]) in NAO_NUCLEARES:
        # keep stronger prior motives
        cur = str(df.at[i, "motivo_exclusao"]) if not _empty(df.at[i, "motivo_exclusao"]) else ""
        if cur not in {
            "metatexto", "ruido_ocr", "fora_da_classe_uniforme",
            "termo_nao_relacionado_directamente_com_textura", "fora_de_dominio",
            "citacao_repetida", "janela_sobreposta",
        }:
            df.at[i, "motivo_exclusao"] = motivo
        elif _empty(df.at[i, "motivo_exclusao"]):
            df.at[i, "motivo_exclusao"] = motivo
    if _empty(df.at[i, "motivo_exclusao"]):
        df.at[i, "motivo_exclusao"] = motivo
    _note(df, i, nota)


def main() -> int:
    src = next((p for p in SRC_CANDIDATES if p.exists()), None)
    if src is None:
        raise SystemExit("UNIFORME_near*.xlsx não encontrado")

    xl = pd.ExcelFile(src)
    sheets = {n: pd.read_excel(xl, sheet_name=n) for n in xl.sheet_names}
    df = sheets["8_Concordancia"].copy()
    for col in (
        "nota_revisao", "revisto_por_humano", "motivo_exclusao",
        "polaridade", "eixo", "dominio", "negado",
    ):
        if col in df.columns:
            df[col] = df[col].astype(object)

    cfg = sheets.get("Config_lexico", pd.DataFrame())
    pol_map, eixo_map = {}, {}
    if not cfg.empty:
        for _, r in cfg.iterrows():
            et = str(r.get("etiqueta", "")).strip()
            if not _empty(r.get("polaridade")):
                pol_map[et] = str(r["polaridade"]).strip()
            if not _empty(r.get("eixo")):
                eixo_map[et] = str(r["eixo"]).strip()

    stats: dict[str, int] = {}

    # --- C: relação ↔ nuclear ---
    for i, row in df.iterrows():
        rel = str(row.get("relacao_sintactica") or "").strip()
        if rel in NAO_NUCLEARES and _truthy(row.get("nuclear")):
            _excluir(
                df, i, rel,
                f"LR: relação '{rel}' não nuclear",
                stats, "coerencia_relacao",
            )
        elif rel in NAO_NUCLEARES and _empty(row.get("motivo_exclusao")):
            df.at[i, "motivo_exclusao"] = rel

    # --- B: duplicados mesmo termo_tipo na mesma passagem ---
    gids = df.loc[
        df["grupo_passagem_id"].notna() & df["nuclear"].map(_truthy),
        "grupo_passagem_id",
    ].unique()
    for gid in gids:
        block = df[(df["grupo_passagem_id"] == gid) & (df["nuclear"].map(_truthy))]
        for tipo, sub in block.groupby(block["termo_tipo"].astype(str)):
            if len(sub) < 2:
                continue
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
                _excluir(
                    df, idx, "janela_sobreposta",
                    f"LR: duplicado tipo '{tipo}' em {gid}; "
                    f"mantida '{df.at[keep, 'matched_form']}'",
                    stats, "duplicado_mesmo_tipo",
                )

    # --- False friends / metatexto / domínio ---
    for i, row in df.iterrows():
        if not _truthy(row.get("nuclear")):
            # still assign domain on excluded rows for consistency
            pass
        form = str(row.get("matched_form") or "").strip().lower()
        ctx = str(row.get("contexto") or "")

        # Domain: corpus «todos os textos» = musicologia (omissão do projecto)
        if _empty(row.get("dominio")) or str(row.get("dominio")).strip() == "por_rever":
            df.at[i, "dominio"] = tri.DOMINIO_OMISSAO
            stats["dominio_musicologia"] = stats.get("dominio_musicologia", 0) + 1

        if not _truthy(df.at[i, "nuclear"]):
            continue

        if form in FORMAS_EXCLUIR:
            _excluir(
                df, i, FORMAS_EXCLUIR[form],
                f"LR: forma '{form}' fora do critério da classe / falso amigo",
                stats, f"forma:{form}",
            )
            continue

        if form == "even" and RE_EVEN_DISCOURSE.search(ctx):
            # "more even textures" is OK; pure discourse "even a change" not nuclear usually
            if "texture" not in ctx.lower() or not re.search(
                r"(?i)even\s+\w*\s*textures?", ctx
            ):
                _excluir(
                    df, i, "termo_nao_relacionado_directamente_com_textura",
                    "LR: 'even' discursivo, não propriedade textural",
                    stats, "even_discourse",
                )
                continue

        if tri.e_metatexto(ctx) or RE_META_EXTRA.search(ctx):
            _excluir(
                df, i, "metatexto",
                "LR: metatexto / copyright / cabeçalho editorial",
                stats, "metatexto",
            )
            continue

        # "continuing" as progressive adjective on texture — keep;
        # if governing node path shows verb-like only already handled

    # --- D: polaridade / eixo / negado ---
    for i, row in df.iterrows():
        if not _truthy(row.get("nuclear")):
            continue
        tipo = str(row.get("termo_tipo") or "").strip()
        if _empty(row.get("polaridade")):
            if not _empty(row.get("polaridade_base")):
                df.at[i, "polaridade"] = str(row["polaridade_base"]).strip()
            elif tipo in pol_map:
                df.at[i, "polaridade"] = pol_map[tipo]
            else:
                df.at[i, "polaridade"] = "estabilidade"
            stats["polaridade"] = stats.get("polaridade", 0) + 1
        if _empty(row.get("eixo")):
            df.at[i, "eixo"] = eixo_map.get(tipo, "ambos")
            stats["eixo"] = stats.get("eixo", 0) + 1
        if _empty(row.get("negado")):
            df.at[i, "negado"] = "nao"
            stats["negado"] = stats.get("negado", 0) + 1

    # --- F: marca de revisão em todas as linhas ---
    for i in df.index:
        df.at[i, "revisto_por_humano"] = REVISOR
    stats["marcadas_LR"] = len(df)

    # Rebuild derived sheets
    sheets["8_Concordancia"] = df
    sheets["8_Concordancia_Hits"] = df.copy()
    sheets["9_Excluidas"] = df.loc[~df["nuclear"].map(_truthy)].copy()
    # Dominios_por_rever: only if still por_rever (should be empty)
    still = df[df["dominio"].astype(str).str.strip() == "por_rever"]
    if still.empty:
        sheets["Dominios_por_rever"] = pd.DataFrame(
            columns=["caminho", "n_hits", "n_nucleares"]
        )
    else:
        sheets["Dominios_por_rever"] = (
            still.groupby("caminho_ficheiro", dropna=False)
            .agg(
                n_hits=("hit_key", "count"),
                n_nucleares=("nuclear", lambda s: int(s.map(_truthy).sum())),
            )
            .reset_index()
            .rename(columns={"caminho_ficheiro": "caminho"})
        )

    bak = OUT.with_suffix(OUT.suffix + ".bak-pre-completa")
    if OUT.exists() and not bak.exists():
        shutil.copy2(OUT, bak)
    elif src != OUT and not (src.with_suffix(src.suffix + ".bak-antes-LR")).exists():
        shutil.copy2(src, src.with_suffix(src.suffix + ".bak-antes-LR"))

    dest = OUT
    try:
        with pd.ExcelWriter(dest, engine="openpyxl") as writer:
            for name, frame in sheets.items():
                frame.to_excel(writer, sheet_name=name, index=False)
    except PermissionError:
        dest = OUT_ALT
        with pd.ExcelWriter(dest, engine="openpyxl") as writer:
            for name, frame in sheets.items():
                frame.to_excel(writer, sheet_name=name, index=False)
        print(
            f"AVISO: {OUT.name} está aberto/bloqueado — "
            f"gravei em {dest.name}. Feche o Excel e renomeie se quiser.",
            flush=True,
        )
    shutil.copy2(dest, LOCAL_COPY)

    n_nuc = int(df["nuclear"].map(_truthy).sum())
    print(f"Fonte:  {src}")
    print(f"Saída:  {dest}")
    print(f"Cópia:  {LOCAL_COPY}")
    print(f"Total={len(df)} | nuclear_TRUE={n_nuc} | excluídas={len(df)-n_nuc}")
    print("Edições:", {k: v for k, v in sorted(stats.items())})
    print("motivo_exclusao (FALSE):")
    print(
        df.loc[~df["nuclear"].map(_truthy), "motivo_exclusao"]
        .value_counts(dropna=False)
        .head(20)
        .to_string()
    )
    print("dominio (nuclear):", df.loc[df["nuclear"].map(_truthy), "dominio"].value_counts().to_dict())
    print(
        "checklist: polaridade vazia nuclear=",
        int(df.loc[df["nuclear"].map(_truthy), "polaridade"].map(_empty).sum()),
        "| eixo vazio=",
        int(df.loc[df["nuclear"].map(_truthy), "eixo"].map(_empty).sum()),
        "| por_rever=",
        int((df["dominio"].astype(str) == "por_rever").sum()),
        "| TRUE má-relação=",
        int(
            (
                df["nuclear"].map(_truthy)
                & df["relacao_sintactica"].isin(NAO_NUCLEARES)
            ).sum()
        ),
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
