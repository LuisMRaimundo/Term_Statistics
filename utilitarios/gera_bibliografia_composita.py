#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Gera bibliografia APA 7 Compósita (DOCX/MD/TSV) — documento autónomo.

Fonte: Compósitaa_near_revisto_LR_final_v9.xlsx (genuínas → ~650 obras).
Saída: pasta CLASSES TEXTURAIS/COMPÓSITA/
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from textura_apendice import filtrar_genuinas, ler_concordancia  # noqa: E402
from textura.referencias import (  # noqa: E402
    construir_apa7,
    construir_rascunho,
    escrever_apa7_docx,
    escrever_apa7_md,
    escrever_apa7_tsv,
    escrever_rascunho_tsv,
    relatorio_apa7,
    relatorio_rascunho,
)

XLSX = Path(
    r"C:\Users\lmr20\Desktop\Tesaurus e Dicionários\CLASSES TEXTURAIS"
    r"\COMPÓSITA\Compósitaa_near_revisto_LR_final_v9.xlsx"
)
OUT = Path(
    r"C:\Users\lmr20\Desktop\Tesaurus e Dicionários\CLASSES TEXTURAIS\COMPÓSITA"
)
RAIZ = Path(r"E:\todos os textos")
REPO_REFS = ROOT / "dados" / "referencias"


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    print("A ler Excel…", flush=True)
    brutas = ler_concordancia(XLSX)
    gen = filtrar_genuinas(brutas)
    n_doc = int(gen["doc_id"].nunique())
    print(f"genuinas={len(gen)} doc_ids={n_doc}", flush=True)

    print("A extrair (PDFs)…", flush=True)
    rasc = construir_rascunho(
        gen, xlsx=XLSX, raiz_corpus=RAIZ, permitir_web=False)
    print(relatorio_rascunho(rasc), flush=True)
    REPO_REFS.mkdir(parents=True, exist_ok=True)
    escrever_rascunho_tsv(rasc, REPO_REFS / "referencias_rascunho_composita.tsv")
    escrever_rascunho_tsv(rasc, OUT / "referencias_rascunho_composita.tsv")

    print("A formatar APA 7…", flush=True)
    apa = construir_apa7(rasc)
    print(relatorio_apa7(apa), flush=True)
    escrever_apa7_tsv(apa, REPO_REFS / "referencias_apa7_composita.tsv")
    escrever_apa7_tsv(apa, OUT / "referencias_apa7_composita.tsv")
    escrever_apa7_md(
        apa,
        OUT / "Referencias_APA7_Composita.md",
        titulo="Referências — Compósita (rascunho APA 7)",
    )
    docx = escrever_apa7_docx(
        apa,
        OUT / "Referencias_APA7_Composita.docx",
        titulo="Referências — Compósita (rascunho APA 7)",
        nota=(
            "Bibliografia autónoma (não integrada no apêndice de concordância). "
            "Gerada a partir de Compósitaa_near_revisto_LR_final_v9.xlsx "
            "(atribuições nucleares/genuínas). Extracção automática com "
            "evidência por campo — sem inventar metadados. Revisão humana "
            "obrigatória antes de uso dissertativo. Fase 2 (validação) e "
            "Fases 4–5 (integração no apêndice / documentação) omitidas a pedido."
        ),
    )
    print(f"DOCX {docx}", flush=True)
    print("DONE", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
