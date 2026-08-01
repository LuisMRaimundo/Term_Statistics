#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Inventário, extracção e formatação APA 7 de referências.

Uso::

    python utilitarios/gera_referencias.py --inventario --xlsx NEAR.xlsx
    python utilitarios/gera_referencias.py --extrair --xlsx NEAR.xlsx \\
        --raiz-corpus "E:\\todos os textos"
    python utilitarios/gera_referencias.py --formatar \\
        --entrada dados/referencias/referencias_rascunho.tsv \\
        --saida dados/referencias/referencias_apa7.tsv \\
        --md dados/referencias/referencias_apa7.md \\
        --docx "C:\\caminho\\Referencias_APA7.docx"

Princípio: nunca inventar metadados.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from textura.referencias import (  # noqa: E402
    chaves_doc_id,
    construir_apa7,
    construir_inventario,
    construir_rascunho,
    escrever_apa7_docx,
    escrever_apa7_md,
    escrever_apa7_tsv,
    escrever_inventario_tsv,
    escrever_rascunho_tsv,
    relatorio_apa7,
    relatorio_inventario,
    relatorio_rascunho,
)


def _configurar_consola() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def cmd_inventario(args: argparse.Namespace) -> int:
    from textura_apendice import ler_concordancia

    xlsx = Path(args.xlsx)
    if not xlsx.is_file():
        print(f"Excel não encontrado: {xlsx}", file=sys.stderr)
        return 2
    df = ler_concordancia(xlsx, folha=args.folha)
    if "doc_id" not in df.columns:
        print("Folha sem coluna doc_id.", file=sys.stderr)
        return 2

    n_ref = len(set(chaves_doc_id(df)))
    raiz = Path(args.raiz_corpus) if args.raiz_corpus else None
    prefixos = tuple(args.prefixo_origem) if args.prefixo_origem else None

    inv = construir_inventario(
        df,
        raiz_corpus=raiz,
        prefixos_origem=prefixos,
        ler_pdf=not args.sem_ler_pdf,
    )
    saida = Path(args.saida) if args.saida else (
        ROOT / "dados" / "referencias" / "inventario.tsv")
    escrever_inventario_tsv(inv, saida)
    print(relatorio_inventario(inv, n_ref), flush=True)
    print(f"escrito: {saida}", flush=True)
    return 0 if len(inv) == n_ref else 1


def cmd_extrair(args: argparse.Namespace) -> int:
    from textura_apendice import ler_concordancia

    xlsx = Path(args.xlsx)
    if not xlsx.is_file():
        print(f"Excel não encontrado: {xlsx}", file=sys.stderr)
        return 2
    df = ler_concordancia(xlsx, folha=args.folha)
    if "doc_id" not in df.columns:
        print("Folha sem coluna doc_id.", file=sys.stderr)
        return 2

    n_ref = len(set(chaves_doc_id(df)))
    raiz = Path(args.raiz_corpus) if args.raiz_corpus else None
    prefixos = tuple(args.prefixo_origem) if args.prefixo_origem else None

    rasc = construir_rascunho(
        df,
        xlsx=xlsx,
        raiz_corpus=raiz,
        prefixos_origem=prefixos,
        permitir_web=bool(args.permitir_web),
    )
    saida = Path(args.saida) if args.saida else (
        ROOT / "dados" / "referencias" / "referencias_rascunho.tsv")
    escrever_rascunho_tsv(rasc, saida)
    print(relatorio_rascunho(rasc), flush=True)
    print(f"escrito: {saida}", flush=True)
    if args.permitir_web:
        print("AVISO: --permitir-web activo (evidencia=doi_org).", flush=True)
    return 0 if len(rasc) == n_ref else 1


def cmd_formatar(args: argparse.Namespace) -> int:
    entrada = Path(args.entrada) if args.entrada else (
        ROOT / "dados" / "referencias" / "referencias_rascunho.tsv")
    if not entrada.is_file():
        print(f"TSV não encontrado: {entrada}", file=sys.stderr)
        return 2
    import pandas as pd
    df = pd.read_csv(entrada, sep="\t", dtype=str).fillna("")
    apa = construir_apa7(df)
    saida = Path(args.saida) if args.saida else (
        ROOT / "dados" / "referencias" / "referencias_apa7.tsv")
    escrever_apa7_tsv(apa, saida)
    print(relatorio_apa7(apa), flush=True)
    print(f"escrito: {saida}", flush=True)

    if args.md:
        p = escrever_apa7_md(apa, Path(args.md), titulo=args.titulo_doc)
        print(f"escrito: {p}", flush=True)
    if args.docx:
        p = escrever_apa7_docx(apa, Path(args.docx), titulo=args.titulo_doc)
        print(f"escrito: {p}", flush=True)
    return 0


def main(argv: list[str] | None = None) -> int:
    _configurar_consola()
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    modo = ap.add_mutually_exclusive_group(required=True)
    modo.add_argument("--inventario", action="store_true")
    modo.add_argument("--extrair", action="store_true")
    modo.add_argument("--formatar", action="store_true",
                      help="Fase 3: TSV rascunho/revisto → APA 7 + citação curta")
    ap.add_argument("--xlsx", type=Path, default=None)
    ap.add_argument("--folha", default=None)
    ap.add_argument("--saida", type=Path, default=None)
    ap.add_argument("--entrada", type=Path, default=None,
                    help="(--formatar) TSV de entrada")
    ap.add_argument("--md", type=Path, default=None,
                    help="(--formatar) pré-visualização Markdown")
    ap.add_argument("--docx", type=Path, default=None,
                    help="(--formatar) bibliografia DOCX autónoma")
    ap.add_argument("--titulo-doc", default="Referências",
                    help="título do documento MD/DOCX")
    ap.add_argument("--raiz-corpus", type=Path, default=None)
    ap.add_argument("--prefixo-origem", action="append", default=None)
    ap.add_argument("--sem-ler-pdf", action="store_true")
    ap.add_argument("--permitir-web", action="store_true")
    args = ap.parse_args(argv)

    if args.inventario or args.extrair:
        if args.xlsx is None:
            ap.error("--xlsx é obrigatório com --inventario/--extrair")
    if args.inventario:
        return cmd_inventario(args)
    if args.extrair:
        return cmd_extrair(args)
    if args.formatar:
        return cmd_formatar(args)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
