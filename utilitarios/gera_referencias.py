#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Inventário e extracção/validação de referências APA 7.

Uso::

    python utilitarios/gera_referencias.py --inventario \\
        --xlsx UNIFORME_near.xlsx --raiz-corpus "E:\\todos os textos"

    python utilitarios/gera_referencias.py --extrair \\
        --xlsx UNIFORME_near.xlsx --raiz-corpus "E:\\todos os textos"

Princípio: nunca inventar metadados. Campo preenchido exige ``evidencia_*``.
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
    construir_inventario,
    construir_rascunho,
    escrever_inventario_tsv,
    escrever_rascunho_tsv,
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
    if len(inv) != n_ref:
        print(
            f"ERRO: inventário ({len(inv)}) ≠ doc_id esperados ({n_ref})",
            file=sys.stderr,
        )
        return 1
    return 0


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
        print(
            "AVISO: --permitir-web activo — campos vazios puderam ser "
            "preenchidos via doi.org/Crossref (evidencia=doi_org).",
            flush=True,
        )
    if len(rasc) != n_ref:
        print(
            f"ERRO: rascunho ({len(rasc)}) ≠ doc_id esperados ({n_ref})",
            file=sys.stderr,
        )
        return 1
    return 0


def main(argv: list[str] | None = None) -> int:
    _configurar_consola()
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    modo = ap.add_mutually_exclusive_group(required=True)
    modo.add_argument(
        "--inventario", action="store_true",
        help="Fase 0: agrupar 8_Concordancia por doc_id → inventario.tsv",
    )
    modo.add_argument(
        "--extrair", action="store_true",
        help="Fase 1: extracção com evidência → referencias_rascunho.tsv",
    )
    ap.add_argument("--xlsx", type=Path, required=True,
                    help="Excel com folha 8_Concordancia")
    ap.add_argument("--folha", default=None,
                    help="folha (omissão: 8_Concordancia)")
    ap.add_argument(
        "--saida", type=Path, default=None,
        help="TSV de saída (omissão sob dados/referencias/)",
    )
    ap.add_argument(
        "--raiz-corpus", type=Path, default=None,
        help="raiz local que substitui o prefixo E:\\todos os textos (CI)",
    )
    ap.add_argument(
        "--prefixo-origem", action="append", default=None,
        help="prefixo a remapear (repetível; omissão: corpus E: padrão)",
    )
    ap.add_argument(
        "--sem-ler-pdf", action="store_true",
        help="(--inventario) não abrir PDFs para tipo_provavel",
    )
    ap.add_argument(
        "--permitir-web", action="store_true",
        help="(--extrair) resolver DOIs já presentes no PDF via doi.org/Crossref",
    )
    args = ap.parse_args(argv)

    if args.inventario:
        return cmd_inventario(args)
    if args.extrair:
        return cmd_extrair(args)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
