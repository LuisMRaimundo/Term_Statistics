#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Inventário e (fases seguintes) extracção/validação de referências APA 7.

Uso (Fase 0)::

    python utilitarios/gera_referencias.py --inventario \\
        --xlsx UNIFORME_near.xlsx \\
        --saida dados/referencias/inventario.tsv \\
        --raiz-corpus "D:\\corpus_local"

Princípio: nunca inventar metadados. ``--inventario`` só lista obras e uma
heurística indicativa de tipo; não preenche campos APA.
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
    escrever_inventario_tsv,
    relatorio_inventario,
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


def main(argv: list[str] | None = None) -> int:
    _configurar_consola()
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument(
        "--inventario", action="store_true",
        help="Fase 0: agrupar 8_Concordancia por doc_id → inventario.tsv",
    )
    ap.add_argument("--xlsx", type=Path, default=None,
                    help="Excel com folha 8_Concordancia")
    ap.add_argument("--folha", default=None,
                    help="folha (omissão: 8_Concordancia)")
    ap.add_argument(
        "--saida", type=Path, default=None,
        help="TSV de saída (omissão: dados/referencias/inventario.tsv)",
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
        help="não abrir PDFs para a heurística tipo_provavel (só nome)",
    )
    args = ap.parse_args(argv)

    if not args.inventario:
        ap.error("indique exactamente um modo (--inventario por agora)")
    if args.xlsx is None:
        ap.error("--xlsx é obrigatório")
    return cmd_inventario(args)


if __name__ == "__main__":
    raise SystemExit(main())
