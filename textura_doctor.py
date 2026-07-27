#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
textura_doctor.py — checklist pré-análise da fase 1 (aditivo)
============================================================

Lê ``8_Concordancia`` de um Excel revisto e imprime avisos/erros via
``textura_triagem.checklist_revisao``. Não altera o ficheiro.

Uso::

    python textura_doctor.py --xlsx UNIFORME_near_revisto_LR.xlsx
    python textura_doctor.py --xlsx ... --estrito   # exit 1 se erros
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

import textura_triagem as ttri


def diagnosticar(xlsx: Path) -> dict:
    with pd.ExcelFile(xlsx) as xl:
        sheets = list(xl.sheet_names)
        if "8_Concordancia" not in sheets:
            raise SystemExit(f"Falta a folha 8_Concordancia em {xlsx}")
        conc = pd.read_excel(xl, sheet_name="8_Concordancia")
    return ttri.checklist_revisao(conc)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--xlsx", type=Path, required=True)
    ap.add_argument(
        "--estrito",
        action="store_true",
        help="código de saída 1 se o checklist tiver erros",
    )
    args = ap.parse_args()
    if not args.xlsx.is_file():
        raise SystemExit(f"Ficheiro inexistente: {args.xlsx}")

    chk = diagnosticar(args.xlsx)
    print(f"Doctor: {args.xlsx.name}", flush=True)
    print(
        f"  n={chk['n']}  nuclear={chk['n_nuclear']}  "
        f"score={chk['score']}  ok={chk['ok']}",
        flush=True,
    )
    for a in chk.get("avisos") or []:
        print(f"  aviso: {a}", flush=True)
    for e in chk.get("erros") or []:
        print(f"  ERRO: {e}", flush=True)

    if args.estrito and not chk.get("ok"):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
