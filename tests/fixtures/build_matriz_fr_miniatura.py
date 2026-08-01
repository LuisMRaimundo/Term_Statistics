#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Rebuild ``matriz_fr_miniatura.xlsx`` — FR-only golden fixture.

Kept separate from the EN miniature so French ``texture`` windows do not
enter the EN golden under ``--lingua en``. Corpus warrant: Boulez / Roy /
Vaggione-style attestations in the real KWIC matrix.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

OUT = Path(__file__).resolve().parent / "matriz_fr_miniatura.xlsx"


def _row(node: str, ctx: str, path: str, url: str = "") -> list:
    return ([""] * 5) + [node] + ([""] * 5) + [
        path, url, "raiz", ctx, len(ctx.split()),
    ]


def build_rows() -> list[list]:
    rows: list[list] = []
    add = lambda *a, **k: rows.append(_row(*a, **k))

    # Attributive — texture complexe (cf. corpus FR windows)
    add("texture",
        "la partition offre une texture complexe representee par un "
        "reseau de lignes independantes",
        r"E:\todos os textos\(1963)_Boulez_Releves.pdf")
    # Genitive — complexité de la texture
    add("texture",
        "la complexite de la texture harmonique plus complexe exige "
        "une ecoute analytique soutenue",
        r"E:\todos os textos\(1980)_Roy_Vaggione.pdf")
    # Coordination — heterogeneous conj under fr_core_news_sm
    add("texture",
        "la texture complexe et la dynamique contrastante dans le "
        "passage final de l oeuvre",
        r"E:\todos os textos\(1995)_Textures_FR.pdf")

    assert 3 <= len(rows) <= 8, len(rows)
    return rows


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Neighbor Contexts"
    for r in build_rows():
        ws.append(r)
    wb.save(OUT)
    print(f"wrote {OUT} ({len(build_rows())} rows)")


if __name__ == "__main__":
    main()
