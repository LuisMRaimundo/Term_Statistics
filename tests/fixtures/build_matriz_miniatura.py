#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Rebuild ``matriz_miniatura.xlsx`` (Neighbor Contexts, no header).

Run from repo root::

    python tests/fixtures/build_matriz_miniatura.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

OUT = Path(__file__).resolve().parent / "matriz_miniatura.xlsx"


def _row(node: str, ctx: str, path: str, url: str = "") -> list:
    return ([""] * 5) + [node] + ([""] * 5) + [
        path, url, "raiz", ctx, len(ctx.split()),
    ]


def build_rows() -> list[list]:
    rows: list[list] = []
    add = lambda *a, **k: rows.append(_row(*a, **k))

    # EN — nuclear-ish attributives / graded / negated
    add("texture",
        "the music gained a continuous non-strophic texture it had not "
        "had either in the frottola or in the school",
        r"E:\todos os textos\(1933)_Literature and Music.pdf")
    add("texture",
        "nearly every orchestral strand into a solid vibrant and "
        "homogeneous texture based on a rich fundamental tone",
        r"E:\todos os textos\(1944)_Music in the Making.pdf")
    add("texture",
        "the passage presents a rather uniform texture of interlocking "
        "string lines throughout the movement",
        r"E:\todos os textos\(1960)_Medida.pdf")
    add("texture",
        "the orchestra does not present a uniform texture in the coda "
        "despite the sustained chords",
        r"E:\todos os textos\(1971)_Orchestration.pdf")
    add("texture",
        "the second subject offers a more homogeneous texture than the "
        "opening tutti of the allegro",
        r"E:\todos os textos\(1980)_Form.pdf")

    # hyphenated compound (requires *uniform in campo_miniatura.txt)
    add("texture",
        "spaces of di-uniform texture are defined as rotationally "
        "invariant fields in the analysis",
        r"E:\todos os textos\(1990)_Analysis.pdf")

    # coordination / associative / into-texture
    add("texture",
        "the combination of texture and dynamics was clear in the final "
        "cadence of the work",
        r"E:\todos os textos\(2001)_Cadence.pdf")
    add("textures",
        "the guitar textures combined with the lyrics of the singer "
        "created a dense surface",
        r"E:\todos os textos\(2005)_Song.pdf")
    add("texture",
        "the voices are combined into a homophonic texture of unusual "
        "clarity and warmth",
        r"E:\todos os textos\(2008)_Choir.pdf")

    # cross-document near-citation (shifted window boundaries)
    ctx_a = (
        "clock which measures duration with which to measure the "
        "degree of textural complexity irvine asks the question about form"
    )
    ctx_b = (
        "ck which measures duration with which to measure the "
        "degree of textural complexity"
    )
    add("texture", ctx_a,
        r"E:\todos os textos\(1960)_Medida_Complexity.pdf")
    add("texture", ctx_b,
        r"E:\todos os textos\Copia__ab12cd34.pdf")

    # within-document overlapping windows (same term)
    add("texture",
        "extent to which a homogeneous texture a texture in which all "
        "the parts share the same musical material and are",
        r"E:\todos os textos\(1955)_Homogeny.pdf")
    add("texture",
        "aping of a modem language extent to which a homogeneous texture "
        "a texture in which all the parts share the same",
        r"E:\todos os textos\(1955)_Homogeny.pdf")

    # extra-musical (geology) — campo term present so a hit is emitted
    add("textures",
        "replacement of bioclasts by a mixture of dolomite and "
        "homogeneous textures in breccia facies of the waulsortian mound",
        r"E:\geology\(2012)_Carbonates.pdf")

    # PT
    add("textura",
        "a obra apresenta uma textura homogenea de cordas no andante e "
        "um contraste timbrico no finale",
        r"E:\todos os textos\(2010)_Textura_PT.pdf")
    add("textura",
        "nao se observa uma textura uniforme no coral apesar da escrita "
        "a quatro vozes",
        r"E:\todos os textos\(2011)_Coral_PT.pdf")
    add("texturas",
        "as texturas constantes do acompanhamento sustentam a melodia "
        "principal sem variar o ritmo",
        r"E:\todos os textos\(2014)_Acomp_PT.pdf")

    # fillers (association / diversity mass)
    add("texture",
        "a stable texture of woodwind chords supports the solo line in "
        "the slow movement",
        r"E:\todos os textos\(1975)_Woodwind.pdf")
    add("texture",
        "the varied texture of the scherzo contrasts with the sustained "
        "chords of the trio",
        r"E:\todos os textos\(1977)_Scherzo.pdf")
    add("texture",
        "an irregular texture of pizzicato fragments interrupts the "
        "otherwise calm surface",
        r"E:\todos os textos\(1982)_Pizz.pdf")
    add("textural",
        "textural density increases as the consistent pulse of the "
        "ostinato thickens the fabric",
        r"E:\todos os textos\(1988)_Density.pdf")
    add("texture",
        "the static texture of the opening is soon abandoned for a "
        "changing web of lines",
        r"E:\todos os textos\(1992)_Static.pdf")
    add("texture",
        "critics praised the sustained texture of the adagio as a model "
        "of orchestral balance",
        r"E:\todos os textos\(1995)_Adagio.pdf")
    add("texture",
        "a regular texture of repeated chords marks the dance refrain "
        "of the suite",
        r"E:\todos os textos\(1998)_Dance.pdf")
    add("texture",
        "the immutable texture of the drone underpins the improvisation "
        "of the soloist",
        r"E:\todos os textos\(2003)_Drone.pdf")
    add("texture",
        "an unchanging texture of muted strings accompanies the spoken "
        "narration throughout",
        r"E:\todos os textos\(2006)_Narration.pdf")
    add("texture",
        "a constant texture of muted brass underpins the funeral march "
        "of the symphony",
        r"E:\todos os textos\(2009)_March.pdf")

    assert 20 <= len(rows) <= 30, len(rows)
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
