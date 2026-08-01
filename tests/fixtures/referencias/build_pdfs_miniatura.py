#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Gera PDFs-miniatura para testes da Fase 1 (extracção com evidência).

Corre::

    python tests/fixtures/referencias/build_pdfs_miniatura.py
"""

from __future__ import annotations

from pathlib import Path

import fitz

OUT = Path(__file__).resolve().parent


def _save(nome: str, builder) -> Path:
    path = OUT / nome
    doc = fitz.open()
    builder(doc)
    doc.save(str(path))
    doc.close()
    return path


def build_artigo(doc: fitz.Document) -> None:
    page = doc.new_page(width=595, height=842)
    # corpo (zona média)
    page.insert_text(
        (72, 120),
        "Texture and Sonata Form in Classical String Quartets",
        fontsize=14,
    )
    page.insert_text((72, 150), "Jane Q. Smith", fontsize=11)
    page.insert_text(
        (72, 180),
        "This article examines homogeneous texture in the classical repertoire.",
        fontsize=10,
    )
    page.insert_text((72, 200), "https://doi.org/10.1234/jmt.2019.001", fontsize=9)
    # rodapé (zona inferior)
    page.insert_text(
        (72, 800),
        "Journal of Music Theory, Vol. 63, No. 2, pp. 145-178, 2019",
        fontsize=8,
    )
    # metadados coerentes com a 1.ª página
    doc.set_metadata({
        "title": "Texture and Sonata Form in Classical String Quartets",
        "author": "Smith, J. Q.",
        "subject": "musicology",
    })


def build_grove(doc: fitz.Document) -> None:
    page = doc.new_page(width=595, height=842)
    page.insert_text((72, 100), "Grove Music Online", fontsize=16)
    page.insert_text((72, 130), "Xenakis, Iannis", fontsize=14)
    page.insert_text((72, 160), "by Hoffmann, Peter", fontsize=11)
    page.insert_text(
        (72, 190),
        "Article published online: 20 January 2001",
        fontsize=10,
    )
    page.insert_text(
        (72, 220),
        "https://doi.org/10.1093/gmo/9781561592630.article.30654",
        fontsize=9,
    )
    page.insert_text(
        (72, 260),
        "Greek composer. His music explores complex textures…",
        fontsize=10,
    )
    # meta lixo de digitalização — não deve ser aceite (incoerente)
    doc.set_metadata({
        "title": "Microsoft Word - Document1",
        "author": "Admin",
    })


def build_vazio(doc: fitz.Document) -> None:
    page = doc.new_page(width=595, height=842)
    page.insert_text(
        (72, 200),
        "scanned page without recoverable bibliographic markers",
        fontsize=10,
    )
    # sem metadados úteis


def main() -> None:
    paths = [
        _save("(2019)_Texture_and_Sonata_Form.pdf", build_artigo),
        _save("Grove_Xenakis_entry.pdf", build_grove),
        _save("scan_sem_metadados.pdf", build_vazio),
    ]
    for p in paths:
        print(f"wrote {p}")


if __name__ == "__main__":
    main()
