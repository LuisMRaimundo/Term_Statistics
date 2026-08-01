# -*- coding: utf-8 -*-
"""Fase 3 — formatação APA 7, citação curta, desambiguação, ordenação."""

from __future__ import annotations

import pandas as pd

from textura.referencias import (
    chave_ordenacao_apa,
    citacao_curta_de,
    construir_apa7,
    formatar_referencia_apa7,
)


def test_formatar_artigo():
    apa = formatar_referencia_apa7({
        "tipo": "artigo",
        "autores": "Smith, J. Q.",
        "ano": "2019",
        "titulo": "Texture and sonata form",
        "contentor": "Journal of Music Theory",
        "volume": "63",
        "numero": "2",
        "paginas": "145-178",
        "doi_ou_url": "10.1234/jmt.2019.001",
    })
    assert "Smith, J. Q. (2019)." in apa
    assert "Texture and sonata form." in apa
    assert "*Journal of Music Theory, 63*" in apa
    assert "(2)" in apa
    assert "145-178" in apa
    assert apa.endswith("https://doi.org/10.1234/jmt.2019.001")
    assert not apa.endswith(".")


def test_formatar_verbete_grove():
    apa = formatar_referencia_apa7({
        "tipo": "verbete",
        "autores": "Hoffmann, P.",
        "ano": "2001",
        "titulo": "Xenakis, Iannis",
        "contentor": "Grove Music Online",
        "editora": "Oxford University Press",
        "doi_ou_url": "https://doi.org/10.1093/gmo/9781561592630.article.30654",
    })
    assert "Hoffmann, P. (2001)." in apa
    assert "In *Grove Music Online*." in apa
    assert "Oxford University Press." in apa
    assert "doi.org/10.1093/gmo" in apa


def test_formatar_livro_tese_actas():
    livro = formatar_referencia_apa7({
        "tipo": "livro", "autores": "Roads, C.", "ano": "2015",
        "titulo": "Composing electronic music", "editora": "Oxford University Press",
    })
    assert "*Composing electronic music*" in livro
    assert "Oxford University Press." in livro

    tese = formatar_referencia_apa7({
        "tipo": "tese", "autores": "Besharse, K. E.", "ano": "n.d.",
        "titulo": "The role of texture in French spectral music",
        "editora": "University of Illinois",
    })
    assert "[Doctoral dissertation, University of Illinois]" in tese

    actas = formatar_referencia_apa7({
        "tipo": "actas", "autores": "Doe, J.", "ano": "2020",
        "titulo": "Granular textures", "contentor": "ICMC Proceedings",
        "paginas": "10-15",
    })
    assert "In *ICMC Proceedings*" in actas
    assert "pp. 10-15" in actas


def test_citacao_curta_1_2_3():
    assert citacao_curta_de("Smith, J.", "2019") == "Smith, 2019"
    assert citacao_curta_de("Smith, J.; Jones, A.", "2019") == "Smith & Jones, 2019"
    assert citacao_curta_de(
        "Smith, J.; Jones, A.; Lee, B.", "2019"
    ) == "Smith et al., 2019"
    assert citacao_curta_de("", "") == "(Autor desconhecido, n.d.)"


def test_desambiguacao_deterministica():
    df = pd.DataFrame([
        {"doc_id": "b", "autores": "Smith, J.", "ano": "2004",
         "titulo": "Zebra textures", "tipo": "artigo",
         "contentor": "JMT", "verificar": "nao"},
        {"doc_id": "a", "autores": "Smith, J.", "ano": "2004",
         "titulo": "Alpha textures", "tipo": "artigo",
         "contentor": "JMT", "verificar": "nao"},
    ])
    out = construir_apa7(df)
    by_id = out.set_index("doc_id")
    assert by_id.loc["a", "citacao_curta"] == "Smith, 2004a"
    assert by_id.loc["b", "citacao_curta"] == "Smith, 2004b"
    assert "(2004a)" in by_id.loc["a", "referencia_apa7"]
    assert "(2004b)" in by_id.loc["b", "referencia_apa7"]


def test_ordenacao_com_acentos():
    assert chave_ordenacao_apa("Álvarez") < chave_ordenacao_apa("Smith")
    assert chave_ordenacao_apa("Çelik") < chave_ordenacao_apa("Delta")
    df = pd.DataFrame([
        {"doc_id": "2", "autores": "Smith, A.", "ano": "2000",
         "titulo": "B", "tipo": "livro", "editora": "X", "verificar": "nao"},
        {"doc_id": "1", "autores": "Álvarez, M.", "ano": "2001",
         "titulo": "A", "tipo": "livro", "editora": "X", "verificar": "nao"},
    ])
    out = construir_apa7(df)
    assert list(out["doc_id"]) == ["1", "2"]
