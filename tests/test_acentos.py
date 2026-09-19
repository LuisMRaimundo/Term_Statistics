# -*- coding: utf-8 -*-
"""Casamento PT/FR/ES: léxico com acentos ↔ corpus sem acentos (e o inverso)."""

from __future__ import annotations

import unittest

import textura_lexico as tlex
import textura_query as tq
from textura.tokenizacao import (
    _casa_em,
    _rx_palavra,
    compila_campo,
    encontra_termos,
    sem_diacriticos,
    tokeniza,
)


class TestSemDiacriticos(unittest.TestCase):
    def test_pt_fr_es_ligaduras(self):
        self.assertEqual(sem_diacriticos("estátic"), "estatic")
        self.assertEqual(sem_diacriticos("homogéneo"), "homogeneo")
        self.assertEqual(sem_diacriticos("imóvel"), "imovel")
        self.assertEqual(sem_diacriticos("homogène"), "homogene")
        self.assertEqual(sem_diacriticos("estático"), "estatico")
        self.assertEqual(sem_diacriticos("cœur"), "coeur")
        self.assertEqual(sem_diacriticos("ESTÁTICA"), "estatica")


class TestPesquisaAcentos(unittest.TestCase):
    def test_lexico_acentuado_casa_matriz_ascii(self):
        self.assertTrue(tq.forma_casa_padrao("estatica", "estátic*"))
        self.assertTrue(tq.forma_casa_padrao("homogeneo", "homogéneo"))
        self.assertTrue(tq.forma_casa_padrao("imovel", "imóvel"))
        self.assertTrue(tq.forma_casa_padrao("estaticas", "estátic*"))

    def test_lexico_ascii_casa_corpus_acentuado(self):
        self.assertTrue(tq.forma_casa_padrao("estática", "estatic*"))
        self.assertTrue(tq.forma_casa_padrao("homogéneo", "homogeneo"))
        self.assertTrue(tq.forma_casa_padrao("immobile", "immob*"))

    def test_nao_funde_radicais_diferentes(self):
        self.assertFalse(tq.forma_casa_padrao("static", "estátic*"))
        self.assertFalse(tq.forma_casa_padrao("homogeneous", "homogéneo"))
        self.assertFalse(tq.forma_casa_padrao("stasis", "estátic*"))

    def test_consulta_near_pt_sem_acento_na_matriz(self):
        q = tq.ConsultaBooleana(
            "textur* NEAR/8 estátic*", mesma_frase=True, exigir_sintaxe=False)
        ctx = "a passagem oferece uma textura estatica de cordas"
        toks = tq.tokeniza(tq.normaliza(ctx))
        self.assertTrue(q.avalia(toks, ctx).ok)

    def test_fr_homogene_e_es_estatico(self):
        self.assertTrue(tq.forma_casa_padrao("homogene", "homogène"))
        self.assertTrue(tq.forma_casa_padrao("estatico", "estático"))


class TestNearAcentos(unittest.TestCase):
    def test_rx_e_casa_em(self):
        rx = _rx_palavra("estátic*")
        self.assertTrue(rx.match(sem_diacriticos("estatica")))
        self.assertTrue(rx.match(sem_diacriticos("estática")))
        self.assertFalse(rx.match(sem_diacriticos("static")))

    def test_encontra_termo_acentuado_em_janela_ascii(self):
        campo = compila_campo({"estátic": ["estátic*"]})
        toks = tokeniza("uma textura estatica de sopros")
        ach = encontra_termos(toks, campo)
        self.assertEqual([a["termo_tipo"] for a in ach], ["estátic"])
        self.assertEqual(ach[0]["termo_forma"], "estatica")

    def test_casa_em_dobra_token(self):
        seq = [_rx_palavra("imóvel")]
        toks = tokeniza("o baixo imovel sustenta a textura")
        j = next(i for i, (w, _) in enumerate(toks) if w == "imovel")
        self.assertEqual(_casa_em(toks, j, seq), 1)


class TestCanonicoAcentos(unittest.TestCase):
    campo = {
        "homoge": ["homoge*"],
        "homogéneo": ["homogéneo"],
        "estátic": ["estátic*"],
        "imóvel": ["imóvel"],
        "immob": ["immob*"],
    }

    def test_forma_ascii_cai_na_etiqueta_acentuada(self):
        self.assertEqual(
            tlex.canonical_de_forma("estaticas", self.campo), "estátic")
        self.assertEqual(
            tlex.canonical_de_forma("estaticidade", self.campo), "estátic")
        self.assertEqual(
            tlex.canonical_de_forma("estatico", self.campo), "estátic")
        self.assertEqual(
            tlex.canonical_de_forma("imovel", self.campo), "imóvel")

    def test_homoge_ganha_a_homogeneo(self):
        self.assertEqual(
            tlex.canonical_de_forma("homogeneo", self.campo), "homoge")
        self.assertEqual(
            tlex.canonical_de_forma("homogeneous", self.campo), "homoge")

    def test_nao_funde_static_com_estatic(self):
        self.assertNotEqual(
            tlex.canonical_de_forma("static", self.campo), "estátic")


if __name__ == "__main__":
    unittest.main()
