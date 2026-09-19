# -*- coding: utf-8 -*-
"""Phase 3: language registry — run-level resolution, EN anchor intact."""

from __future__ import annotations

import unittest

from textura.lexico import COPULAS, NEGACAO
from textura.linguas import (
    CODIGOS, LINGUA_OMISSAO, lingua_do_no, obter, resolver_execucao,
)
from textura.pipeline import NEAR_OMISSAO
from textura.relacoes import _PREPS_GENITIVO_OMISSAO


class TestRegistoLinguas(unittest.TestCase):
    def test_codigos_esperados(self):
        self.assertEqual(set(CODIGOS), {"en", "pt", "fr", "de"})

    def test_en_anchor_matches_pre_phase3_constants(self):
        en = obter("en")
        self.assertEqual(en.modelo_spacy, "en_core_web_sm")
        self.assertEqual(en.status, "validado")
        self.assertEqual(en.preps_genitivo, _PREPS_GENITIVO_OMISSAO)
        self.assertEqual(en.copulas, frozenset(COPULAS))
        self.assertEqual(en.negadores, frozenset(NEGACAO))
        self.assertEqual(en.preps_associativa, frozenset({"with"}))

    def test_fr_pt_validado_de_nao(self):
        self.assertEqual(obter("fr").status, "validado")
        self.assertEqual(obter("pt").status, "validado")
        self.assertEqual(obter("de").status, "nao_validado")

    def test_omissao_e_todas(self):
        self.assertEqual(LINGUA_OMISSAO, "todas")
        self.assertEqual(NEAR_OMISSAO, 8)
        ex = resolver_execucao("", None)
        self.assertEqual(ex.lingua_cli, "todas")
        self.assertEqual(ex.cfg.codigo, "en")

    def test_todas_keeps_en_model_and_preps(self):
        ex = resolver_execucao("todas", None)
        self.assertEqual(ex.cfg.codigo, "en")
        self.assertEqual(ex.modelo_spacy, "en_core_web_sm")
        self.assertIn("spaCy despachado", ex.aviso)
        self.assertIn("textura*", ex.aviso)

    def test_pt_resolves_pt_model(self):
        ex = resolver_execucao("pt", None)
        self.assertEqual(ex.modelo_spacy, "pt_core_news_sm")
        self.assertIn("de", ex.cfg.preps_genitivo)
        self.assertIn("com", ex.cfg.preps_associativa)

    def test_modelo_cli_override(self):
        ex = resolver_execucao("en", "en_core_web_md")
        self.assertEqual(ex.modelo_spacy, "en_core_web_md")

    def test_nao_validado_surfaced_in_aviso(self):
        ex = resolver_execucao("de", None)
        self.assertIn("não validado", ex.aviso)

    def test_nos_fr_registered(self):
        self.assertIn("texture", obter("fr").nos)

    def test_lingua_do_no_exclusivos_e_partilhados(self):
        self.assertEqual(lingua_do_no("textura"), "pt")
        self.assertEqual(lingua_do_no("texturas"), "pt")
        self.assertEqual(lingua_do_no("texturais"), "pt")
        self.assertEqual(lingua_do_no("textured"), "en")
        self.assertEqual(lingua_do_no("texturally"), "en")
        self.assertEqual(lingua_do_no("texture"), "en")
        self.assertEqual(lingua_do_no("textures"), "en")
        self.assertEqual(lingua_do_no("textural"), "en")
        self.assertEqual(lingua_do_no("texturen"), "de")


if __name__ == "__main__":
    unittest.main()
