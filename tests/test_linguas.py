# -*- coding: utf-8 -*-
"""Phase 3: language registry — run-level resolution, EN anchor intact."""

from __future__ import annotations

import unittest

from textura.lexico import COPULAS, NEGACAO
from textura.linguas import CODIGOS, obter, resolver_execucao
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

    def test_todas_keeps_en_model_and_preps(self):
        ex = resolver_execucao("todas", None)
        self.assertEqual(ex.cfg.codigo, "en")
        self.assertEqual(ex.modelo_spacy, "en_core_web_sm")
        self.assertIn("sem detecção de língua por linha", ex.aviso)
        self.assertIn(
            "janelas não-EN classificadas com modelo/preposições EN",
            ex.aviso,
        )

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


if __name__ == "__main__":
    unittest.main()
