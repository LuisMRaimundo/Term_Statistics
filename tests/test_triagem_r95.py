# -*- coding: utf-8 -*-
"""R95 triage: false friends, metatexto, domain default, checklist."""

from __future__ import annotations

import unittest

import pandas as pd

import textura_triagem as ttri


class TestFalsosAmigos(unittest.TestCase):
    def test_continuo_excluded(self):
        df = pd.DataFrame({
            "matched_form": ["continuo", "continuous", "uniform"],
            "nuclear": [True, True, True],
            "motivo_exclusao": ["", "", ""],
            "contexto": ["figured continuo", "continuous texture", "uniform texture"],
            "caminho_ficheiro": ["a.pdf", "a.pdf", "a.pdf"],
            "relacao_sintactica": ["incidental", "atributiva", "atributiva"],
        })
        out = ttri.aplicar_falsos_amigos(df)
        self.assertFalse(bool(out.loc[0, "nuclear"]))
        self.assertTrue(bool(out.loc[1, "nuclear"]))
        self.assertTrue(bool(out.loc[2, "nuclear"]))
        self.assertIn("nao_relacionado", str(out.loc[0, "motivo_exclusao"]))

    def test_continuum_excluded(self):
        self.assertEqual(
            ttri.motivo_falso_amigo("continuum"),
            "fora_da_classe_uniforme",
        )


class TestMetatexto(unittest.TestCase):
    def test_jstor_terms(self):
        self.assertTrue(
            ttri.e_metatexto("This content downloaded from JSTOR.org/terms")
        )

    def test_musical_ok(self):
        self.assertFalse(
            ttri.e_metatexto("a continuous texture of interlocking lines")
        )


class TestDominioOmissao(unittest.TestCase):
    def test_default_musicologia(self):
        df = pd.DataFrame({
            "matched_form": ["uniform"],
            "nuclear": [True],
            "motivo_exclusao": [""],
            "contexto": ["uniform texture throughout"],
            "caminho_ficheiro": ["obscure/path/book.pdf"],
            "relacao_sintactica": ["atributiva"],
        })
        out, _, _ = ttri.aplicar_triagem(
            df, regras_dominio=[], dominio_omissao="musicologia",
        )
        self.assertEqual(out.loc[0, "dominio"], "musicologia")
        self.assertTrue(bool(out.loc[0, "nuclear"]))


class TestChecklist(unittest.TestCase):
    def test_false_friend_warning(self):
        df = pd.DataFrame({
            "matched_form": ["continuo"],
            "nuclear": [True],
            "relacao_sintactica": ["atributiva"],
            "polaridade": ["+"],
            "eixo": ["uniformidade"],
            "dominio": ["musicologia"],
            "revisto_por_humano": ["LR"],
        })
        chk = ttri.checklist_revisao(df)
        self.assertTrue(chk["ok"])
        self.assertTrue(any("falsos amigos" in a for a in chk["avisos"]))

    def test_bad_nuclear_relation_error(self):
        df = pd.DataFrame({
            "matched_form": ["uniform"],
            "nuclear": [True],
            "relacao_sintactica": ["incidental"],
            "polaridade": ["+"],
            "eixo": ["uniformidade"],
            "dominio": ["musicologia"],
            "revisto_por_humano": ["LR"],
        })
        chk = ttri.checklist_revisao(df)
        self.assertFalse(chk["ok"])
        self.assertTrue(chk["erros"])


if __name__ == "__main__":
    unittest.main()
