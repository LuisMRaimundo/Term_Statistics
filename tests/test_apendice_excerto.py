# -*- coding: utf-8 -*-
"""Tipografia e limpeza de excertos do apêndice DOCX."""

from __future__ import annotations

import unittest

import textura_apendice as ta


class TestFormatarExcerto(unittest.TestCase):
    def test_aspas_rectas_e_reticencias(self):
        out = ta.formatar_excerto(
            "to schoenberg, they are the one body of sound that represents "
            "a basic textural constant, in chamber music, orchestral music, "
            "or the concerto, in ora",
            matched_form="constant",
            no="textural",
        )
        self.assertTrue(out.startswith('"… '))
        self.assertTrue(out.endswith('"'))
        self.assertNotIn("«", out)
        self.assertNotIn("»", out)
        self.assertIn("constant", out)
        self.assertNotIn("in ora", out)
        self.assertIn("concerto", out)

    def test_nunca_corta_termo_de_pesquisa_na_cauda(self):
        # «even» no fim — termo protegido, não é coto a remover
        out = ta.formatar_excerto(
            "he had already devised more even",
            matched_form="even",
            no="textures",
        )
        self.assertIn("even", out.casefold())

    def test_nunca_corta_termo_apos_virgula(self):
        ctx = ("the passage shows a remarkably homogeneous, "
               "almost static field")
        # cauda «almost static field» tem 3 palavras — cortaria sem protecção
        # se homogeneous estivesse só na cauda... aqui homogeneous está antes
        out = ta.formatar_excerto(
            "layers remain continuous, in ora",
            matched_form="continuous",
            no="texture",
        )
        self.assertIn("continuous", out)
        self.assertNotIn("in ora", out)

    def test_protecao_se_termo_esta_na_cauda_curta(self):
        # se o termo está na cauda, NÃO cortar a cauda
        out = ta.formatar_excerto(
            "the score presents a dense, uniform",
            matched_form="uniform",
            no="texture",
        )
        self.assertIn("uniform", out)

    def test_desembrulha_guillemets_antigos(self):
        out = ta.formatar_excerto(
            "«… a homogeneous texture throughout …»",
            matched_form="homogeneous",
            no="texture",
        )
        self.assertTrue(out.startswith('"… '))
        self.assertIn("homogeneous", out)
        self.assertNotIn("«", out)

    def test_remove_particulas_finais_of_the_in(self):
        out = ta.formatar_excerto(
            "granular synthesis. furthermore, the use of",
            matched_form="homogeneous",
            no="textures",
        )
        self.assertNotRegex(out, r"\bof…\"$")
        self.assertNotRegex(out, r"\bthe use of")
        self.assertIn("furthermore", out)

    def test_nao_remove_even_mesmo_sendo_curto(self):
        out = ta.formatar_excerto(
            "he had already devised more even",
            matched_form="even",
            no="textures",
        )
        self.assertIn("even", out.casefold())

    def test_pagina_segue_o_excerto_nao_a_fonte(self):
        ex = ta.formatar_excerto(
            "a basic textural constant in chamber music",
            matched_form="constant",
            no="textural",
            pagina_ref=ta.PaginaRef(impressa="45"),
        )
        self.assertTrue(ex.startswith('"… '))
        self.assertTrue(ex.endswith("(p. 45)"))
        fo = ta.resolver_fonte(
            __import__("pandas").Series({
                "caminho_ficheiro": r"E:\corpus\book.pdf",
            }),
            {},
            pagina_ref=ta.PaginaRef(impressa="45"),
        )
        self.assertNotIn("p. 45", fo)
        self.assertNotIn("PDF p.", fo)


if __name__ == "__main__":
    unittest.main()
