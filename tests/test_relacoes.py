# -*- coding: utf-8 -*-
"""Testes T1/T2 — emparelhamento e tokenização com hífen."""

from __future__ import annotations

import unittest

import pytest

import textura_near as tn


class TestTokenizacaoHifen(unittest.TestCase):
    def test_di_uniform_um_token(self):
        toks = tn.tokeniza(tn.normaliza("di-uniform texture spaces"))
        formas = [w for w, _ in toks]
        self.assertIn("di-uniform", formas)
        self.assertNotIn("di", formas)
        self.assertEqual(formas.count("uniform"), 0)

    def test_uniform_estrela_nao_casa_di_uniform(self):
        campo = tn.compila_campo({"uniform": ["uniform*"]})
        toks = tn.tokeniza(tn.normaliza(
            "di-uniform texture spaces are defined as follows"))
        nos = tn.NOS["en"]
        ach, _, _ = tn.emparelha_contexto(toks, nos, campo, near=4,
                                         limites=[], mesma_frase=False)
        self.assertEqual(ach, [], "uniform* não deve casar di-uniform")

    def test_estrela_uniform_casa_di_uniform(self):
        campo = tn.compila_campo({"uniform": ["*uniform"]})
        toks = tn.tokeniza(tn.normaliza("di-uniform texture spaces"))
        nos = tn.NOS["en"]
        ach, _, _ = tn.emparelha_contexto(toks, nos, campo, near=4,
                                         limites=[], mesma_frase=False)
        self.assertTrue(ach)
        self.assertTrue(ach[0]["forma_em_composto"])

    def test_invaria_nao_casa_rotation_invariant(self):
        campo = tn.compila_campo({"invariable": ["invaria*"]})
        toks = tn.tokeniza(tn.normaliza(
            "multiresolution gray-scale and rotation-invariant "
            "texture classification"))
        nos = tn.NOS["en"]
        ach, _, _ = tn.emparelha_contexto(toks, nos, campo, near=8,
                                         limites=[], mesma_frase=False)
        self.assertEqual(ach, [], "invaria* não deve casar rotation-invariant")


@pytest.mark.core
class TestDependencias(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            import spacy
            cls.nlp = spacy.load("en_core_web_sm",
                                 disable=["ner", "lemmatizer", "textcat"])
        except Exception as exc:  # noqa: BLE001
            raise unittest.SkipTest(f"spaCy indisponível: {exc}") from exc

    def _classifica(self, texto: str, forma_no: str, forma_termo: str):
        ctx = tn.normaliza(texto)
        doc = self.nlp(ctx)
        toks = tn.tokeniza(ctx)
        i_no = next(i for i, (w, _) in enumerate(toks) if w == forma_no)
        i_te = next(i for i, (w, _) in enumerate(toks) if w == forma_termo)
        return tn.relacao_dependencia(doc, toks[i_no][1], toks[i_te][1])

    def test_atributiva_uniform_texture(self):
        r = self._classifica(
            "long sections are marked conspicuously by uniform texture",
            "texture", "uniform")
        self.assertEqual(r["relacao_sintactica"], "atributiva")
        self.assertTrue(r["nuclear"])

    def test_predicativa_not_uniform(self):
        r = self._classifica(
            "the texture is not uniform, there being a number of additions",
            "texture", "uniform")
        self.assertEqual(r["relacao_sintactica"], "predicativa")
        self.assertTrue(r["nuclear"])

    def test_nominal_composto_textural_diversity(self):
        r = self._classifica(
            "one is struck by the textural diversity",
            "textural", "diversity")
        self.assertEqual(r["relacao_sintactica"], "nominal_composto")
        self.assertEqual(r["orientacao"], "no_sobre_termo")
        self.assertTrue(r["nuclear"])

    def test_nominal_genitiva_uniformity_of_texture(self):
        r = self._classifica(
            "a general tendency toward uniformity of texture",
            "texture", "uniformity")
        self.assertEqual(r["relacao_sintactica"], "nominal_genitiva")
        self.assertTrue(r["nuclear"])

    def test_adverbial_texturally_uniform(self):
        r = self._classifica(
            "the genre is almost completely texturally uniform",
            "texturally", "uniform")
        self.assertEqual(r["relacao_sintactica"], "adverbial")
        self.assertTrue(r["nuclear"])

    def test_incidental_static_block(self):
        r = self._classifica(
            "the full strings enter underneath the penultimate brass texture "
            "with a static harmonic block",
            "texture", "static")
        self.assertEqual(r["relacao_sintactica"], "incidental")
        self.assertFalse(r["nuclear"])
        self.assertIn(r["governante"].lower(), {"block", "harmonic"})

    def test_negacao_no_opus(self):
        ctx = "piece for orchestra no. 1 (1961), consist of static textures"
        toks = tn.tokeniza(tn.normaliza(ctx))
        i_te = next(i for i, (w, _) in enumerate(toks) if w == "static")
        neg, grad, mod = tn.anota_polaridade_linear(toks, i_te, tn.normaliza(ctx))
        self.assertFalse(neg)


class TestRegressoesR1R8(unittest.TestCase):
    def test_textur_estrela_igual_uniform_vira_campo(self):
        import tempfile
        from pathlib import Path
        import textura_lexico as tlex
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "c.txt"
            p.write_text("textur* = uniform*\n", encoding="utf-8")
            campo = tlex.carregar_campo_termos(
                p, campo_ref={"uniform": ["uniform*"]})
        self.assertIn("uniform", campo)
        self.assertNotIn("textur*", campo)
        tlex.assert_output_sem_no(["uniform"])
        with self.assertRaises(SystemExit):
            tlex.assert_output_sem_no(["textur*"])

    def test_canonical_polaridade_eixo(self):
        import textura_lexico as tlex
        campo = {"uniform": ["uniform*"]}
        tlex.registar_campo(campo)
        can = tlex.canonical_de_forma("uniformity", campo)
        self.assertEqual(can, "uniform")
        self.assertEqual(tlex.polaridade(can), "estabilidade")
        self.assertEqual(tlex.eixo_semantico(can), "homogeneidade_sincronica")

    def test_doc_id_estavel_por_ficheiro(self):
        import textura_lexico as tlex
        a = tlex.doc_id_de_caminho(r"E:\corpus\Levy_1982.pdf")
        b = tlex.doc_id_de_caminho(r"E:\outro\Levy_1982.pdf")
        c = tlex.doc_id_de_caminho(r"E:\corpus\Xenakis.pdf")
        self.assertEqual(a, b)
        self.assertNotEqual(a, c)

    def test_parece_caminho_ficheiro(self):
        import textura_lexico as tlex
        self.assertTrue(tlex.parece_caminho_ficheiro(
            r"E:\textos\artigo.pdf"))
        self.assertFalse(tlex.parece_caminho_ficheiro(
            r"E:\todos os textos"))

    def test_metatexto_example_sem_numero_nao_exclui(self):
        import textura_triagem as ttri
        self.assertFalse(ttri.e_metatexto(
            "An example is a slow movement of a piano sonata by Mozart. "
            "Although its textures are more uniform"))
        self.assertTrue(ttri.e_metatexto("Example 6a. Uniform texture."))
        self.assertTrue(ttri.e_metatexto(
            "this content downloaded from 132.174.255.1 on JSTOR"))

    def test_modalizado_can(self):
        toks = tn.tokeniza(tn.normaliza(
            "a particular static texture can be constructed by "
            "repetitive or random placement"))
        i = next(i for i, (w, _) in enumerate(toks) if w == "static")
        neg, grad, mod = tn.anota_polaridade_linear(toks, i)
        self.assertTrue(mod)

    @pytest.mark.core
    def test_governante_nao_e_matched(self):
        try:
            import spacy
            nlp = spacy.load("en_core_web_sm",
                             disable=["ner", "lemmatizer", "textcat"])
        except Exception as exc:  # noqa: BLE001
            raise unittest.SkipTest(str(exc)) from exc
        ctx = tn.normaliza("one is struck by the textural diversity")
        doc = nlp(ctx)
        toks = tn.tokeniza(ctx)
        i_no = next(i for i, (w, _) in enumerate(toks) if w == "textural")
        i_te = next(i for i, (w, _) in enumerate(toks) if w == "diversity")
        r = tn.relacao_dependencia(doc, toks[i_no][1], toks[i_te][1])
        self.assertNotEqual(r["governante"].lower(), "diversity")


class TestPrompt3(unittest.TestCase):
    def test_static_estrela_nao_casa_static_chordal(self):
        campo = tn.compila_campo({"static": ["static*"]})
        toks = tn.tokeniza(tn.normaliza(
            "the static-chordal texture remains"))
        ach, _, _ = tn.emparelha_contexto(
            toks, tn.NOS["en"], campo, near=4, limites=[], mesma_frase=False)
        self.assertEqual(ach, [], "static* nao deve casar static-chordal")

    def test_logdice_denominador_janelas_no(self):
        import textura_analise as ta
        # O11=112, O12=117, n_janelas=58013 -> ~5.94
        val = ta.logdice(112, 117, 58013)
        self.assertGreater(val, 5.0)
        self.assertLess(val, 7.0)
        # denominador errado 450 daria ~12.4
        errado = ta.logdice(112, 117, 450)
        self.assertGreater(errado, 11.0)

    def test_fase2_recusa_sem_instrucoes(self):
        import tempfile
        from pathlib import Path
        import textura_analise as ta
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "x.xlsx"
            pd = __import__("pandas")
            with pd.ExcelWriter(p) as xw:
                pd.DataFrame({"a": [1]}).to_excel(xw, sheet_name="8_Concordancia",
                                                  index=False)
            with self.assertRaises(SystemExit):
                ta.validar_fase1(p)

    def test_colinearidade_bloqueia_eixo(self):
        import pandas as pd
        import textura_analise as ta
        df = pd.DataFrame({
            "canonical_term": ["static"] * 5 + ["uniform"] * 5 + ["varied"] * 5,
            "eixo": (["invariancia_diacronica"] * 5
                     + ["homogeneidade_sincronica"] * 10),
        })
        self.assertTrue(ta.colinear_deterministica(
            df, "canonical_term", "eixo"))

    def test_folha_declara_N(self):
        import textura_analise as ta
        meta = ta._meta_linha("linhas nucleares (apos revisao manual)", 275)
        self.assertEqual(int(meta.loc[0, "N"]), 275)
        self.assertIn("unidade", meta.columns)


class TestEmparelhamento(unittest.TestCase):
    def test_static_texture_par_mais_proximo(self):
        ctx = ("static texture (mm. 434-440) abruptly negates "
               "the atonal contrapuntal texture")
        campo = tn.compila_campo({"static": ["static"]})
        toks = tn.tokeniza(tn.normaliza(ctx))
        ach, _, n_nos = tn.emparelha_contexto(
            toks, tn.NOS["en"], campo, near=8,
            limites=tn.fronteiras_frase(tn.normaliza(ctx)),
            mesma_frase=True)
        self.assertEqual(len(ach), 1)
        self.assertEqual(ach[0]["distancia"], 1)
        self.assertEqual(ach[0]["lado"], "esq")
        self.assertEqual(ach[0]["no"], "texture")
        self.assertGreaterEqual(n_nos, 2)

    def test_uniform_texture_adjacente(self):
        ctx = "long sections are marked conspicuously by uniform texture"
        campo = tn.compila_campo({"uniform": ["uniform*"]})
        toks = tn.tokeniza(tn.normaliza(ctx))
        ach, _, _ = tn.emparelha_contexto(
            toks, tn.NOS["en"], campo, near=4,
            limites=[], mesma_frase=False)
        self.assertEqual(len(ach), 1)
        self.assertEqual(ach[0]["distancia"], 1)

    def test_invariante_distancia_lado(self):
        ctx = "long sections are marked conspicuously by uniform texture"
        campo = tn.compila_campo({"uniform": ["uniform*"]})
        toks = tn.tokeniza(tn.normaliza(ctx))
        ach, _, _ = tn.emparelha_contexto(
            toks, tn.NOS["en"], campo, near=4,
            limites=[], mesma_frase=False)
        self.assertTrue(ach)
        rec = tn.recalcular_distancia_lado(
            ctx, ach[0]["no"], ach[0]["matched_form"],
            mesma_frase=False, nos_validos=tn.NOS["en"])
        self.assertIsNotNone(rec)
        self.assertEqual(rec["distancia"], ach[0]["distancia"])
        self.assertEqual(rec["lado"], ach[0]["lado"])


if __name__ == "__main__":
    unittest.main()
