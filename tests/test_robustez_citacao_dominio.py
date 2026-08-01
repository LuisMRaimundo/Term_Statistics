# -*- coding: utf-8 -*-
"""Testes de robustez — citações entre documentos, coordenação heterogénea
e domínio por janela (patches 2026-08)."""

from __future__ import annotations

import unittest

import pandas as pd
import pytest

import textura_near as tn

try:
    import spacy
    _NLP = spacy.load("en_core_web_sm",
                      disable=["ner", "lemmatizer", "textcat"])
except Exception:                                    # pragma: no cover
    _NLP = None

requer_spacy = pytest.mark.skipif(_NLP is None,
                                  reason="modelo spaCy indisponível")


def _linha(i, doc, ctx, termo="combin", forma="combination", nuclear=True,
           caminho=r"E:\todos os textos\(2000)_Obra.pdf"):
    return {
        "source_matrix_row": i,
        "texture_occurrence_id": f"{doc}::ROW_{i}",
        "doc_id": doc,
        "canonical_term": termo,
        "matched_form": forma,
        "contexto": ctx,
        "nuclear": nuclear,
        "motivo_exclusao": "",
        "caminho_ficheiro": caminho,
    }


class TestCitacaoEntreDocumentos(unittest.TestCase):
    """A detecção deixa de exigir igualdade exacta da janela inteira."""

    CTX_A = ("clock, which measures duration, with which to measure the "
             "degree of textural complexity irvine asks the question")
    # mesma passagem, fronteira de janela deslocada (variante de edição)
    CTX_B = ("ck, which measures duration, with which to measure the "
             "degree of textural complexity")

    def _res(self):
        return pd.DataFrame([
            _linha(1, "docA", self.CTX_A, termo="complex",
                   forma="complexity",
                   caminho=r"E:\todos os textos\(1960)_Medida.pdf"),
            _linha(2, "docB", self.CTX_B, termo="complex",
                   forma="complexity",
                   caminho=r"E:\todos os textos\Copia__ab12cd34.pdf"),
        ])

    def test_janela_deslocada_e_agrupada(self):
        out = tn.fundir_janelas_e_marcar_duplicados(self._res(), ngrama=8)
        tags = out["candidato_duplicado"].str.contains(
            "citacao_entre_doc_ids", na=False)
        self.assertTrue(tags.all(), "ambas as linhas devem ser etiquetadas")

    def test_demove_apenas_uma_copia(self):
        out = tn.fundir_janelas_e_marcar_duplicados(self._res(), ngrama=8)
        self.assertEqual(int(out["nuclear"].sum()), 1)
        demovida = out.loc[~out["nuclear"].astype(bool)]
        self.assertEqual(demovida["motivo_exclusao"].iloc[0],
                         "citacao_repetida")

    def test_sobrevivente_prefere_ficheiro_catalogado(self):
        out = tn.fundir_janelas_e_marcar_duplicados(self._res(), ngrama=8)
        viva = out.loc[out["nuclear"].astype(bool), "caminho_ficheiro"].iloc[0]
        self.assertIn("(1960)_", viva)

    def test_termos_distintos_nao_sao_demovidos(self):
        res = self._res()
        res.loc[1, "canonical_term"] = "blend"
        res.loc[1, "matched_form"] = "blend"
        out = tn.fundir_janelas_e_marcar_duplicados(res, ngrama=8)
        # passagem etiquetada, mas nenhum hit demovido (termos diferentes)
        self.assertEqual(int(out["nuclear"].sum()), 2)

    def test_documento_unico_nao_e_citacao(self):
        res = self._res()
        res["doc_id"] = "docA"
        res["texture_occurrence_id"] = ["docA::ROW_1", "docA::ROW_2"]
        out = tn.fundir_janelas_e_marcar_duplicados(res, ngrama=8)
        self.assertFalse(out["candidato_duplicado"].str.contains(
            "citacao_entre_doc_ids", na=False).any())


@requer_spacy
class TestCoordenacaoHeterogenea(unittest.TestCase):
    def _rel(self, frase, no, termo):
        doc = _NLP(frase)
        return tn.relacao_dependencia(doc, frase.index(no),
                                      frase.index(termo))

    def test_of_texture_and_dynamics_sinalizado(self):
        r = self._rel("the combination of texture and dynamics was clear",
                      "texture", "combination")
        self.assertTrue(r["nuclear"])
        self.assertIn("coordenacao_heterogenea", r["revisao_sugerida"])

    def test_of_textures_sem_aviso(self):
        r = self._rel("the combination of textures was clear",
                      "textures", "combination")
        self.assertTrue(r["nuclear"])
        self.assertNotIn("coordenacao_heterogenea",
                         r.get("revisao_sugerida") or "")

    def test_textures_combined_with_lyrics(self):
        r = self._rel("the guitar textures combined with the lyrics of "
                      "the singer", "textures", "combined")
        if r["nuclear"]:
            self.assertIn("associativa_com_nao_textural",
                          r["revisao_sugerida"])

    def test_combined_into_texture_sem_aviso_associativo(self):
        r = self._rel("the voices are combined into a homophonic texture",
                      "texture", "combined")
        self.assertNotIn("associativa_com_nao_textural",
                         r.get("revisao_sugerida") or "")


class TestDominioJanela(unittest.TestCase):
    def test_geologia(self):
        self.assertEqual(
            tn.dominio_janela("replacement of bioclasts by a mixture of "
                              "dolomite, fe textures in breccia facies"),
            "geologia")

    def test_musical_sem_dominio(self):
        self.assertEqual(
            tn.dominio_janela("a complex polyphonic texture of "
                              "superimposed rhythmic layers"), "")

    def test_haptica(self):
        self.assertEqual(
            tn.dominio_janela("libraries of heterogeneous textures for a "
                              "haptic understanding of materials"),
            "haptica_materiais")


if __name__ == "__main__":              # pragma: no cover
    unittest.main()
