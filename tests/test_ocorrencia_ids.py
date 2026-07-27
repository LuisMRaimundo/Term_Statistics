# -*- coding: utf-8 -*-
"""Testes schema ≥2 — ocorrência mestra vs hit NEAR vs janela."""

from __future__ import annotations

import unittest

import pandas as pd

import textura_near as tn


class TestIdentidadeOcorrencia(unittest.TestCase):
    def test_occurrence_id_estavel(self):
        self.assertEqual(
            tn.occurrence_id_de("fdc5c3371bafedcd", 18452),
            "fdc5c3371bafedcd::ROW_18452",
        )

    def test_hit_key_distingue_posicao(self):
        a = tn.hit_key_de("D::ROW_1", "continu", "continuous", 10, 2)
        b = tn.hit_key_de("D::ROW_1", "homoge", "homogeneous", 10, 4)
        c = tn.hit_key_de("D::ROW_1", "continu", "continuous", 10, 2)
        self.assertNotEqual(a, b)
        self.assertEqual(a, c)

    def test_match_ids_por_ocorrencia(self):
        res = pd.DataFrame([
            {
                "texture_occurrence_id": "d::ROW_10",
                "canonical_term": "homoge",
                "matched_form": "homogeneous",
                "off_no": 20, "off_termo": 10,
                "idx_termo": 5,
            },
            {
                "texture_occurrence_id": "d::ROW_10",
                "canonical_term": "continu",
                "matched_form": "continuous",
                "off_no": 20, "off_termo": 5,
                "idx_termo": 2,
            },
            {
                "texture_occurrence_id": "d::ROW_11",
                "canonical_term": "uniform",
                "matched_form": "uniform",
                "off_no": 8, "off_termo": 3,
                "idx_termo": 1,
            },
        ])
        out = tn.atribuir_match_ids(res)
        m10 = out.loc[out["texture_occurrence_id"] == "d::ROW_10"].sort_values(
            "match_id")
        self.assertEqual(list(m10["match_id"]), ["M001", "M002"])
        self.assertEqual(
            out.loc[out["texture_occurrence_id"] == "d::ROW_11", "match_id"]
            .iloc[0],
            "M001",
        )
        self.assertTrue(out["hit_key"].str.startswith("d::ROW_").all())


class TestFusaoJanelas(unittest.TestCase):
    def _base_row(self, **kw):
        base = {
            "source_matrix_row": 1,
            "texture_occurrence_id": "doc::ROW_1",
            "match_id": "M001",
            "hit_key": "x",
            "grupo_passagem_id": "",
            "candidato_duplicado": "",
            "canonical_term": "homoge",
            "matched_form": "homogeneous",
            "doc_id": "docA",
            "caminho_ficheiro": r"E:\t\book.pdf",
            "url": "",
            "no": "texture",
            "contexto": (
                "extent to which a homogeneous texture a texture in which "
                "all the parts share the same musical material and are"
            ),
            "motivo_exclusao": "",
            "nuclear": True,
            "n_janelas_fundidas": 1,
            "off_no": 20,
            "off_termo": 10,
        }
        base.update(kw)
        return base

    def test_mesmo_termo_janelas_deslocadas_funde(self):
        a = self._base_row(
            source_matrix_row=292,
            texture_occurrence_id="docA::ROW_292",
            contexto=(
                "aping of a modem language 201 80 extent to which a "
                "homogeneous texture a texture in which all the parts "
                "share the same musical material and are unif"
            ),
        )
        b = self._base_row(
            source_matrix_row=293,
            texture_occurrence_id="docA::ROW_293",
            contexto=(
                "modem language 201 80 extent to which a homogeneous "
                "texture a texture in which all the parts share the same "
                "musical material and are unified"
            ),
            nuclear=True,
        )
        out = tn.fundir_janelas_e_marcar_duplicados(pd.DataFrame([a, b]))
        nucs = out["nuclear"].astype(bool).tolist()
        self.assertEqual(sum(nucs), 1, "só um hit nuclear após fusão")
        self.assertTrue(
            (out["motivo_exclusao"] == "janela_sobreposta").any()
        )
        self.assertGreaterEqual(int(out.loc[out["nuclear"], "n_janelas_fundidas"].iloc[0]), 2)
        self.assertTrue(
            out["grupo_passagem_id"].astype(str).str.startswith("P").all()
        )

    def test_termos_diferentes_na_mesma_passagem_nao_exclui(self):
        ctx = (
            "a continuous and relatively homogeneous texture that remains "
            "stable throughout the entire movement without abrupt change"
        )
        a = self._base_row(
            source_matrix_row=100,
            texture_occurrence_id="docA::ROW_100",
            canonical_term="continu",
            matched_form="continuous",
            contexto=ctx,
        )
        b = self._base_row(
            source_matrix_row=101,
            texture_occurrence_id="docA::ROW_101",
            canonical_term="homoge",
            matched_form="homogeneous",
            contexto="and relatively " + ctx,
        )
        out = tn.fundir_janelas_e_marcar_duplicados(pd.DataFrame([a, b]))
        self.assertTrue(bool(out["nuclear"].all()),
                        "hits de termos distintos devem permanecer")
        self.assertTrue(
            out["candidato_duplicado"].astype(str).str.contains(
                "passagem_sobreposta").all()
        )

    def test_agregar_ocorrencias_dois_hits(self):
        res = pd.DataFrame([
            self._base_row(
                source_matrix_row=50,
                texture_occurrence_id="docA::ROW_50",
                match_id="M001",
                canonical_term="continu",
                matched_form="continuous",
                nuclear=True,
            ),
            self._base_row(
                source_matrix_row=50,
                texture_occurrence_id="docA::ROW_50",
                match_id="M002",
                canonical_term="homoge",
                matched_form="homogeneous",
                nuclear=True,
            ),
        ])
        res = tn.atribuir_match_ids(res)
        occ = tn.agregar_ocorrencias(res)
        self.assertEqual(len(occ), 1)
        self.assertEqual(int(occ.iloc[0]["n_matches"]), 2)
        self.assertIn("continu", occ.iloc[0]["matched_terms"])
        self.assertIn("homoge", occ.iloc[0]["matched_terms"])

    def test_mesma_ocorrencia_dois_hits_nao_marca_passagem(self):
        ctx = (
            "a continuous and relatively homogeneous texture remains "
            "stable throughout the entire movement without abrupt change"
        )
        res = pd.DataFrame([
            self._base_row(
                source_matrix_row=50,
                texture_occurrence_id="docA::ROW_50",
                canonical_term="continu",
                matched_form="continuous",
                contexto=ctx,
            ),
            self._base_row(
                source_matrix_row=50,
                texture_occurrence_id="docA::ROW_50",
                canonical_term="homoge",
                matched_form="homogeneous",
                contexto=ctx,
            ),
        ])
        out = tn.fundir_janelas_e_marcar_duplicados(res)
        self.assertTrue(bool(out["nuclear"].all()))
        self.assertTrue(
            (out["grupo_passagem_id"].astype(str).str.strip() == "").all()
        )


class TestDedupeAnalise(unittest.TestCase):
    def test_modos_ocorrencia(self):
        import textura_analise as ta
        df = pd.DataFrame([
            {"texture_occurrence_id": "A::ROW_1", "canonical_term": "uniform",
             "doc_id": "A", "contexto": "x"},
            {"texture_occurrence_id": "A::ROW_1", "canonical_term": "homoge",
             "doc_id": "A", "contexto": "x"},
            {"texture_occurrence_id": "A::ROW_2", "canonical_term": "uniform",
             "doc_id": "A", "contexto": "y"},
        ])
        o, m = ta.aplicar_desduplicacao(df, "ocorrencia", "doc_id")
        self.assertEqual(m, "ocorrencia")
        self.assertEqual(len(o), 2)
        ot, m2 = ta.aplicar_desduplicacao(df, "ocorrencia_termo", "doc_id")
        self.assertEqual(m2, "ocorrencia_termo")
        self.assertEqual(len(ot), 3)


if __name__ == "__main__":
    unittest.main()
