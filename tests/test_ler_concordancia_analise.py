# -*- coding: utf-8 -*-
"""Concordância em *_analise.xlsx (banner meta) deve ser legível."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pandas as pd

import textura_analise as ta


class LerConcordanciaAnaliseTests(unittest.TestCase):
    def test_banner_meta_header3(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "demo_analise.xlsx"
            meta = pd.DataFrame({"unidade": ["hits"], "N": [2]})
            body = pd.DataFrame({
                "canonical_term": ["complex", "mix"],
                "nuclear": [True, True],
                "relacao_sintactica": ["atributiva", "predicativa"],
            })
            with pd.ExcelWriter(path, engine="openpyxl") as xw:
                meta.to_excel(xw, sheet_name="8_Concordancia", index=False, startrow=0)
                body.to_excel(xw, sheet_name="8_Concordancia", index=False, startrow=3)
                pd.DataFrame({"chave": ["fase"], "valor": ["2"]}).to_excel(
                    xw, sheet_name="0_Instrucoes", index=False)
                pd.DataFrame({"indicador": ["x"], "valor": [1]}).to_excel(
                    xw, sheet_name="1_Resumo", index=False)
            conc = ta.ler_folha_concordancia(path)
            self.assertIn("relacao_sintactica", conc.columns)
            self.assertEqual(len(conc), 2)

    def test_sugerir_entrada(self):
        with tempfile.TemporaryDirectory() as td:
            rev = Path(td) / "Comp_near_revisto_v2.xlsx"
            ana = Path(td) / "Comp_near_revisto_v2_analise.xlsx"
            rev.write_bytes(b"x")
            ana.write_bytes(b"x")
            self.assertEqual(ta._sugerir_entrada_revisao(ana), str(rev))

    def test_sincronizar_hits(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "rev.xlsx"
            conc = pd.DataFrame({
                "hit_key": ["a", "b", "c"],
                "nuclear": [True, True, False],
                "relacao_sintactica": ["atributiva", "predicativa", "incidental"],
                "canonical_term": ["complex", "mix", "blend"],
            })
            hits_old = conc.copy()
            hits_old.loc[0, "nuclear"] = False  # desactualizado
            with pd.ExcelWriter(path, engine="openpyxl") as xw:
                conc.to_excel(xw, sheet_name="8_Concordancia", index=False)
                hits_old.to_excel(xw, sheet_name="8_Concordancia_Hits", index=False)
            out = ta.sincronizar_hits_com_concordancia(path)
            self.assertTrue(out["ok"])
            self.assertEqual(out["n_conc"], 2)
            self.assertEqual(out["n_hits_antes"], 1)
            hits = pd.read_excel(path, sheet_name="8_Concordancia_Hits")
            self.assertEqual(ta._nuclear_true_count(hits["nuclear"]), 2)


if __name__ == "__main__":
    unittest.main()
