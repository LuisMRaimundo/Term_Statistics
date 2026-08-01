# -*- coding: utf-8 -*-
"""Default dominios.tsv must resolve at the project root, not textura/."""

from __future__ import annotations

import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]


class TestDominiosDefaultPath(unittest.TestCase):
    def test_pipeline_default_cand_is_project_root(self):
        """Regression: after Phase 1, __file__ is textura/pipeline.py."""
        import textura.pipeline as pipe

        pipeline_file = Path(pipe.__file__).resolve()
        self.assertEqual(pipeline_file.parent.name, "textura")
        # Exact idiom used in pipeline.main for --dominios default
        cand = pipeline_file.parents[1] / "dominios.tsv"
        self.assertEqual(cand, ROOT / "dominios.tsv")
        self.assertTrue(
            cand.is_file(),
            f"expected committed dominios.tsv at project root: {cand}",
        )
        # Must NOT point inside the package
        self.assertFalse(
            (pipeline_file.parent / "dominios.tsv").is_file()
            and cand == pipeline_file.parent / "dominios.tsv"
        )

    def test_default_branch_loads_when_dominios_omitted(self):
        """Smoke: carregar_dominios on the resolved default path succeeds."""
        import textura_triagem as ttri

        regras = ttri.carregar_dominios(ROOT / "dominios.tsv")
        self.assertGreater(len(regras), 0)
        # musicologia rule for «todos os textos» is the default corpus cue
        dom = ttri.classificar_dominio(
            r"E:\todos os textos\(2000)_X.pdf", regras)
        self.assertEqual(dom, "musicologia")


if __name__ == "__main__":
    unittest.main()
