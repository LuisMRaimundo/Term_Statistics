# -*- coding: utf-8 -*-
"""Default dominios TSV resolution via caminho_dominios_path."""

from __future__ import annotations

import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class TestDominiosDefaultPath(unittest.TestCase):
    def test_caminho_dominios_path_is_dados_lexicos(self):
        """Canonical path rules live under dados/lexicos/."""
        from textura.lexico import caminho_dominios_path

        cand = caminho_dominios_path()
        self.assertEqual(cand, ROOT / "dados" / "lexicos" / "dominios_path.tsv")
        self.assertTrue(cand.is_file(), f"expected {cand}")
        self.assertNotEqual(cand.parent.name, "textura")

    def test_pipeline_default_uses_caminho_helper(self):
        import inspect

        import textura.pipeline as pipe

        src = inspect.getsource(pipe.main)
        self.assertIn("caminho_dominios_path", src)

    def test_default_branch_loads_when_dominios_omitted(self):
        import textura_triagem as ttri
        from textura.lexico import caminho_dominios_path

        regras = ttri.carregar_dominios(caminho_dominios_path())
        self.assertGreater(len(regras), 0)
        dom = ttri.classificar_dominio(
            r"E:\todos os textos\(2000)_X.pdf", regras)
        self.assertEqual(dom, "musicologia")

    def test_root_is_comment_stub_without_rules(self):
        """Shipped root dominios.tsv must not silently shadow the canonical file."""
        from textura.lexico import _linhas_regras_dominio

        root = ROOT / "dominios.tsv"
        self.assertTrue(root.is_file())
        self.assertEqual(
            _linhas_regras_dominio(root),
            frozenset(),
            "root dominios.tsv must be comments-only; put rules in "
            "dados/lexicos/dominios_path.tsv",
        )


if __name__ == "__main__":
    unittest.main()
