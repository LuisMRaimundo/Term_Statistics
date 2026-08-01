# -*- coding: utf-8 -*-
"""Default dominios TSV must resolve via caminho_dominios_path (not textura/)."""

from __future__ import annotations

import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class TestDominiosDefaultPath(unittest.TestCase):
    def test_caminho_dominios_path_is_dados_lexicos(self):
        """Phase 2: canonical path rules live under dados/lexicos/."""
        from textura.lexico import caminho_dominios_path

        cand = caminho_dominios_path()
        self.assertEqual(cand, ROOT / "dados" / "lexicos" / "dominios_path.tsv")
        self.assertTrue(cand.is_file(), f"expected {cand}")
        # Must NOT point inside the package directory
        self.assertNotEqual(cand.parent.name, "textura")

    def test_pipeline_default_uses_caminho_helper(self):
        """pipeline.main must call caminho_dominios_path when --dominios omitted."""
        import inspect

        import textura.pipeline as pipe

        src = inspect.getsource(pipe.main)
        self.assertIn("caminho_dominios_path", src)

    def test_default_branch_loads_when_dominios_omitted(self):
        """Smoke: carregar_dominios on the resolved default path succeeds."""
        import textura_triagem as ttri
        from textura.lexico import caminho_dominios_path

        regras = ttri.carregar_dominios(caminho_dominios_path())
        self.assertGreater(len(regras), 0)
        dom = ttri.classificar_dominio(
            r"E:\todos os textos\(2000)_X.pdf", regras)
        self.assertEqual(dom, "musicologia")

    def test_root_stub_still_loads_if_present(self):
        """Compat: root dominios.tsv remains readable if scripts hard-code it."""
        import textura_triagem as ttri

        root = ROOT / "dominios.tsv"
        if not root.is_file():
            self.skipTest("root dominios.tsv absent")
        regras = ttri.carregar_dominios(root)
        self.assertGreater(len(regras), 0)


if __name__ == "__main__":
    unittest.main()
