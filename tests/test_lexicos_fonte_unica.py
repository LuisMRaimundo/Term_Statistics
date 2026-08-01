# -*- coding: utf-8 -*-
"""Phase 2: TSV lexicons load non-empty; taxonomy covers path + janela domains."""

from __future__ import annotations

import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DIR = ROOT / "dados" / "lexicos"

TSVS_OBRIGATORIOS = (
    "nos.tsv",
    "negacao.tsv",
    "graduacao.tsv",
    "modalidade.tsv",
    "copulas.tsv",
    "abreviaturas.tsv",
    "polo_estabilidade.tsv",
    "polo_variabilidade.tsv",
    "relacoes_nucleares.tsv",
    "relacoes_nao_nucleares.tsv",
    "dominio_janela.tsv",
    "dominio_taxonomia.tsv",
    "dominios_path.tsv",
    "falsos_amigos.tsv",
)


class TestLexicosFonteUnica(unittest.TestCase):
    def test_tsv_files_present_and_non_empty(self):
        for nome in TSVS_OBRIGATORIOS:
            path = DIR / nome
            self.assertTrue(path.is_file(), f"missing {path}")
            self.assertGreater(path.stat().st_size, 20, f"empty {path}")

    def test_loader_exports_non_empty(self):
        from textura import lexico as lx

        self.assertGreater(len(lx.NOS), 0)
        self.assertIn("en", lx.NOS)
        self.assertGreater(len(lx.NEGACAO), 5)
        self.assertGreater(len(lx.ABREVIATURAS), 10)
        self.assertGreater(len(lx.POLO_ESTABILIDADE), 5)
        self.assertGreater(len(lx.POLO_VARIABILIDADE), 3)
        self.assertGreater(len(lx.RELACOES_NUCLEARES), 3)
        self.assertGreater(len(lx.RELACOES_NAO_NUCLEARES), 2)
        self.assertGreater(len(lx.DOMINIO_JANELA_LEXICO), 3)
        self.assertGreater(len(lx.FALSOS_AMIGOS_FORMAS), 3)
        self.assertIn("musicologia", lx.DOMINIOS_VALIDOS)

    def test_taxonomia_covers_path_and_janela(self):
        from textura import lexico as lx

        tax = lx.carregar_dominio_taxonomia()
        for dom in lx.DOMINIOS_VALIDOS:
            self.assertIn(dom, tax)
            self.assertIn(tax[dom], {"path", "ambos"})
        for dom in lx.DOMINIO_JANELA_LEXICO:
            self.assertIn(dom, tax)
            self.assertIn(tax[dom], {"janela", "ambos"})

    def test_consumers_share_same_object_identity_for_frozensets(self):
        """Re-exports must not redefine literals — same frozenset objects."""
        from textura import config, lexico
        import textura_triagem as ttri
        import textura_query as tq

        self.assertIs(config.ABREVIATURAS, lexico.ABREVIATURAS)
        self.assertIs(tq.ABREVIATURAS, lexico.ABREVIATURAS)
        self.assertIs(config.RELACOES_NUCLEARES, lexico.RELACOES_NUCLEARES)
        self.assertIs(ttri.RELACOES_NUCLEARES, lexico.RELACOES_NUCLEARES)
        self.assertIs(ttri.DOMINIOS_VALIDOS, lexico.DOMINIOS_VALIDOS)

    def test_no_duplicate_literal_assignments_in_py_sources(self):
        """Grep guard: inventory lists must not be re-literalised outside lexico."""
        import re

        # Only brace-literal redefinitions (not set(tsv) / import aliases).
        ban = re.compile(
            r"^(NOS|NEGACAO|GRADUACAO|MODALIDADE|COPULAS|ABREVIATURAS|"
            r"POLO_ESTABILIDADE|POLO_VARIABILIDADE|RELACOES_NUCLEARES|"
            r"RELACOES_NAO_NUCLEARES|DOMINIO_JANELA_LEXICO|DOMAIN_LEXICON|"
            r"FALSOS_AMIGOS_FORMAS|DOMINIOS_VALIDOS)\s*=\s*\{"
        )
        allowed = {
            ROOT / "textura" / "lexico.py",
        }
        offenders = []
        for path in ROOT.rglob("*.py"):
            if "tests" in path.parts or ".venv" in path.parts:
                continue
            if path.resolve() in {p.resolve() for p in allowed}:
                continue
            text = path.read_text(encoding="utf-8")
            for i, line in enumerate(text.splitlines(), 1):
                if ban.match(line.strip()):
                    offenders.append(f"{path.relative_to(ROOT)}:{i}:{line.strip()}")
        self.assertEqual(offenders, [], "literal lexicon redefinitions:\n" + "\n".join(offenders))


if __name__ == "__main__":
    unittest.main()
