# -*- coding: utf-8 -*-
"""Phase 2: single-source lexicons, taxonomy consistency, loud loader failures."""

from __future__ import annotations

import contextlib
import io
import os
import shutil
import tempfile
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

CTX_GEO = (
    "replacement of bioclasts by a mixture of dolomite and "
    "homogeneous textures in breccia facies of the Waulsortian"
)


class TestLexicosFonteUnica(unittest.TestCase):
    def tearDown(self):
        from textura import lexico as lx

        lx.definir_dir_lexicos_para_teste(None)

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

    def test_janela_labels_subset_of_merged_taxonomy(self):
        """Every dominio_janela label must exist in the merged taxonomy.

        Prevents dominio_janela:geologia naming a domain path-triage
        cannot recognise after a lone TSV edit.
        """
        from textura import lexico as lx

        tax = lx.carregar_dominio_taxonomia()
        missing = sorted(set(lx.carregar_dominio_janela()) - set(tax))
        self.assertEqual(
            missing,
            [],
            "dominio_janela labels absent from dominio_taxonomia.tsv: "
            + ", ".join(missing),
        )

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
                    offenders.append(
                        f"{path.relative_to(ROOT)}:{i}:{line.strip()}"
                    )
        self.assertEqual(
            offenders, [],
            "literal lexicon redefinitions:\n" + "\n".join(offenders),
        )

    def test_dominios_precedence_root_only_wins_with_stderr_aviso(self):
        from textura import lexico as lx

        with tempfile.TemporaryDirectory() as tmp:
            td = Path(tmp)
            raiz = td / "dominios.tsv"
            canon = td / "dominios_path.tsv"
            raiz.write_text(
                "# custom\n(?i)meu_corpus\tmusicologia\n", encoding="utf-8"
            )
            buf = io.StringIO()
            with contextlib.redirect_stderr(buf):
                path = lx.caminho_dominios_path(
                    raiz=raiz, canon=canon, avisar=True
                )
            err = buf.getvalue()
            self.assertEqual(path, raiz)
            self.assertIn("AVISO:", err)
            self.assertIn("legado", err)
            self.assertIn("dominios_path.tsv", err)

    def test_dominios_precedence_conflict_raises(self):
        from textura import lexico as lx

        with tempfile.TemporaryDirectory() as tmp:
            td = Path(tmp)
            raiz = td / "dominios.tsv"
            canon = td / "dominios_path.tsv"
            raiz.write_text("(?i)custom\tmusicologia\n", encoding="utf-8")
            canon.write_text("(?i)other\tmir_visao\n", encoding="utf-8")
            with self.assertRaises(lx.LexicoError) as ctx:
                lx.caminho_dominios_path(raiz=raiz, canon=canon, avisar=False)
            msg = str(ctx.exception)
            self.assertIn(str(raiz), msg)
            self.assertIn(str(canon), msg)

    def test_dominios_identical_prefers_canonical(self):
        from textura import lexico as lx

        with tempfile.TemporaryDirectory() as tmp:
            td = Path(tmp)
            raiz = td / "dominios.tsv"
            canon = td / "dominios_path.tsv"
            body = "(?i)todos os textos\tmusicologia\n"
            raiz.write_text(body, encoding="utf-8")
            canon.write_text(body, encoding="utf-8")
            buf = io.StringIO()
            with contextlib.redirect_stderr(buf):
                path = lx.caminho_dominios_path(
                    raiz=raiz, canon=canon, avisar=True
                )
            err = buf.getvalue()
            self.assertEqual(path, canon)
            self.assertIn("AVISO:", err)
            self.assertIn("duplicado obsoleto", err)

    def test_dominios_aviso_visible_under_default_warning_filters(self):
        """CLI users must see the notice without enabling DeprecationWarning."""
        import subprocess
        import sys

        script = (
            "import warnings\n"
            # Default filters (no pytest 'always' for DeprecationWarning)
            "warnings.resetwarnings()\n"
            "from pathlib import Path\n"
            "import tempfile\n"
            "from textura.lexico import caminho_dominios_path\n"
            "td = Path(tempfile.mkdtemp())\n"
            "raiz = td / 'dominios.tsv'\n"
            "canon = td / 'dominios_path.tsv'\n"
            "raiz.write_text('(?i)x\\tmusicologia\\n', encoding='utf-8')\n"
            "caminho_dominios_path(raiz=raiz, canon=canon, avisar=True)\n"
        )
        env = os.environ.copy()
        env["PYTHONPATH"] = str(ROOT)
        proc = subprocess.run(
            [sys.executable, "-c", script],
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
        )
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("AVISO:", proc.stderr)
        self.assertIn("dominios_path.tsv", proc.stderr)

    def test_malformed_tsv_raises_naming_file(self):
        from textura import lexico as lx

        with tempfile.TemporaryDirectory() as tmp:
            td = Path(tmp)
            shutil.copytree(DIR, td / "lexicos")
            alvo = td / "lexicos" / "negacao.tsv"
            alvo.write_text(
                "# broken\ntoken\nnot\twith\ttabs\n", encoding="utf-8"
            )
            lx.definir_dir_lexicos_para_teste(td / "lexicos")
            with self.assertRaises(lx.LexicoError) as ctx:
                lx._load_frozenset("negacao.tsv")
            self.assertIn("negacao.tsv", str(ctx.exception))

    def test_empty_tsv_raises_naming_file(self):
        from textura import lexico as lx

        with tempfile.TemporaryDirectory() as tmp:
            td = Path(tmp)
            shutil.copytree(DIR, td / "lexicos")
            alvo = td / "lexicos" / "copulas.tsv"
            alvo.write_text("# only comments\n", encoding="utf-8")
            lx.definir_dir_lexicos_para_teste(td / "lexicos")
            with self.assertRaises(lx.LexicoError) as ctx:
                lx._load_frozenset("copulas.tsv")
            self.assertIn("copulas.tsv", str(ctx.exception))

    def test_loader_anchored_to_project_root_not_cwd(self):
        from textura import lexico as lx

        here = Path.cwd()
        with tempfile.TemporaryDirectory() as tmp:
            os.chdir(tmp)
            try:
                self.assertEqual(lx.dir_lexicos(), DIR)
                self.assertEqual(lx.raiz_projecto(), ROOT)
                lx._limpar_caches()
                nos = lx.carregar_nos()
                self.assertIn("en", nos)
                self.assertEqual(lx.dominio_janela(CTX_GEO), "geologia")
            finally:
                os.chdir(here)

    def test_live_tsv_perturbation_changes_dominio_janela(self):
        """Removing geologia cues from a TSV copy must clear classification."""
        from textura import lexico as lx

        self.assertEqual(lx.dominio_janela(CTX_GEO), "geologia")
        with tempfile.TemporaryDirectory() as tmp:
            td = Path(tmp)
            shutil.copytree(DIR, td / "lexicos")
            janela = td / "lexicos" / "dominio_janela.tsv"
            lines = [
                ln for ln in janela.read_text(encoding="utf-8").splitlines()
                if not ln.startswith("geologia\t")
            ]
            janela.write_text("\n".join(lines) + "\n", encoding="utf-8")
            lx.definir_dir_lexicos_para_teste(td / "lexicos")
            self.assertEqual(lx.dominio_janela(CTX_GEO), "")
            # Taxonomy invariant still holds on the perturbed copy
            tax = lx.carregar_dominio_taxonomia()
            for dom in lx.carregar_dominio_janela():
                self.assertIn(dom, tax)


class TestLexicoLiveOnMiniature(unittest.TestCase):
    """Golden-adjacent: TSV edit changes dominio_janela on matrix context."""

    def tearDown(self):
        from textura import lexico as lx

        lx.definir_dir_lexicos_para_teste(None)

    def test_matrix_row_follows_perturbed_lexicon(self):
        import pandas as pd
        from textura import lexico as lx

        df = pd.read_excel(
            ROOT / "tests" / "fixtures" / "matriz_miniatura.xlsx",
            header=None,
        )
        ctx = None
        for col in df.columns:
            for val in df[col].tolist():
                if not isinstance(val, str):
                    continue
                if "bioclast" in val.lower():
                    ctx = val
                    break
            if ctx:
                break
        self.assertIsNotNone(ctx)
        self.assertEqual(lx.dominio_janela(ctx), "geologia")

        with tempfile.TemporaryDirectory() as tmp:
            td = Path(tmp)
            shutil.copytree(DIR, td / "lexicos")
            janela = td / "lexicos" / "dominio_janela.tsv"
            lines = [
                ln for ln in janela.read_text(encoding="utf-8").splitlines()
                if not ln.startswith("geologia\t")
            ]
            janela.write_text("\n".join(lines) + "\n", encoding="utf-8")
            lx.definir_dir_lexicos_para_teste(td / "lexicos")
            self.assertEqual(lx.dominio_janela(ctx), "")


if __name__ == "__main__":
    unittest.main()
