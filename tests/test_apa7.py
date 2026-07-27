# -*- coding: utf-8 -*-
"""APA 7 formatters and DOI helpers (offline — no network)."""

from __future__ import annotations

import unittest

import textura_apa7 as apa
import textura_apendice as ta


class TestDoi(unittest.TestCase):
    def test_extract_bare_and_url(self):
        self.assertEqual(
            apa.extrair_doi("10.1234/abc.def"),
            "10.1234/abc.def",
        )
        self.assertEqual(
            apa.extrair_doi("https://doi.org/10.1234/abc.def"),
            "10.1234/abc.def",
        )
        self.assertEqual(
            apa.extrair_doi("https://dx.doi.org/10.1037/a0028240"),
            "10.1037/a0028240",
        )


class TestFormatApa7(unittest.TestCase):
    def test_journal_article(self):
        m = apa.Meta(
            authors=["Smith, J. A.", "Doe, R."],
            year="2019",
            title="Texture and form in chamber music",
            container="Journal of Music Theory",
            volume="63",
            issue="2",
            pages="123-145",
            doi="10.1234/jmt.2019",
            tipo="journal-article",
            source="crossref",
        )
        ref = apa.format_apa7(m)
        self.assertIn("Smith, J. A., & Doe, R. (2019).", ref)
        self.assertIn("Texture and form in chamber music.", ref)
        self.assertIn("*Journal of Music Theory*", ref)
        self.assertIn("*63*", ref)
        self.assertIn("(2)", ref)
        self.assertIn("123–145", ref)
        self.assertTrue(ref.endswith("https://doi.org/10.1234/jmt.2019"))
        self.assertFalse(ref.endswith("."))

    def test_book(self):
        m = apa.Meta(
            authors=["Berry, W."],
            year="1987",
            title="Structural functions in music",
            publisher="Dover",
            tipo="book",
            source="crossref",
        )
        ref = apa.format_apa7(m)
        self.assertIn("Berry, W. (1987).", ref)
        self.assertIn("*Structural functions in music*.", ref)
        self.assertIn("Dover.", ref)

    def test_sentence_case(self):
        self.assertEqual(
            apa.sentence_case("Texture And Form: A Study"),
            "Texture and form: A study",
        )


class TestResolverFonteNoPeriodAfterDoi(unittest.TestCase):
    def test_does_not_append_period_after_doi(self):
        import pandas as pd

        row = pd.Series({
            "fonte_apa": (
                "Smith, J. (2019). Title. *Journal*, *1*(2), 3–4. "
                "https://doi.org/10.1234/abc"
            ),
            "caminho_ficheiro": "x.pdf",
        })
        out = ta.resolver_fonte(row, {})
        self.assertTrue(out.endswith("https://doi.org/10.1234/abc"))
        self.assertFalse(out.endswith("abc."))


class TestFilenameFallback(unittest.TestCase):
    def test_parses_author_year_title(self):
        m = apa.fallback_from_filename("Smith_2019_Texture_and_form.pdf")
        self.assertEqual(m.year, "2019")
        self.assertTrue(m.title.lower().startswith("texture"))
        self.assertEqual(m.source, "filename")


if __name__ == "__main__":
    unittest.main()
