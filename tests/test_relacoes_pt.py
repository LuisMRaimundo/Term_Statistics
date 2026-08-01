# -*- coding: utf-8 -*-
"""Phase 3: PT dependency classification (skip if pt_core_news_sm absent).

PT sentences live here — not in matriz_miniatura — so the EN golden chain
stays byte-stable under ``--lingua todas``.
"""

from __future__ import annotations

import unittest

import pytest

from textura.linguas import obter
from textura.relacoes import relacao_dependencia
from textura.tokenizacao import normaliza, tokeniza

pt_cfg = obter("pt")


def _pt_nlp():
    try:
        import spacy
        return spacy.load(
            pt_cfg.modelo_spacy, disable=["ner", "lemmatizer", "textcat"]
        )
    except Exception as exc:  # noqa: BLE001
        raise unittest.SkipTest(
            f"modelo {pt_cfg.modelo_spacy} indisponível: {exc}"
        ) from exc


@pytest.mark.core
class TestDependenciasPT(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.nlp = _pt_nlp()

    def _classifica(self, texto: str, forma_no: str, forma_termo: str):
        ctx = normaliza(texto)
        doc = self.nlp(ctx)
        toks = tokeniza(ctx)
        i_no = next(i for i, (w, _) in enumerate(toks) if w == forma_no)
        i_te = next(i for i, (w, _) in enumerate(toks) if w == forma_termo)
        return relacao_dependencia(
            doc, toks[i_no][1], toks[i_te][1],
            preps_genitivo=pt_cfg.preps_genitivo,
            preps_associativa=pt_cfg.preps_associativa,
        )

    def test_atributiva_textura_uniforme(self):
        r = self._classifica(
            "a passagem apresenta uma textura uniforme de cordas entrelaçadas",
            "textura", "uniforme",
        )
        self.assertEqual(r["relacao_sintactica"], "atributiva")
        self.assertTrue(r["nuclear"])

    def test_genitiva_uniformidade_da_textura(self):
        r = self._classifica(
            "a uniformidade da textura caracteriza o segundo tema",
            "textura", "uniformidade",
        )
        self.assertEqual(r["relacao_sintactica"], "nominal_genitiva")
        self.assertTrue(r["nuclear"])


@pytest.mark.core
class TestGenitivoDEScaffold(unittest.TestCase):
    @pytest.mark.xfail(
        reason="DE genitive case without von/vom not yet handled (Phase 3 TODO)",
        strict=False,
    )
    def test_genitivo_morfologico_sem_prep(self):
        """Scaffold: morphological genitive should eventually classify."""
        de = obter("de")
        try:
            import spacy
            nlp = spacy.load(
                de.modelo_spacy, disable=["ner", "lemmatizer", "textcat"]
            )
        except Exception as exc:  # noqa: BLE001
            raise unittest.SkipTest(
                f"modelo {de.modelo_spacy} indisponível: {exc}"
            ) from exc
        ctx = normaliza(
            "die Gleichmässigkeit der Textur bestimmt den Satz"
        )
        doc = nlp(ctx)
        toks = tokeniza(ctx)
        i_no = next(i for i, (w, _) in enumerate(toks) if w == "textur")
        i_te = next(
            i for i, (w, _) in enumerate(toks) if w.startswith("gleich")
        )
        r = relacao_dependencia(
            doc, toks[i_no][1], toks[i_te][1],
            preps_genitivo=de.preps_genitivo,
        )
        self.assertEqual(r["relacao_sintactica"], "nominal_genitiva")


if __name__ == "__main__":
    unittest.main()
