# -*- coding: utf-8 -*-
"""Overflow-safe Bayes factor for large concordance n."""

from __future__ import annotations

import math
import unittest

from textura_stats import bayes_factor_proporcao


class BayesFactorProporcaoTests(unittest.TestCase):
    def test_balanced_small(self):
        bf = bayes_factor_proporcao(5, 10, p0=0.5)
        self.assertTrue(math.isfinite(bf))
        self.assertGreater(bf, 0)

    def test_large_n_no_overflow(self):
        # Case that crashed Compósita análise (~1157 nucleares, proporção extrema):
        # BF10 → ∞ (evidência forte contra p=0,5); nunca OverflowError.
        bf = bayes_factor_proporcao(1157, 1157, p0=0.5)
        self.assertEqual(bf, float("inf"))
        bf0 = bayes_factor_proporcao(0, 1157, p0=0.5)
        self.assertEqual(bf0, float("inf"))
        mid = bayes_factor_proporcao(578, 1157, p0=0.5)
        self.assertTrue(math.isfinite(mid))
        self.assertLess(mid, 1.0)

    def test_invalid(self):
        self.assertTrue(math.isnan(bayes_factor_proporcao(1, 0)))
        self.assertTrue(math.isnan(bayes_factor_proporcao(-1, 10)))
        self.assertTrue(math.isnan(bayes_factor_proporcao(3, 10, p0=0)))


if __name__ == "__main__":
    unittest.main()
