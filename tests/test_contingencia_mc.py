# -*- coding: utf-8 -*-
"""Monte Carlo contingency: degenerate sims must not deflate p."""

from __future__ import annotations

import unittest

import numpy as np
import pandas as pd

import textura_analise as ta


class TestContingenciaMonteCarlo(unittest.TestCase):
    def test_structural_zero_p_not_deflated(self):
        """2×2 with a structural zero → MC path; p must not be anticonservative.

        Degenerate multinomial draws (zero expected cell) are resampled /
        skipped, not scored as χ²=0. A biased guard would push p toward 1.
        """
        tab = pd.DataFrame(
            [[5, 0], [1, 3]],
            index=["atributiva", "predicativa"],
            columns=["estabilidade", "variabilidade"],
        )
        r = ta.teste_contingencia(tab, n_perm=2000, semente=20260725)
        self.assertIn("Monte Carlo", r["metodo"])
        self.assertIn("validas", r["metodo"])
        self.assertTrue(np.isfinite(r["p"]), r)
        # Observed association is strong (5 vs 0 on first row); an honest
        # MC p should stay clearly below a deflated ceiling near 1.0.
        self.assertLess(r["p"], 0.25, r)

    def test_skip_path_keeps_honest_denominator(self):
        """Forced ValueError on sims → still finishes with finite p and 'validas'."""
        tab = pd.DataFrame([[4, 1], [1, 3]])
        real = ta.stats.chi2_contingency
        state = {"calls": 0}

        def wrapper(x, correction=False):
            state["calls"] += 1
            arr = np.asarray(x)
            # After the observed-table calls (setup + cramers_v), fail some sims
            # that look like MC draws (same shape, later call index).
            if state["calls"] > 3 and arr.shape == (2, 2) and state["calls"] % 3 == 0:
                raise ValueError("expected zero (forced)")
            return real(x, correction=correction)

        orig = ta.stats.chi2_contingency
        ta.stats.chi2_contingency = wrapper
        try:
            r = ta.teste_contingencia(tab, n_perm=30, semente=7)
        finally:
            ta.stats.chi2_contingency = orig
        self.assertIn("validas", r["metodo"])
        self.assertTrue(np.isfinite(r["p"]), r)


if __name__ == "__main__":
    unittest.main()
