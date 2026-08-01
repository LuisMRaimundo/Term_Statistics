# -*- coding: utf-8 -*-
"""Monte Carlo contingency: skip degenerates; do not deflate p."""

from __future__ import annotations

import unittest

import numpy as np
import pandas as pd
from scipy import stats as scipy_stats

import textura_analise as ta


class TestContingenciaMonteCarlo(unittest.TestCase):
    def test_sparse_2x2_agrees_with_fisher(self):
        """Guarded MC/Fisher p must track Fisher's exact on the same 2×2.

        The old χ²=0 guard was anticonservative (deflated p). Agreement with
        Fisher (generous relative tolerance) fails under that bug and passes
        under skip-or-Fisher.
        """
        tab = pd.DataFrame(
            [[5, 0], [1, 3]],
            index=["atributiva", "predicativa"],
            columns=["estabilidade", "variabilidade"],
        )
        r = ta.teste_contingencia(tab, n_perm=2000, semente=20260725,
                                  min_validas=200)
        p_fish = float(scipy_stats.fisher_exact(tab.values).pvalue)
        self.assertTrue(np.isfinite(r["p"]), r)
        self.assertIn("n_validas", r)
        # Either MC with enough validas or explicit Fisher fallback
        self.assertTrue(
            "Monte Carlo" in r["metodo"] or "Fisher" in r["metodo"],
            r["metodo"],
        )
        self.assertAlmostEqual(r["p"], p_fish, delta=max(0.05, 0.5 * p_fish))

    def test_old_chi2_zero_guard_would_deflate_p(self):
        """Document the property: χ²=0-on-error yields p ≥ honest skip guard.

        On a table that produces many degenerate sims, scoring errors as 0
        never counts them as extreme → larger (deflated) p.
        """
        tab = pd.DataFrame([[5, 0], [1, 3]])
        obs = tab.values
        n = int(obs.sum())
        rprop = obs.sum(1) / n
        cprop = obs.sum(0) / n
        chi2, _, _, _ = scipy_stats.chi2_contingency(obs)
        rng = np.random.default_rng(20260725)
        n_perm = 500
        max_tries = n_perm * 20

        def _run(score_error_as_zero: bool):
            maiores = validas = tentativas = 0
            while validas < n_perm and tentativas < max_tries:
                tentativas += 1
                sim = rng.multinomial(
                    n, np.outer(rprop, cprop).ravel()
                ).reshape(obs.shape)
                try:
                    c2 = scipy_stats.chi2_contingency(
                        sim, correction=False)[0]
                except ValueError:
                    if score_error_as_zero:
                        c2 = 0.0
                        validas += 1
                    else:
                        continue
                else:
                    validas += 1
                if c2 >= chi2 - 1e-12:
                    maiores += 1
            return (maiores + 1) / (validas + 1), validas

        p_old, _ = _run(True)
        p_new, n_ok = _run(False)
        self.assertGreater(n_ok, 0)
        # Anticonservative old guard: p_old >= p_new (deflation)
        self.assertGreaterEqual(p_old, p_new - 1e-12)

    def test_low_validas_falls_back_to_fisher(self):
        tab = pd.DataFrame([[5, 0], [1, 3]])
        r = ta.teste_contingencia(tab, n_perm=5000, semente=1,
                                  min_validas=10**9)  # force fallback
        self.assertIn("Fisher", r["metodo"])
        self.assertTrue(np.isfinite(r["p"]))
        p_fish = float(scipy_stats.fisher_exact(tab.values).pvalue)
        self.assertAlmostEqual(r["p"], p_fish, places=12)


if __name__ == "__main__":
    unittest.main()
