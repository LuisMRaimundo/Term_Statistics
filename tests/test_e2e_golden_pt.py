# -*- coding: utf-8 -*-
"""PT golden: miniature matrix under ``--lingua pt`` + ``pt_core_news_sm``.

Locks hit counts and ``relacao_sintactica`` so CI model/code drift on the PT
path cannot pass silently. EN golden (``test_e2e_golden.py``) stays separate
and byte-stable; PT rows live in the shared miniature but are selected by NOS.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = Path(__file__).resolve().parent / "fixtures"
MATRIZ = FIXTURES / "matriz_miniatura.xlsx"
CAMPO = FIXTURES / "campo_miniatura.txt"
GOLDEN = FIXTURES / "golden_pt_near.json"


def _truthy_nuclear(s: pd.Series) -> pd.Series:
    return s.map(
        lambda v: v is True or str(v).strip().lower() in {"true", "1", "sim"}
    )


def _run(args: list[str]) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, *args],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


def _pt_model_ok() -> bool:
    try:
        import spacy
        spacy.load("pt_core_news_sm", disable=["ner", "lemmatizer", "textcat"])
        return True
    except Exception:
        return False


@pytest.fixture(scope="module")
def pt_near_xlsx(tmp_path_factory) -> Path:
    if not _pt_model_ok():
        pytest.skip("pt_core_news_sm unavailable")
    assert GOLDEN.is_file()
    out_dir = tmp_path_factory.mktemp("golden_pt_near")
    out = out_dir / "golden_pt_near.xlsx"
    proc = _run([
        "textura_near.py",
        "--xlsx", str(MATRIZ),
        "--saida", str(out),
        "--near", "4",
        "--lingua", "pt",
        "--termos", str(CAMPO),
        "--dominio-omissao", "musicologia",
    ])
    assert proc.returncode == 0, (
        f"textura_near --lingua pt failed ({proc.returncode})\n"
        f"STDOUT:\n{proc.stdout}\nSTDERR:\n{proc.stderr}"
    )
    assert "pt_core_news_sm" in proc.stdout
    assert out.is_file()
    return out


class TestPtNearGolden:
    def test_counts_and_relacoes(self, pt_near_xlsx: Path):
        expect = json.loads(GOLDEN.read_text(encoding="utf-8"))
        df = pd.read_excel(pt_near_xlsx, sheet_name="8_Concordancia")
        nuc = _truthy_nuclear(df["nuclear"])
        assert len(df) == expect["n_hits"]
        assert int(nuc.sum()) == expect["n_nuclear"]

        got = [
            {
                "no": str(r["no"]),
                "termo_tipo": str(r["termo_tipo"]),
                "matched_form": str(r["matched_form"]).lower(),
                "relacao_sintactica": str(r["relacao_sintactica"]),
                "nuclear": bool(n),
            }
            for r, n in zip(df.to_dict("records"), nuc.tolist())
        ]
        # Order-stable on this fixture (matrix row order)
        want = expect["rows"]
        assert len(got) == len(want)
        for g, w in zip(got, want):
            assert g["no"] == w["no"]
            assert g["termo_tipo"] == w["termo_tipo"]
            assert g["matched_form"] == w["matched_form"]
            assert g["relacao_sintactica"] == w["relacao_sintactica"]
            assert g["nuclear"] is w["nuclear"]
