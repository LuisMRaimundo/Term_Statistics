# -*- coding: utf-8 -*-
"""FR golden: matriz_fr_miniatura under ``--lingua fr`` + ``fr_core_news_sm``.

Separate matrix so French ``texture`` windows never enter the EN golden.
Locks attributive / genitive / heterogeneous coordination — the minimum
needed for dissertation FR counts under a validated path.
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
MATRIZ = FIXTURES / "matriz_fr_miniatura.xlsx"
CAMPO = FIXTURES / "campo_miniatura.txt"
GOLDEN = FIXTURES / "golden_fr_near.json"


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


def _fr_model_ok() -> bool:
    try:
        import spacy
        spacy.load("fr_core_news_sm", disable=["ner", "lemmatizer", "textcat"])
        return True
    except Exception:
        return False


def _norm_revisao(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v != v:
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


@pytest.fixture(scope="module")
def fr_near_xlsx(tmp_path_factory) -> Path:
    if not _fr_model_ok():
        pytest.skip("fr_core_news_sm unavailable")
    assert GOLDEN.is_file()
    assert MATRIZ.is_file()
    out_dir = tmp_path_factory.mktemp("golden_fr_near")
    out = out_dir / "golden_fr_near.xlsx"
    proc = _run([
        "textura_near.py",
        "--xlsx", str(MATRIZ),
        "--saida", str(out),
        "--near", "4",
        "--lingua", "fr",
        "--termos", str(CAMPO),
        "--dominio-omissao", "musicologia",
    ])
    assert proc.returncode == 0, (
        f"textura_near --lingua fr failed ({proc.returncode})\n"
        f"STDOUT:\n{proc.stdout}\nSTDERR:\n{proc.stderr}"
    )
    assert "fr_core_news_sm" in proc.stdout
    assert out.is_file()
    return out


class TestFrNearGolden:
    def test_counts_and_relacoes(self, fr_near_xlsx: Path):
        expect = json.loads(GOLDEN.read_text(encoding="utf-8"))
        df = pd.read_excel(fr_near_xlsx, sheet_name="8_Concordancia")
        nuc = _truthy_nuclear(df["nuclear"])
        assert len(df) == expect["n_hits"]
        assert int(nuc.sum()) == expect["n_nuclear"]

        got = [
            {
                "no": str(r["no"]),
                "termo_tipo": str(r["termo_tipo"]),
                "matched_form": str(r["matched_form"]).lower(),
                "relacao_sintactica": str(r["relacao_sintactica"]),
                "revisao_sugerida": _norm_revisao(r.get("revisao_sugerida")),
                "nuclear": bool(n),
            }
            for r, n in zip(df.to_dict("records"), nuc.tolist())
        ]
        want = expect["rows"]
        assert len(got) == len(want)
        for g, w in zip(got, want):
            assert g["no"] == w["no"]
            assert g["termo_tipo"] == w["termo_tipo"]
            assert g["matched_form"] == w["matched_form"]
            assert g["relacao_sintactica"] == w["relacao_sintactica"]
            assert g["revisao_sugerida"] == w.get("revisao_sugerida", "")
            assert g["nuclear"] is w["nuclear"]

        revs = " | ".join(g["revisao_sugerida"] for g in got)
        assert "genitiva_por_complemento" in revs
        assert "coordenacao_heterogenea" in revs
