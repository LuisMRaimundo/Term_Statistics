# -*- coding: utf-8 -*-
"""Phase 5: column dictionary stays synchronised with live structures."""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
DICT_MD = ROOT / "dados" / "dicionario_colunas.md"
GERADOR = ROOT / "utilitarios" / "gera_dicionario_colunas.py"
MATRIZ = ROOT / "tests" / "fixtures" / "matriz_miniatura.xlsx"
CAMPO = ROOT / "tests" / "fixtures" / "campo_miniatura.txt"


def _load_gerador():
    import importlib.util

    spec = importlib.util.spec_from_file_location("gera_dicionario_colunas", GERADOR)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    return mod


def test_dicionario_matches_live_priority_list():
    from textura.exportacao import COLUNAS_HITS_PRIORIDADE, DESCRICAO_COLUNAS_HITS

    assert DICT_MD.is_file(), f"missing {DICT_MD}; run {GERADOR.name}"
    missing_desc = [
        c for c in COLUNAS_HITS_PRIORIDADE if c not in DESCRICAO_COLUNAS_HITS
    ]
    assert missing_desc == [], f"columns without description: {missing_desc}"
    fresh = _load_gerador().render()
    assert DICT_MD.read_text(encoding="utf-8") == fresh
    cols_in_doc = re.findall(
        r"^\| `([^`]+)` \|", DICT_MD.read_text(encoding="utf-8"), re.M
    )
    assert cols_in_doc == list(COLUNAS_HITS_PRIORIDADE)


def test_dicionario_lists_revisao_vocabulary():
    from textura.revisao import REVISAO_EXACTAS, REVISAO_PREFIXOS

    text = DICT_MD.read_text(encoding="utf-8")
    for t in REVISAO_EXACTAS:
        assert f"`{t}`" in text
    for t in REVISAO_PREFIXOS:
        assert f"`{t}" in text


def test_manual_points_to_generated_dictionary():
    manual = (ROOT / "TECHNICAL_MANUAL.md").read_text(encoding="utf-8")
    assert "dados/dicionario_colunas.md" in manual
    guia = (ROOT / "GUIA_REVISAO_FASE1.md").read_text(encoding="utf-8")
    assert "dicionario_colunas" in guia or "revisao_sugerida" in guia


@pytest.fixture(scope="module")
def near_xlsx_for_dict(tmp_path_factory) -> Path:
    out = tmp_path_factory.mktemp("dict_near") / "near.xlsx"
    proc = subprocess.run(
        [
            sys.executable, "textura_near.py",
            "--xlsx", str(MATRIZ),
            "--saida", str(out),
            "--near", "4",
            "--lingua", "todas",
            "--termos", str(CAMPO),
            "--dominio-omissao", "musicologia",
        ],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    assert proc.returncode == 0, proc.stderr + proc.stdout
    return out


def test_workbook_columns_subset_of_dictionary(near_xlsx_for_dict: Path):
    from textura.exportacao import COLUNAS_HITS_PRIORIDADE

    df = pd.read_excel(near_xlsx_for_dict, sheet_name="8_Concordancia")
    extras = [c for c in df.columns if c not in COLUNAS_HITS_PRIORIDADE]
    assert extras == [], f"workbook columns missing from live dictionary: {extras}"


def test_revisao_sugerida_tags_in_vocabulary(near_xlsx_for_dict: Path):
    from textura.revisao import validar_revisao_sugerida

    df = pd.read_excel(near_xlsx_for_dict, sheet_name="8_Concordancia")
    bad = []
    for i, blob in enumerate(df["revisao_sugerida"].fillna("")):
        inv = validar_revisao_sugerida(blob)
        if inv:
            bad.append((i, blob, inv))
    assert bad == [], f"invalid revisao_sugerida tags: {bad[:5]}"
