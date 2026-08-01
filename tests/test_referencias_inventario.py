# -*- coding: utf-8 -*-
"""Fase 0 — inventário de referências (doc_id, remapeamento, tipo_provavel)."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from textura.referencias import (
    COLUNAS_INVENTARIO,
    chaves_doc_id,
    construir_inventario,
    remapear_caminho,
    tipo_provavel,
    escrever_inventario_tsv,
)

ROOT = Path(__file__).resolve().parents[1]


def _pdf(path: Path, texto: str) -> Path:
    fitz = pytest.importorskip("fitz")
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), texto, fontsize=11)
    doc.save(str(path))
    doc.close()
    return path


def test_remapear_caminho_prefixo_e_file_url(tmp_path: Path):
    raiz = tmp_path / "corpus"
    alvo = raiz / "sub" / "obra.pdf"
    alvo.parent.mkdir(parents=True)
    alvo.write_bytes(b"%PDF-1.4")

    p1 = remapear_caminho(
        r"E:\todos os textos\sub\obra.pdf", raiz_corpus=raiz)
    assert p1 == alvo
    assert p1.is_file()

    p2 = remapear_caminho(
        "file:///E:/todos%20os%20textos/sub/obra.pdf", raiz_corpus=raiz)
    assert p2 == alvo

    # sem raiz: path tal qual
    p3 = remapear_caminho(r"E:\todos os textos\sub\obra.pdf")
    assert p3 == Path(r"E:\todos os textos\sub\obra.pdf")


def test_tipo_provavel_heuristica():
    assert tipo_provavel("(2010)_Grove_entry.pdf",
                         "Grove Music Online Oxford") == "grove"
    assert tipo_provavel("My_PhD_thesis.pdf", "") == "tese"
    assert tipo_provavel("ISMIR_proceedings_2019.pdf", "") == "actas"
    assert tipo_provavel("paper_in_Journal_vol_12.pdf", "") == "artigo"
    assert tipo_provavel("random_scan.pdf", "") == "desconhecido"


def test_construir_inventario_sem_perder_doc_ids(tmp_path: Path):
    raiz = tmp_path / "corpus"
    grove = _pdf(
        raiz / "Grove_Xenakis.pdf",
        "Grove Music Online\nXenakis, Iannis",
    )
    tese = _pdf(raiz / "Smith_PhD_thesis.pdf", "A doctoral dissertation")
    # ficheiro anunciado mas ausente
    ausente_rel = "Missing_Book.pdf"

    df = pd.DataFrame([
        {
            "doc_id": "aaa111",
            "caminho_ficheiro": rf"E:\todos os textos\{grove.name}",
            "nuclear": True,
            "contexto": "texture a",
        },
        {
            "doc_id": "aaa111",
            "caminho_ficheiro": rf"E:\todos os textos\{grove.name}",
            "nuclear": False,
            "contexto": "texture b",
        },
        {
            "doc_id": "bbb222",
            "caminho_ficheiro": rf"E:\todos os textos\{tese.name}",
            "nuclear": True,
            "contexto": "texture c",
        },
        {
            "doc_id": "ccc333",
            "caminho_ficheiro": rf"E:\todos os textos\{ausente_rel}",
            "nuclear": True,
            "contexto": "texture d",
        },
    ])
    n_ref = len(set(chaves_doc_id(df)))
    inv = construir_inventario(df, raiz_corpus=raiz, ler_pdf=True)
    assert len(inv) == n_ref == 3
    assert list(inv.columns) == list(COLUNAS_INVENTARIO)

    by_id = inv.set_index("doc_id")
    assert int(by_id.loc["aaa111", "n_hits"]) == 2
    assert int(by_id.loc["aaa111", "n_hits_nucleares"]) == 1
    assert by_id.loc["aaa111", "ficheiro_existe"] == "sim"
    assert by_id.loc["aaa111", "tipo_provavel"] == "grove"
    assert by_id.loc["bbb222", "tipo_provavel"] == "tese"
    assert by_id.loc["ccc333", "ficheiro_existe"] == "nao"

    out = escrever_inventario_tsv(inv, tmp_path / "inventario.tsv")
    reload = pd.read_csv(out, sep="\t")
    assert len(reload) == 3


def test_cli_inventario(tmp_path: Path):
    import subprocess
    import sys

    raiz = tmp_path / "corpus"
    pdf = _pdf(raiz / "Oxford_Music_Online_entry.pdf",
               "Oxford Music Online\nentry text")
    xlsx = tmp_path / "mini_near.xlsx"
    df = pd.DataFrame([{
        "doc_id": "d1",
        "caminho_ficheiro": rf"E:\todos os textos\{pdf.name}",
        "nuclear": True,
        "contexto": "a homogeneous texture appears",
        "query_pattern": "homogeneous*",
        "canonical_term": "homogeneous",
        "url": "",
    }])
    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="8_Concordancia", index=False)

    saida = tmp_path / "inventario.tsv"
    proc = subprocess.run(
        [
            sys.executable,
            str(ROOT / "utilitarios" / "gera_referencias.py"),
            "--inventario",
            "--xlsx", str(xlsx),
            "--saida", str(saida),
            "--raiz-corpus", str(raiz),
        ],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "coincide: sim" in proc.stdout
    assert saida.is_file()
    inv = pd.read_csv(saida, sep="\t")
    assert len(inv) == 1
    assert inv.iloc[0]["tipo_provavel"] == "grove"
