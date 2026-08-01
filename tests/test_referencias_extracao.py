# -*- coding: utf-8 -*-
"""Fase 1 — extracção com evidência (PDFs-miniatura)."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest

from textura.referencias import (
    CAMPOS_BIBLIO,
    assert_evidencias_consistentes,
    construir_rascunho,
    extrair_obra,
    marcar_duplicados,
    RascunhoObra,
)

ROOT = Path(__file__).resolve().parents[1]
FIXT = Path(__file__).resolve().parent / "fixtures" / "referencias"
BUILDER = FIXT / "build_pdfs_miniatura.py"


@pytest.fixture(scope="module")
def pdfs_miniatura() -> dict[str, Path]:
    pytest.importorskip("fitz")
    subprocess.run(
        [sys.executable, str(BUILDER)],
        cwd=str(ROOT),
        check=True,
        capture_output=True,
    )
    return {
        "artigo": FIXT / "(2019)_Texture_and_Sonata_Form.pdf",
        "grove": FIXT / "Grove_Xenakis_entry.pdf",
        "vazio": FIXT / "scan_sem_metadados.pdf",
    }


def test_artigo_rodape_e_meta(pdfs_miniatura: dict[str, Path]):
    path = pdfs_miniatura["artigo"]
    r = extrair_obra(
        "doc_art",
        rf"E:\todos os textos\{path.name}",
        raiz_corpus=path.parent,
    )
    d = r.para_dict()
    assert d["tipo"] == "artigo"
    assert d["evidencia_tipo"] in {"rodape", "pagina1", "nome_ficheiro"}
    assert "Journal of Music Theory" in d["contentor"]
    assert d["evidencia_contentor"] == "rodape"
    assert d["volume"] == "63"
    assert d["numero"] == "2"
    assert "145" in d["paginas"]
    assert d["ano"] == "2019"
    assert "10.1234/jmt.2019.001" in d["doi_ou_url"]
    assert d["titulo"]
    assert d["evidencia_titulo"] != "vazio"
    assert d["autores"]
    # meta coerente pode contribuir; nunca vazio com valor
    for c in CAMPOS_BIBLIO:
        if d[c]:
            assert d[f"evidencia_{c}"] != "vazio"


def test_grove_verbete_sem_ano_de_download(pdfs_miniatura: dict[str, Path]):
    path = pdfs_miniatura["grove"]
    r = extrair_obra("doc_grove", str(path), raiz_corpus=None)
    d = r.para_dict()
    assert d["tipo"] == "verbete"
    assert d["contentor"] == "Grove Music Online"
    assert d["editora"] == "Oxford University Press"
    assert d["evidencia_contentor"] == "pagina1"
    assert "Xenakis" in d["titulo"]
    assert "Hoffmann" in d["autores"]
    assert d["ano"] == "2001"  # published online no PDF
    assert "10.1093/gmo" in d["doi_ou_url"]
    # meta lixo rejeitada
    assert "Microsoft" not in d["titulo"]
    assert d["autores"].casefold() != "admin"


def test_grove_ano_vazio_se_nao_impresso(pdfs_miniatura: dict[str, Path], tmp_path: Path):
    fitz = pytest.importorskip("fitz")
    path = tmp_path / "Grove_no_date.pdf"
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), "Grove Music Online", fontsize=14)
    page.insert_text((72, 100), "Texture", fontsize=12)
    page.insert_text((72, 130), "by Smith, A.", fontsize=11)
    doc.save(str(path))
    doc.close()
    r = extrair_obra("g2", str(path))
    d = r.para_dict()
    assert d["tipo"] == "verbete"
    assert d["ano"] == ""
    assert d["evidencia_ano"] == "vazio"
    assert d["verificar"] == "sim"


def test_sem_metadados_tudo_vazio_verificar(pdfs_miniatura: dict[str, Path]):
    path = pdfs_miniatura["vazio"]
    r = extrair_obra("doc_vazio", str(path))
    d = r.para_dict()
    assert d["tipo"] == "desconhecido"
    assert d["verificar"] == "sim"
    assert d["confianca"] == "baixa"
    for c in CAMPOS_BIBLIO:
        if c == "tipo":
            continue
        assert d[c] == "", f"{c} deveria estar vazio"
        assert d[f"evidencia_{c}"] == "vazio"


def test_nenhum_campo_sem_evidencia(pdfs_miniatura: dict[str, Path]):
    rows = []
    for key, doc_id in (("artigo", "a"), ("grove", "g"), ("vazio", "v")):
        path = pdfs_miniatura[key]
        rows.append({
            "doc_id": doc_id,
            "caminho_ficheiro": str(path),
            "nuclear": True,
            "contexto": "texture",
        })
    df = pd.DataFrame(rows)
    out = construir_rascunho(df, raiz_corpus=None)
    assert_evidencias_consistentes(out)
    assert len(out) == 3


def test_duplicado_de():
    a = RascunhoObra("id1")
    a.set("autores", "Smith, J.", "pagina1")
    a.set("ano", "2019", "pagina1")
    a.set("titulo", "Texture Studies", "pagina1")
    a.set("tipo", "artigo", "pagina1")
    b = RascunhoObra("id2")
    b.set("autores", "Smith, J.", "pagina1")
    b.set("ano", "2019", "pagina1")
    b.set("titulo", "Texture Studies", "pagina1")
    b.set("tipo", "artigo", "pagina1")
    marcar_duplicados([a, b])
    assert a.duplicado_de == ""
    assert b.duplicado_de == "id1"


def test_cli_extrair(pdfs_miniatura: dict[str, Path], tmp_path: Path):
    path = pdfs_miniatura["artigo"]
    xlsx = tmp_path / "mini.xlsx"
    df = pd.DataFrame([{
        "doc_id": "x1",
        "caminho_ficheiro": rf"E:\todos os textos\{path.name}",
        "nuclear": True,
        "contexto": "a homogeneous texture",
        "query_pattern": "homogeneous*",
        "canonical_term": "homogeneous",
        "url": "",
    }])
    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="8_Concordancia", index=False)
        pd.DataFrame({"caminho": [str(path)]}).to_excel(
            xw, sheet_name="Manifesto_corpus", index=False)

    saida = tmp_path / "rascunho.tsv"
    proc = subprocess.run(
        [
            sys.executable,
            str(ROOT / "utilitarios" / "gera_referencias.py"),
            "--extrair",
            "--xlsx", str(xlsx),
            "--saida", str(saida),
            "--raiz-corpus", str(path.parent),
        ],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "obras (doc_id): 1" in proc.stdout
    out = pd.read_csv(saida, sep="\t", dtype=str).fillna("")
    assert_evidencias_consistentes(out)
    assert out.iloc[0]["tipo"] == "artigo"
