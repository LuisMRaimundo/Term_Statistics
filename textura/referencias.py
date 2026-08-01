#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Referências bibliográficas (APA 7) — inventário e (fases seguintes) formatação.

Princípio: nunca inventar metadados. A Fase 0 só inventaria obras a partir
da folha ``8_Concordancia`` e classifica ``tipo_provavel`` de forma
*indicativa* (heurística de nome/1.ª página), sem preencher campos APA.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable
from urllib.parse import unquote, urlparse

import pandas as pd

# Prefixo típico do corpus local (CI não o possui — usar ``--raiz-corpus``).
PREFIXOS_CORPUS_OMISSAO: tuple[str, ...] = (
    r"E:\todos os textos",
    r"E:/todos os textos",
    "file:///E:/todos%20os%20textos",
    "file:///E:/todos os textos",
    "file://E:/todos%20os%20textos",
)

COLUNAS_INVENTARIO: tuple[str, ...] = (
    "doc_id",
    "caminho_ficheiro",
    "n_hits",
    "n_hits_nucleares",
    "ficheiro_existe",
    "tipo_provavel",
)

_TIPOS = (
    "grove",
    "tese",
    "actas",
    "artigo",
    "livro",
    "desconhecido",
)

# Nomes de ficheiro usam «_»; não exigir \b (underscore é carácter de palavra).
_RX_GROVE = re.compile(
    r"grove\s+music\s+online|oxford\s+music\s+online", re.I)
_RX_TESE = re.compile(
    r"thesis|dissertation|doutoramento|doutorado|(?<![a-z])tese(?![a-z])|"
    r"ph\.?\s*d\.?|m\.?\s*phil|mestrado|master'?s\s+thesis",
    re.I,
)
_RX_ACTAS = re.compile(
    r"proceedings|actas|atas|conference|symposium|congresso|"
    r"colloquium|workshop",
    re.I,
)
_RX_ARTIGO = re.compile(
    r"journal|revista|article|artigo|vol\.?\s*\d|pp\.?\s*\d",
    re.I,
)
_RX_LIVRO = re.compile(
    r"oxford\s+university\s+press|cambridge\s+university\s+press|"
    r"routledge|springer|handbook|edited\s+by",
    re.I,
)


def _truthy_nuclear(val) -> bool:
    if isinstance(val, bool):
        return val
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return False
    return str(val).strip().lower() in {
        "1", "true", "verdadeiro", "yes", "sim", "t", "y",
    }


def _strip_file_url(caminho: str) -> str:
    s = str(caminho or "").strip()
    if not s:
        return ""
    if s.lower().startswith("file:"):
        parsed = urlparse(s)
        path = unquote(parsed.path or "")
        # Windows: file:///E:/foo → /E:/foo
        if re.match(r"^/[A-Za-z]:", path):
            path = path[1:]
        return path.replace("/", "\\") if re.match(r"^[A-Za-z]:", path) else path
    return s


def remapear_caminho(
    caminho: str,
    raiz_corpus: Path | None = None,
    prefixos_origem: Iterable[str] | None = None,
) -> Path:
    """Resolve o caminho local, opcionalmente remapeando o prefixo do corpus.

    Se ``raiz_corpus`` for dada e o caminho (após descodificar ``file://``)
    começar por um dos ``prefixos_origem``, a parte relativa é juntada a
    ``raiz_corpus``. Caso contrário devolve o path tal qual (normalizado).
    """
    bruto = _strip_file_url(caminho)
    if not bruto:
        return Path()
    p = Path(bruto)
    if raiz_corpus is None:
        return p
    raiz = Path(raiz_corpus)
    prefixos = tuple(prefixos_origem) if prefixos_origem else PREFIXOS_CORPUS_OMISSAO
    # Comparação casefold + barras normalizadas
    norm = bruto.replace("\\", "/").casefold()
    for pref in prefixos:
        pref_s = _strip_file_url(pref).replace("\\", "/").rstrip("/")
        pref_n = pref_s.casefold()
        if norm == pref_n or norm.startswith(pref_n + "/"):
            rel = bruto.replace("\\", "/")[len(pref_s):].lstrip("/\\")
            return (raiz / rel) if rel else raiz
    return p


def ler_amostra_pdf(caminho: Path, *, max_paginas: int = 1,
                    max_chars: int = 2500) -> str:
    """Texto das primeiras páginas (vazio se ficheiro ausente / ilegível)."""
    if not caminho.is_file():
        return ""
    try:
        import fitz
    except ImportError:
        return ""
    try:
        doc = fitz.open(caminho)
    except Exception:
        return ""
    try:
        partes: list[str] = []
        n = min(max_paginas, doc.page_count)
        for i in range(n):
            try:
                partes.append(doc.load_page(i).get_text("text") or "")
            except Exception:
                continue
        return "\n".join(partes)[:max_chars]
    finally:
        doc.close()


def tipo_provavel(caminho: str, texto_pagina1: str = "") -> str:
    """Heurística *indicativa* — não é classificação bibliográfica final."""
    nome = Path(_strip_file_url(caminho)).name
    blob = f"{nome}\n{texto_pagina1}"
    if _RX_GROVE.search(blob):
        return "grove"
    if _RX_TESE.search(blob):
        return "tese"
    if _RX_ACTAS.search(blob):
        return "actas"
    if _RX_ARTIGO.search(blob):
        return "artigo"
    if _RX_LIVRO.search(blob):
        return "livro"
    return "desconhecido"


def _caminho_representativo(serie: pd.Series) -> str:
    vals = [
        str(v).strip() for v in serie
        if v is not None and not (isinstance(v, float) and pd.isna(v))
        and str(v).strip() and str(v).strip().lower() != "nan"
    ]
    if not vals:
        return ""
    # Mais frequente; empate → ordem lexicográfica estável
    contagem: dict[str, int] = {}
    for v in vals:
        contagem[v] = contagem.get(v, 0) + 1
    return sorted(contagem.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]


def chave_doc_id(row: pd.Series) -> str:
    """Chave de agrupamento estável (doc_id ou fallback pelo nome do ficheiro)."""
    d = row.get("doc_id")
    if d is not None and not (isinstance(d, float) and pd.isna(d)):
        s = str(d).strip()
        if s and s.lower() != "nan":
            return s
    c = row.get("caminho_ficheiro", "")
    nome = Path(_strip_file_url(str(c))).name if c is not None else ""
    return f"sem_doc_id::{nome}" if nome else "sem_doc_id::"


def chaves_doc_id(df: pd.DataFrame) -> list[str]:
    return [chave_doc_id(r) for _, r in df.iterrows()]


def construir_inventario(
    df: pd.DataFrame,
    *,
    raiz_corpus: Path | None = None,
    prefixos_origem: Iterable[str] | None = None,
    ler_pdf: bool = True,
) -> pd.DataFrame:
    """Agrupa ``8_Concordancia`` por ``doc_id`` → linhas de inventário."""
    if "doc_id" not in df.columns:
        raise ValueError(
            "Folha sem coluna 'doc_id'. Corra textura_near antes do inventário."
        )
    work = df.copy()
    work["doc_id"] = [chave_doc_id(r) for _, r in work.iterrows()]

    col_path = "caminho_ficheiro" if "caminho_ficheiro" in work.columns else None
    nuc = (
        work["nuclear"].map(_truthy_nuclear)
        if "nuclear" in work.columns
        else pd.Series(False, index=work.index)
    )
    work = work.assign(_nuclear=nuc)

    rows: list[dict] = []
    for doc_id, grp in work.groupby("doc_id", sort=True):
        caminho = _caminho_representativo(grp[col_path]) if col_path else ""
        local = remapear_caminho(
            caminho, raiz_corpus=raiz_corpus, prefixos_origem=prefixos_origem)
        existe = bool(local.is_file()) if str(local) else False
        amostra = ler_amostra_pdf(local) if (ler_pdf and existe) else ""
        rows.append({
            "doc_id": str(doc_id),
            "caminho_ficheiro": caminho,
            "n_hits": int(len(grp)),
            "n_hits_nucleares": int(grp["_nuclear"].sum()),
            "ficheiro_existe": "sim" if existe else "nao",
            "tipo_provavel": tipo_provavel(caminho, amostra),
        })
    out = pd.DataFrame(rows, columns=list(COLUNAS_INVENTARIO))
    return out.reset_index(drop=True)


def escrever_inventario_tsv(df: pd.DataFrame, saida: Path) -> Path:
    saida = Path(saida)
    saida.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(saida, sep="\t", index=False, encoding="utf-8", lineterminator="\n")
    return saida


def relatorio_inventario(df: pd.DataFrame, n_doc_excel: int) -> str:
    por_tipo = df["tipo_provavel"].value_counts().to_dict()
    n_miss = int((df["ficheiro_existe"] == "nao").sum())
    n_ok = int((df["ficheiro_existe"] == "sim").sum())
    linhas = [
        f"doc_id únicos (inventário): {len(df)}",
        f"doc_id únicos (Excel nunique): {n_doc_excel}",
        f"coincide: {'sim' if len(df) == n_doc_excel else 'NAO'}",
        f"ficheiros encontrados: {n_ok}",
        f"ficheiros em falta: {n_miss}",
        "tipo_provavel (indicativo):",
    ]
    for t in _TIPOS:
        linhas.append(f"  {t}: {int(por_tipo.get(t, 0))}")
    extras = sorted(set(por_tipo) - set(_TIPOS))
    for t in extras:
        linhas.append(f"  {t}: {int(por_tipo[t])}")
    return "\n".join(linhas)
