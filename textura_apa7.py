#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
textura_apa7.py — catálogo APA 7.ª ed. a partir das hiperligações
=================================================================

Script **à parte** da fase 3 (DOCX). Não mexe na localização de páginas
no PDF (isso continua em ``textura_apendice.py --paginas-pdf``).

Problema: quando falta ``fonte_apa`` / catálogo ``--refs``, o apêndice
cai no fallback do *filename* (``Smith_2019_Texture.pdf`` → título
corrupto). Este script:

  1. Lê as obras únicas de ``8_Concordancia`` (``url`` / ``caminho_ficheiro``).
  2. Segue a hiperligação (DOI → Crossref; URL → metadados HTML; PDF local
     → metadados / DOI embutido).
  3. Formata a referência em APA 7 (regras alinhadas com APA Style /
     Purdue OWL: autor invertido, ano, *sentence case* no título da obra,
     itálico em periódicos/livros via ``*…*``, DOI como ``https://doi.org/…``
     **sem** ponto final).
  4. Escreve um catálogo para ``textura_apendice.py --refs …`` e, opcionalmente,
     a coluna ``fonte_apa`` no próprio Excel.

Uso:
    python textura_apa7.py --xlsx UNIFORME_near_revisto.xlsx
    python textura_apa7.py --xlsx UNIFORME_near.xlsx --saida refs_apa7.xlsx
    python textura_apa7.py --xlsx UNIFORME_near.xlsx --escrever-xlsx
    python textura_apendice.py --xlsx UNIFORME_near.xlsx --refs refs_apa7.xlsx

Referência normativa (APA 7 / OWL):
  https://owl.purdue.edu/owl/research_and_citation/apa_style/apa_formatting_and_style_guide/index.html
"""

from __future__ import annotations

import argparse
import html as html_lib
import json
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

import pandas as pd

import textura_apendice as ta

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

USER_AGENT = (
    "TermStatistics-APA7/1.0 (research; mailto:local@local; "
    "compatible; +https://github.com/LuisMRaimundo/Term_Statistics)"
)
CROSSREF = "https://api.crossref.org/works/"
TIMEOUT_S = 18
_RX_DOI = re.compile(
    r"\b(10\.\d{4,9}/[-._;()/:A-Z0-9]+)\b", re.I
)
_RX_DOI_URL = re.compile(
    r"https?://(?:dx\.)?doi\.org/(10\.\d{4,9}/[-._;()/:A-Z0-9]+)", re.I
)
_RX_YEAR = re.compile(r"\b(19|20)\d{2}\b")
_RX_TAG = re.compile(r"<[^>]+>")
_RX_WS = re.compile(r"\s+")


# ---------------------------------------------------------------------------
# Modelo
# ---------------------------------------------------------------------------

@dataclass
class Meta:
    authors: list[str] = field(default_factory=list)  # "Surname, A. A."
    year: str = ""
    title: str = ""
    container: str = ""          # journal / book title
    volume: str = ""
    issue: str = ""
    pages: str = ""
    publisher: str = ""
    doi: str = ""
    url: str = ""
    tipo: str = "unknown"        # journal-article | book | chapter | webpage | pdf
    source: str = ""             # crossref | html | pdf | filename
    notes: str = ""


# ---------------------------------------------------------------------------
# APA 7 — formatação (offline)
# ---------------------------------------------------------------------------

def sentence_case(title: str) -> str:
    """APA sentence case: first word + proper-noun heuristic (minimal)."""
    t = _RX_WS.sub(" ", (title or "").strip())
    if not t:
        return ""
    # keep existing capitals after colon for subtitle start
    parts = re.split(r"(:\s*)", t)
    out: list[str] = []
    for i, part in enumerate(parts):
        if i % 2 == 1:  # separator
            out.append(part)
            continue
        if not part:
            continue
        lower = part[:1].upper() + part[1:].lower() if len(part) > 1 else part.upper()
        # restore likely acronyms / Roman numerals already ALLCAPS short tokens
        tokens = []
        for w, orig in zip(lower.split(), part.split()):
            if orig.isupper() and 2 <= len(orig) <= 5 and orig.isalpha():
                tokens.append(orig)
            else:
                tokens.append(w)
        out.append(" ".join(tokens))
    return "".join(out)


def title_case_periodical(name: str) -> str:
    """Title case for journal names (major words)."""
    small = {
        "a", "an", "the", "and", "but", "or", "for", "nor", "on", "at",
        "to", "from", "by", "of", "in", "with", "as",
    }
    words = _RX_WS.sub(" ", (name or "").strip()).split()
    out = []
    for i, w in enumerate(words):
        low = w.lower()
        if i > 0 and low in small:
            out.append(low)
        else:
            out.append(w[:1].upper() + w[1:] if w else w)
    return " ".join(out)


def format_authors_apa(authors: list[str]) -> str:
    """Authors already as 'Surname, A. A.' — join with APA 7 commas / &."""
    names = [a.strip() for a in authors if a and a.strip()]
    if not names:
        return ""
    if len(names) == 1:
        return names[0]
    if len(names) <= 20:
        return ", ".join(names[:-1]) + ", & " + names[-1]
    # 21+: first 19, ellipsis, last
    head = ", ".join(names[:19])
    return f"{head}, … {names[-1]}"


def crossref_author_to_apa(a: dict[str, Any]) -> str:
    family = (a.get("family") or "").strip()
    given = (a.get("given") or "").strip()
    if not family and given:
        return given
    if not family:
        return (a.get("name") or "").strip()
    initials = ""
    if given:
        parts = re.split(r"[\s\-]+", given)
        initials = " ".join(
            (p[0].upper() + ".") for p in parts if p and p[0].isalpha()
        )
    return f"{family}, {initials}".strip().rstrip(",")


def format_apa7(meta: Meta) -> str:
    """Build one APA 7 reference string (``*italics*`` for the appendix)."""
    author = format_authors_apa(meta.authors)
    year = (meta.year or "").strip() or "n.d."
    title = sentence_case(meta.title) if meta.title else ""
    doi = (meta.doi or "").strip()
    if doi.lower().startswith("https://doi.org/"):
        doi_url = doi
        doi = doi.split("doi.org/", 1)[-1]
    elif doi:
        doi_url = f"https://doi.org/{doi}"
    else:
        doi_url = ""

    tipo = (meta.tipo or "unknown").lower()
    bits: list[str] = []

    if author:
        bits.append(f"{author} ({year}).")
    else:
        # Title moves to front when no author (APA)
        if title:
            if tipo in {"book", "report", "pdf"}:
                bits.append(f"*{title}*. ({year}).")
            else:
                bits.append(f"{title}. ({year}).")
            title = ""  # consumed
        else:
            bits.append(f"({year}).")

    if title:
        if tipo in {"journal-article", "article", "proceedings-article"}:
            bits.append(f"{title}.")
        elif tipo in {"book", "monograph", "report", "dissertation", "pdf"}:
            bits.append(f"*{title}*.")
        elif tipo in {"book-chapter", "chapter"}:
            bits.append(f"{title}.")
        else:
            bits.append(f"{title}.")

    if tipo in {"journal-article", "article", "proceedings-article"}:
        cont = title_case_periodical(meta.container) if meta.container else ""
        if cont:
            vol = (meta.volume or "").strip()
            iss = (meta.issue or "").strip()
            pages = (meta.pages or "").strip().replace("-", "–")
            mid = f"*{cont}*"
            if vol:
                mid += f", *{vol}*"
                if iss:
                    mid += f"({iss})"
            if pages:
                mid += f", {pages}"
            bits.append(mid + ".")
    elif tipo in {"book-chapter", "chapter"}:
        cont = sentence_case(meta.container) if meta.container else ""
        if cont:
            ed_bit = f"In *{cont}*"
            pages = (meta.pages or "").strip().replace("-", "–")
            if pages:
                ed_bit += f" (pp. {pages})"
            ed_bit += "."
            bits.append(ed_bit)
        if meta.publisher:
            bits.append(f"{meta.publisher.strip()}.")
    elif tipo in {"book", "monograph", "report", "dissertation", "pdf"}:
        if meta.publisher:
            bits.append(f"{meta.publisher.strip()}.")
    elif tipo in {"webpage", "web"}:
        if meta.container:
            bits.append(f"{meta.container.strip()}.")

    if doi_url:
        bits.append(doi_url)  # no trailing period (APA 7)
    elif meta.url and not meta.url.lower().startswith("file:"):
        bits.append(meta.url.strip())

    ref = " ".join(b for b in bits if b).strip()
    ref = _RX_WS.sub(" ", ref)
    # tidy double periods (not after doi.org)
    ref = re.sub(r"\.\s*\.", ". ", ref)
    return ref


# ---------------------------------------------------------------------------
# Resolução de URL / DOI / PDF
# ---------------------------------------------------------------------------

def extrair_doi(texto: str | None) -> str:
    if not texto:
        return ""
    s = str(texto).strip()
    m = _RX_DOI_URL.search(s)
    if m:
        return m.group(1).rstrip(").,;")
    m = _RX_DOI.search(s)
    if m:
        return m.group(1).rstrip(").,;")
    if s.startswith("10.") and "/" in s:
        return s.split()[0].rstrip(").,;")
    return ""


def _http_get(url: str, *, accept: str = "*/*") -> tuple[int, bytes, str]:
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": accept,
        },
        method="GET",
    )
    with urllib.request.urlopen(req, timeout=TIMEOUT_S) as resp:
        ctype = resp.headers.get("Content-Type", "")
        return resp.getcode(), resp.read(), ctype


def fetch_crossref(doi: str) -> Optional[Meta]:
    doi = extrair_doi(doi) or doi.strip()
    if not doi:
        return None
    url = CROSSREF + urllib.parse.quote(doi)
    try:
        code, raw, _ = _http_get(url, accept="application/json")
    except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError, ValueError):
        return None
    if code != 200:
        return None
    try:
        msg = json.loads(raw.decode("utf-8", errors="replace")).get("message") or {}
    except json.JSONDecodeError:
        return None

    authors = [crossref_author_to_apa(a) for a in (msg.get("author") or [])]
    authors = [a for a in authors if a]
    year = ""
    for key in ("published-print", "published-online", "issued", "created"):
        parts = (msg.get(key) or {}).get("date-parts") or []
        if parts and parts[0]:
            year = str(parts[0][0])
            break
    titles = msg.get("title") or []
    title = titles[0] if titles else ""
    containers = msg.get("container-title") or []
    container = containers[0] if containers else ""
    page = (msg.get("page") or "").replace("-", "–")
    tipo_cr = (msg.get("type") or "").lower()
    tipo_map = {
        "journal-article": "journal-article",
        "proceedings-article": "proceedings-article",
        "book-chapter": "book-chapter",
        "book": "book",
        "monograph": "book",
        "edited-book": "book",
        "dissertation": "dissertation",
        "report": "report",
        "posted-content": "webpage",
    }
    tipo = tipo_map.get(tipo_cr, "journal-article" if container else "book")
    publisher = (msg.get("publisher") or "").strip()
    return Meta(
        authors=authors,
        year=year,
        title=title,
        container=container,
        volume=str(msg.get("volume") or ""),
        issue=str(msg.get("issue") or ""),
        pages=page,
        publisher=publisher,
        doi=doi,
        url=f"https://doi.org/{doi}",
        tipo=tipo,
        source="crossref",
    )


def _meta_content(html: str, *names: str) -> str:
    for name in names:
        # <meta name="…" content="…"> or property=
        rx = re.compile(
            rf'<meta[^>]+(?:name|property)\s*=\s*["\']{re.escape(name)}["\']'
            rf'[^>]+content\s*=\s*["\']([^"\']+)["\']',
            re.I,
        )
        m = rx.search(html)
        if m:
            return html_lib.unescape(m.group(1)).strip()
        rx2 = re.compile(
            rf'<meta[^>]+content\s*=\s*["\']([^"\']+)["\'][^>]+'
            rf'(?:name|property)\s*=\s*["\']{re.escape(name)}["\']',
            re.I,
        )
        m = rx2.search(html)
        if m:
            return html_lib.unescape(m.group(1)).strip()
    return ""


def _all_meta(html: str, name: str) -> list[str]:
    vals = []
    for m in re.finditer(
        rf'<meta[^>]+(?:name|property)\s*=\s*["\']{re.escape(name)}["\']'
        rf'[^>]+content\s*=\s*["\']([^"\']+)["\']',
        html,
        re.I,
    ):
        vals.append(html_lib.unescape(m.group(1)).strip())
    return [v for v in vals if v]


def parse_html_meta(html: str, url: str) -> Meta:
    title = (
        _meta_content(html, "citation_title", "dc.title", "DC.Title", "og:title")
        or ""
    )
    if not title:
        m = re.search(r"<title[^>]*>(.*?)</title>", html, re.I | re.S)
        if m:
            title = _RX_WS.sub(
                " ", _RX_TAG.sub("", html_lib.unescape(m.group(1)))
            ).strip()
    authors_raw = _all_meta(html, "citation_author") or _all_meta(html, "dc.creator")
    authors = []
    for a in authors_raw:
        # "Given Family" → "Family, G."
        parts = a.replace(",", " ").split()
        if len(parts) >= 2:
            family = parts[-1]
            given = parts[:-1]
            initials = " ".join(p[0].upper() + "." for p in given if p)
            authors.append(f"{family}, {initials}".strip())
        else:
            authors.append(a)
    year = _meta_content(html, "citation_publication_date", "citation_date", "dc.date")
    ym = _RX_YEAR.search(year or "")
    year = ym.group(0) if ym else (year[:4] if year and year[:4].isdigit() else "")
    journal = _meta_content(
        html, "citation_journal_title", "og:site_name", "dc.publisher"
    )
    vol = _meta_content(html, "citation_volume")
    issue = _meta_content(html, "citation_issue")
    pages = _meta_content(html, "citation_firstpage")
    last = _meta_content(html, "citation_lastpage")
    if pages and last:
        pages = f"{pages}–{last}"
    doi = extrair_doi(_meta_content(html, "citation_doi", "dc.identifier")) or extrair_doi(html[:8000])
    publisher = _meta_content(html, "citation_publisher", "dc.publisher")
    tipo = "journal-article" if journal else "webpage"
    return Meta(
        authors=authors,
        year=year,
        title=title,
        container=journal,
        volume=vol,
        issue=issue,
        pages=pages,
        publisher=publisher,
        doi=doi,
        url=url,
        tipo=tipo,
        source="html",
    )


def fetch_html(url: str) -> Optional[Meta]:
    try:
        code, raw, ctype = _http_get(url, accept="text/html,application/xhtml+xml")
    except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError, ValueError):
        return None
    if code != 200:
        return None
    if "pdf" in (ctype or "").lower():
        return None
    try:
        html = raw.decode("utf-8", errors="replace")
    except Exception:
        html = raw.decode("latin-1", errors="replace")
    meta = parse_html_meta(html, url)
    if meta.doi and (not meta.title or not meta.authors):
        cr = fetch_crossref(meta.doi)
        if cr:
            # prefer Crossref bibliographic quality
            return cr
    if not meta.title and not meta.doi:
        return None
    return meta


def fetch_pdf_local(path: Path) -> Optional[Meta]:
    if not path.exists() or path.suffix.lower() != ".pdf":
        return None
    try:
        import fitz  # pymupdf
    except ImportError:
        return None
    try:
        doc = fitz.open(path)
    except Exception:
        return None
    try:
        info = doc.metadata or {}
        title = (info.get("title") or "").strip()
        author = (info.get("author") or "").strip()
        # scan first pages for DOI
        blob = ""
        for i in range(min(3, doc.page_count)):
            try:
                blob += doc.load_page(i).get_text("text") + "\n"
            except Exception:
                continue
        doi = extrair_doi(blob)
        if doi:
            cr = fetch_crossref(doi)
            if cr:
                return cr
        authors = []
        if author:
            # "A. Smith; B. Jones" or "Smith, A."
            for chunk in re.split(r"[;|/]| and ", author):
                chunk = chunk.strip()
                if not chunk:
                    continue
                if "," in chunk:
                    authors.append(chunk)
                else:
                    parts = chunk.split()
                    if len(parts) >= 2:
                        authors.append(
                            f"{parts[-1]}, "
                            + " ".join(p[0].upper() + "." for p in parts[:-1])
                        )
                    else:
                        authors.append(chunk)
        year = ""
        m = _RX_YEAR.search(blob[:2000]) or _RX_YEAR.search(title)
        if m:
            year = m.group(0)
        if not title and not authors and not doi:
            return None
        return Meta(
            authors=authors,
            year=year,
            title=title or path.stem.replace("_", " "),
            doi=doi,
            tipo="pdf",
            source="pdf",
            notes="pdf-metadata",
        )
    finally:
        doc.close()


def fallback_from_filename(caminho: str | None, url: str | None = None) -> Meta:
    """Last resort — parse Author_Year_Title stem (still better than raw stem)."""
    raw = caminho or url or ""
    stem = Path(str(raw).replace("\\", "/").split("?")[0]).stem
    stem = urllib.parse.unquote(stem)
    rotulo = stem.replace("_", " ").replace("-", " ").strip()
    year = ""
    m = _RX_YEAR.search(rotulo)
    if m:
        year = m.group(0)
    authors: list[str] = []
    title = rotulo
    # Pattern: Surname Year Title…
    m2 = re.match(
        r"^([A-ZÁÉÍÓÚÄËÏÖÜ][\w'’\-]+)(?:\s+([A-Z][\w'’\-]+))?\s+(\d{4})\s+(.+)$",
        rotulo,
    )
    if m2:
        a1, a2, year, title = m2.group(1), m2.group(2), m2.group(3), m2.group(4)
        if a2 and a2.lower() not in {"and", "e"}:
            authors = [f"{a1}, {a2[0]}."]
        else:
            authors = [f"{a1}, A."]
    return Meta(
        authors=authors,
        year=year,
        title=title,
        url=(url or "") if url and "://" in url else "",
        tipo="pdf",
        source="filename",
        notes="fallback-filename",
    )


def resolver_meta(url: str | None, caminho: str | None) -> Meta:
    """Follow hyperlink / local file → bibliographic Meta."""
    url = (url or "").strip() or None
    caminho = (caminho or "").strip() or None
    if url and url.lower() in {"nan", "none", "nat"}:
        url = None
    if caminho and caminho.lower() in {"nan", "none", "nat"}:
        caminho = None

    doi = extrair_doi(url) or ""
    if not doi and url:
        # sometimes path-like DOI
        doi = extrair_doi(urllib.parse.unquote(url))

    if doi:
        cr = fetch_crossref(doi)
        if cr:
            return cr

    if url and (url.startswith("http://") or url.startswith("https://")):
        # doi.org already tried via doi extract; still try HTML for publisher pages
        html_meta = fetch_html(url)
        if html_meta:
            if html_meta.doi and html_meta.source == "html":
                cr = fetch_crossref(html_meta.doi)
                if cr:
                    return cr
            return html_meta

    if caminho:
        p = Path(caminho)
        if p.suffix.lower() == ".pdf":
            pdf_meta = fetch_pdf_local(p)
            if pdf_meta:
                return pdf_meta

    return fallback_from_filename(caminho, url)


# ---------------------------------------------------------------------------
# Excel I/O
# ---------------------------------------------------------------------------

def obras_unicas(df: pd.DataFrame) -> pd.DataFrame:
    """One row per underlying work (doc_id / path / url)."""
    rows = []
    seen: set[str] = set()
    for _, row in df.iterrows():
        caminho = ta._primeira_col(row, ("caminho_ficheiro", "caminho", "doc_id"))
        url = None
        for c in ("url", "URL", "link"):
            if c in row.index and pd.notna(row[c]) and str(row[c]).strip():
                url = str(row[c]).strip()
                break
        chave = ta._norm_chave(str(caminho or url or ""))
        if not chave or chave in seen:
            continue
        seen.add(chave)
        rows.append({
            "chave": chave,
            "doc_id": row["doc_id"] if "doc_id" in row.index else "",
            "caminho_ficheiro": caminho or "",
            "url": url or "",
            "fonte_apa_actual": ta._primeira_col(row, ta.COLS_FONTE_DIRECTA) or "",
        })
    return pd.DataFrame(rows)


def enriquecer_obras(
    obras: pd.DataFrame,
    *,
    pause_s: float = 0.15,
    limite: int | None = None,
) -> pd.DataFrame:
    out_rows = []
    n = len(obras) if limite is None else min(limite, len(obras))
    for i, row in obras.head(n).iterrows():
        print(
            f"  [{len(out_rows)+1}/{n}] {row['chave'][:60]} …",
            flush=True,
        )
        meta = resolver_meta(
            str(row["url"]) if row["url"] else None,
            str(row["caminho_ficheiro"]) if row["caminho_ficheiro"] else None,
        )
        apa = format_apa7(meta)
        out_rows.append({
            "underlying_file": Path(str(row["caminho_ficheiro"])).name
            if row["caminho_ficheiro"] else row["chave"],
            "ficheiro": Path(str(row["caminho_ficheiro"])).name
            if row["caminho_ficheiro"] else "",
            "caminho_ficheiro": row["caminho_ficheiro"],
            "doc_id": row["doc_id"],
            "url": row["url"] or meta.url,
            "doi": meta.doi,
            "apa7": apa,
            "fonte_apa": apa,
            "meta_source": meta.source,
            "meta_tipo": meta.tipo,
            "meta_title": meta.title,
            "meta_year": meta.year,
            "fonte_apa_actual": row.get("fonte_apa_actual", ""),
            "notes": meta.notes,
        })
        if pause_s:
            time.sleep(pause_s)
    return pd.DataFrame(out_rows)


def escrever_catalogo(df: pd.DataFrame, saida: Path) -> Path:
    saida.parent.mkdir(parents=True, exist_ok=True)
    if saida.suffix.lower() in {".xlsx", ".xlsm"}:
        df.to_excel(saida, index=False)
    else:
        df.to_csv(saida, index=False, encoding="utf-8-sig")
    return saida


def _escrever_fonte_apa_no_xlsx(
    xlsx: Path,
    catalogo: pd.DataFrame,
    *,
    folha: str | None = None,
) -> Path:
    chave_apa: dict[str, str] = {}
    for _, r in catalogo.iterrows():
        apa = str(r.get("apa7") or r.get("fonte_apa") or "").strip()
        if not apa:
            continue
        for col in ("underlying_file", "ficheiro", "caminho_ficheiro", "doc_id", "url"):
            if col in catalogo.columns and pd.notna(r[col]) and str(r[col]).strip():
                chave_apa[ta._norm_chave(str(r[col]))] = apa

    df = ta.ler_concordancia(xlsx, folha=folha)
    if "fonte_apa" not in df.columns:
        df["fonte_apa"] = ""

    n_fill = 0
    for i, row in df.iterrows():
        actual = str(row.get("fonte_apa") or "").strip()
        # overwrite filename-like / empty / (n.d.) stubs
        stub = (
            not actual
            or actual.endswith("(n.d.).")
            or re.match(r"^[A-Za-z0-9].*\(n\.d\.\)\.$", actual) is not None
        )
        if actual and not stub:
            continue
        hit = None
        for c in ("caminho_ficheiro", "caminho", "doc_id", "url"):
            if c not in row.index or pd.isna(row[c]):
                continue
            k = ta._norm_chave(str(row[c]))
            if k in chave_apa:
                hit = chave_apa[k]
                break
        if hit:
            df.at[i, "fonte_apa"] = hit
            n_fill += 1

    bak = xlsx.with_name(xlsx.stem + xlsx.suffix + ".bak-apa7")
    if not bak.exists():
        import shutil
        shutil.copy2(xlsx, bak)

    # preserve other sheets from backup/original
    with pd.ExcelFile(bak if bak.exists() else xlsx) as xl:
        nomes = list(xl.sheet_names)
        outros = {
            n: pd.read_excel(xl, sheet_name=n)
            for n in nomes
            if n != (folha or "8_Concordancia")
        }
    alvo = folha or "8_Concordancia"
    with pd.ExcelWriter(xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=alvo, index=False)
        for n, frame in outros.items():
            frame.to_excel(writer, sheet_name=n, index=False)
    print(f"  fonte_apa actualizada em {n_fill} linhas")
    print(f"  backup: {bak.name}")
    return xlsx


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv: list[str] | None = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("--xlsx", type=Path, required=True,
                    help="Excel com folha 8_Concordancia (fase 1/2)")
    ap.add_argument("--folha", default=None,
                    help="folha (omissão: 8_Concordancia)")
    ap.add_argument("--saida", type=Path, default=None,
                    help="catálogo APA7 (omissão: <xlsx>_refs_apa7.xlsx)")
    ap.add_argument("--escrever-xlsx", action="store_true",
                    help="escrever coluna fonte_apa no Excel de entrada")
    ap.add_argument("--limite", type=int, default=None,
                    help="processar só as primeiras N obras (teste)")
    ap.add_argument("--pause", type=float, default=0.15,
                    help="pausa entre pedidos HTTP (s)")
    ap.add_argument("--incluir-nao-nucleares", action="store_true",
                    help="não filtrar nuclear/genuína ao listar obras")
    args = ap.parse_args(argv)

    if not args.xlsx.exists():
        print(f"Ficheiro não encontrado: {args.xlsx}", file=sys.stderr)
        return 2

    print(f"APA7 · hiperligações ← {args.xlsx.name}", flush=True)
    df = ta.ler_concordancia(args.xlsx, folha=args.folha)
    df = ta.filtrar_genuinas(
        df, incluir_nao_nucleares=args.incluir_nao_nucleares
    )
    obras = obras_unicas(df)
    print(f"  obras únicas: {len(obras)}", flush=True)
    if obras.empty:
        print("Nada a fazer (sem caminho/url).", file=sys.stderr)
        return 1

    cat = enriquecer_obras(obras, pause_s=args.pause, limite=args.limite)
    saida = args.saida or args.xlsx.with_name(args.xlsx.stem + "_refs_apa7.xlsx")
    escrever_catalogo(cat, saida)
    n_cr = int((cat["meta_source"] == "crossref").sum())
    n_html = int((cat["meta_source"] == "html").sum())
    n_pdf = int((cat["meta_source"] == "pdf").sum())
    n_fn = int((cat["meta_source"] == "filename").sum())
    print(
        f"  catálogo → {saida.name} "
        f"(crossref={n_cr} html={n_html} pdf={n_pdf} filename={n_fn})",
        flush=True,
    )

    if args.escrever_xlsx:
        _escrever_fonte_apa_no_xlsx(args.xlsx, cat, folha=args.folha)

    print(
        "\n=== APA7 concluído ===\n"
        f"  Catálogo: {saida}\n"
        f"  Fase 3:   python textura_apendice.py --xlsx \"{args.xlsx}\" "
        f"--refs \"{saida}\""
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
