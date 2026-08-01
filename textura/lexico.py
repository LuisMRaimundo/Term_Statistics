#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fonte única de léxicos (dados/lexicos/*.tsv).

Listas e taxonomias partilhadas entre near, query, triagem, QA e estatística
carregam-se aqui. Não duplicar literais noutros módulos — importar deste.
"""

from __future__ import annotations

import re
from functools import lru_cache
from pathlib import Path
from typing import Dict, FrozenSet, Mapping, Tuple

_DIR_LEXICOS = Path(__file__).resolve().parents[1] / "dados" / "lexicos"
_RE_HIFEN_QUEBRA = re.compile(r"(\w)-\s+(\w)")


def dir_lexicos() -> Path:
    return _DIR_LEXICOS


def caminho_dominios_path() -> Path:
    """TSV de regras path (canónico em dados/lexicos/; fallback raiz)."""
    preferido = _DIR_LEXICOS / "dominios_path.tsv"
    if preferido.is_file():
        return preferido
    return Path(__file__).resolve().parents[1] / "dominios.tsv"


def _ler_linhas(path: Path) -> list[str]:
    if not path.is_file():
        raise FileNotFoundError(f"Léxico em falta: {path}")
    out: list[str] = []
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        out.append(line)
    return out


def _load_frozenset(nome: str, *, header: str | None = "token") -> FrozenSet[str]:
    linhas = _ler_linhas(_DIR_LEXICOS / nome)
    if header and linhas and linhas[0].split("\t")[0] == header:
        linhas = linhas[1:]
    return frozenset(linhas)


@lru_cache(maxsize=1)
def carregar_nos() -> Dict[str, FrozenSet[str]]:
    linhas = _ler_linhas(_DIR_LEXICOS / "nos.tsv")
    if linhas and linhas[0].startswith("lingua"):
        linhas = linhas[1:]
    acc: Dict[str, set[str]] = {}
    for line in linhas:
        parts = line.split("\t")
        if len(parts) != 2:
            continue
        lang, forma = parts[0].strip().lower(), parts[1].strip().lower()
        acc.setdefault(lang, set()).add(forma)
    return {k: frozenset(v) for k, v in acc.items()}


@lru_cache(maxsize=1)
def carregar_dominio_janela() -> Dict[str, Tuple[str, ...]]:
    """Ordem de domínios = ordem no TSV (prioridade de classificação)."""
    linhas = _ler_linhas(_DIR_LEXICOS / "dominio_janela.tsv")
    if linhas and linhas[0].startswith("dominio"):
        linhas = linhas[1:]
    ordem: list[str] = []
    acc: Dict[str, list[str]] = {}
    for line in linhas:
        parts = line.split("\t")
        if len(parts) != 2:
            continue
        dom, pista = parts[0].strip(), parts[1].strip().lower()
        if dom not in acc:
            ordem.append(dom)
            acc[dom] = []
        if pista not in acc[dom]:
            acc[dom].append(pista)
    return {d: tuple(acc[d]) for d in ordem}


@lru_cache(maxsize=1)
def carregar_falsos_amigos() -> Dict[str, str]:
    linhas = _ler_linhas(_DIR_LEXICOS / "falsos_amigos.tsv")
    if linhas and linhas[0].startswith("forma"):
        linhas = linhas[1:]
    out: Dict[str, str] = {}
    for line in linhas:
        parts = line.split("\t", 1)
        if len(parts) != 2:
            continue
        out[parts[0].strip().lower()] = parts[1].strip()
    return out


@lru_cache(maxsize=1)
def carregar_dominio_taxonomia() -> Dict[str, str]:
    """dominio -> fonte (path|janela|ambos|sistema)."""
    linhas = _ler_linhas(_DIR_LEXICOS / "dominio_taxonomia.tsv")
    if linhas and linhas[0].startswith("dominio"):
        linhas = linhas[1:]
    out: Dict[str, str] = {}
    for line in linhas:
        parts = line.split("\t")
        if len(parts) != 2:
            continue
        out[parts[0].strip()] = parts[1].strip()
    return out


@lru_cache(maxsize=1)
def carregar_dominios_validos_path() -> FrozenSet[str]:
    tax = carregar_dominio_taxonomia()
    return frozenset(
        d for d, fonte in tax.items() if fonte in {"path", "ambos"}
    )


def _normaliza_janela(texto: str) -> str:
    texto = str(texto or "").replace("\n", " ").replace("\r", " ")
    texto = _RE_HIFEN_QUEBRA.sub(r"\1\2", texto)
    return re.sub(r"\s+", " ", texto).strip()


# --- API estável -----------------------------------------------------------

NOS: Mapping[str, FrozenSet[str]] = carregar_nos()
NEGACAO: FrozenSet[str] = _load_frozenset("negacao.tsv")
GRADUACAO: FrozenSet[str] = _load_frozenset("graduacao.tsv")
MODALIDADE: FrozenSet[str] = _load_frozenset("modalidade.tsv")
COPULAS: FrozenSet[str] = _load_frozenset("copulas.tsv")
ABREVIATURAS: FrozenSet[str] = _load_frozenset("abreviaturas.tsv")
POLO_ESTABILIDADE: FrozenSet[str] = _load_frozenset("polo_estabilidade.tsv")
POLO_VARIABILIDADE: FrozenSet[str] = _load_frozenset("polo_variabilidade.tsv")
RELACOES_NUCLEARES: FrozenSet[str] = _load_frozenset("relacoes_nucleares.tsv")
RELACOES_NAO_NUCLEARES: FrozenSet[str] = _load_frozenset(
    "relacoes_nao_nucleares.tsv"
)
DOMINIO_JANELA_LEXICO: Mapping[str, Tuple[str, ...]] = carregar_dominio_janela()
DOMAIN_LEXICON: Mapping[str, Tuple[str, ...]] = DOMINIO_JANELA_LEXICO
FALSOS_AMIGOS_FORMAS: Mapping[str, str] = carregar_falsos_amigos()
DOMINIOS_VALIDOS: FrozenSet[str] = carregar_dominios_validos_path()


def dominio_janela(contexto: str) -> str:
    """Domínio extramusical sugerido pela própria janela ('' se nenhum)."""
    c = _normaliza_janela(contexto).lower()
    for dom, pistas in DOMINIO_JANELA_LEXICO.items():
        for p in pistas:
            if p in c:
                return dom
    return ""
