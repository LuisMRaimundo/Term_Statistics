#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fonte única de léxicos (dados/lexicos/*.tsv).

Listas e taxonomias partilhadas entre near, query, triagem, QA e estatística
carregam-se aqui. Não duplicar literais noutros módulos — importar deste.

Caminhos ancoram-se à raiz do projecto (pasta-mãe de ``textura/``), nunca ao CWD.
"""

from __future__ import annotations

import re
import warnings
from functools import lru_cache
from pathlib import Path
from typing import Dict, FrozenSet, Mapping, Tuple

_RAIZ_PROJECTO = Path(__file__).resolve().parents[1]
_DIR_LEXICOS = _RAIZ_PROJECTO / "dados" / "lexicos"
_DIR_OVERRIDE: Path | None = None
_RE_HIFEN_QUEBRA = re.compile(r"(\w)-\s+(\w)")


class LexicoError(ValueError):
    """Falha ao ler ou validar um ficheiro em dados/lexicos/."""


def raiz_projecto() -> Path:
    return _RAIZ_PROJECTO


def dir_lexicos() -> Path:
    return _DIR_OVERRIDE if _DIR_OVERRIDE is not None else _DIR_LEXICOS


def _limpar_caches() -> None:
    carregar_nos.cache_clear()
    carregar_dominio_janela.cache_clear()
    carregar_falsos_amigos.cache_clear()
    carregar_dominio_taxonomia.cache_clear()
    carregar_dominios_validos_path.cache_clear()
    _load_frozenset.cache_clear()


def definir_dir_lexicos_para_teste(path: Path | None) -> None:
    """Só para testes: aponta o loader a uma cópia temporária dos TSV."""
    global _DIR_OVERRIDE
    _DIR_OVERRIDE = Path(path) if path is not None else None
    _limpar_caches()


def _linhas_regras_dominio(path: Path) -> frozenset[str]:
    """Conjunto normalizado «padrão\\tdomínio» (ignora comentários/vazios)."""
    if not path.is_file():
        return frozenset()
    out_set: set[str] = set()
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.split("#", 1)[0].strip()
        if not line or "\t" not in line:
            continue
        pad, dom = line.split("\t", 1)
        pad, dom = pad.strip(), dom.strip().lower()
        if pad and dom:
            out_set.add(f"{pad}\t{dom}")
    return frozenset(out_set)


def caminho_dominios_path(
    *,
    avisar: bool = True,
    raiz: Path | None = None,
    canon: Path | None = None,
) -> Path:
    """Resolve o TSV path→domínio com precedência explícita.

    1. Só raiz ``dominios.tsv`` (com regras) → usa raiz + DeprecationWarning.
    2. Só ``dados/lexicos/dominios_path.tsv`` → usa canónico.
    3. Ambos com o mesmo conteúdo normativo → canónico; avisa para remover a raiz.
    4. Ambos com conteúdo diferente → ``LexicoError`` (migração consciente).
    5. Nenhum com regras → devolve o canónico (o chamador trata ficheiro em falta).
    """
    raiz = Path(raiz) if raiz is not None else _RAIZ_PROJECTO / "dominios.tsv"
    canon = (
        Path(canon) if canon is not None else dir_lexicos() / "dominios_path.tsv"
    )
    regras_raiz = _linhas_regras_dominio(raiz)
    regras_canon = _linhas_regras_dominio(canon)

    if regras_raiz and regras_canon:
        if regras_raiz != regras_canon:
            raise LexicoError(
                f"Conflito em regras path→domínio: {raiz} e {canon} "
                f"diferem materialmente ({len(regras_raiz)} vs "
                f"{len(regras_canon)} regras normalizadas). "
                f"Migre o conteúdo desejado para {canon} e remova ou "
                f"alinhe o ficheiro na raiz — não há escolha silenciosa."
            )
        if avisar:
            warnings.warn(
                f"{raiz} é um duplicado obsoleto de {canon}; "
                f"remova o ficheiro na raiz (o pipeline usa o canónico).",
                DeprecationWarning,
                stacklevel=2,
            )
        return canon

    if regras_raiz:
        if avisar:
            warnings.warn(
                f"A usar {raiz} (legado). Migre as regras para {canon} "
                f"e remova o ficheiro na raiz.",
                DeprecationWarning,
                stacklevel=2,
            )
        return raiz

    return canon


def _ler_linhas(path: Path) -> list[str]:
    if not path.is_file():
        raise LexicoError(f"Léxico em falta: {path}")
    try:
        texto = path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError as exc:
        raise LexicoError(
            f"Léxico com encoding inválido (esperado UTF-8): {path}"
        ) from exc
    out: list[str] = []
    for raw in texto.splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        out.append(line)
    if not out:
        raise LexicoError(
            f"Léxico vazio ou só com comentários (sem entradas): {path}"
        )
    return out


@lru_cache(maxsize=32)
def _load_frozenset(nome: str, header: str = "token") -> FrozenSet[str]:
    path = dir_lexicos() / nome
    linhas = _ler_linhas(path)
    if header and linhas and linhas[0].split("\t")[0] == header:
        linhas = linhas[1:]
    if not linhas:
        raise LexicoError(
            f"Léxico sem entradas após o cabeçalho '{header}': {path}"
        )
    tokens: list[str] = []
    for i, line in enumerate(linhas, start=1):
        if "\t" in line:
            raise LexicoError(
                f"Léxico {path}: linha de dados {i} tem tabulador "
                f"(esperado uma coluna '{header}'): {line!r}"
            )
        tokens.append(line)
    return frozenset(tokens)


@lru_cache(maxsize=1)
def carregar_nos() -> Dict[str, FrozenSet[str]]:
    path = dir_lexicos() / "nos.tsv"
    linhas = _ler_linhas(path)
    if linhas[0].startswith("lingua"):
        linhas = linhas[1:]
    if not linhas:
        raise LexicoError(f"Léxico sem entradas após o cabeçalho: {path}")
    acc: Dict[str, set[str]] = {}
    for i, line in enumerate(linhas, start=1):
        parts = line.split("\t")
        if len(parts) != 2:
            raise LexicoError(
                f"Léxico {path}: linha {i} deve ter 2 colunas "
                f"(lingua, forma); obtido {len(parts)}: {line!r}"
            )
        lang, forma = parts[0].strip().lower(), parts[1].strip().lower()
        if not lang or not forma:
            raise LexicoError(
                f"Léxico {path}: linha {i} com lingua/forma vazia: {line!r}"
            )
        acc.setdefault(lang, set()).add(forma)
    return {k: frozenset(v) for k, v in acc.items()}


@lru_cache(maxsize=1)
def carregar_dominio_janela() -> Dict[str, Tuple[str, ...]]:
    """Ordem de domínios = ordem no TSV (prioridade de classificação)."""
    path = dir_lexicos() / "dominio_janela.tsv"
    linhas = _ler_linhas(path)
    if linhas[0].startswith("dominio"):
        linhas = linhas[1:]
    if not linhas:
        raise LexicoError(f"Léxico sem entradas após o cabeçalho: {path}")
    ordem: list[str] = []
    acc: Dict[str, list[str]] = {}
    for i, line in enumerate(linhas, start=1):
        parts = line.split("\t")
        if len(parts) != 2:
            raise LexicoError(
                f"Léxico {path}: linha {i} deve ter 2 colunas "
                f"(dominio, pista); obtido {len(parts)}: {line!r}"
            )
        dom, pista = parts[0].strip(), parts[1].strip().lower()
        if not dom or not pista:
            raise LexicoError(
                f"Léxico {path}: linha {i} com dominio/pista vazia: {line!r}"
            )
        if dom not in acc:
            ordem.append(dom)
            acc[dom] = []
        if pista not in acc[dom]:
            acc[dom].append(pista)
    return {d: tuple(acc[d]) for d in ordem}


@lru_cache(maxsize=1)
def carregar_falsos_amigos() -> Dict[str, str]:
    path = dir_lexicos() / "falsos_amigos.tsv"
    linhas = _ler_linhas(path)
    if linhas[0].startswith("forma"):
        linhas = linhas[1:]
    if not linhas:
        raise LexicoError(f"Léxico sem entradas após o cabeçalho: {path}")
    out: Dict[str, str] = {}
    for i, line in enumerate(linhas, start=1):
        parts = line.split("\t", 1)
        if len(parts) != 2:
            raise LexicoError(
                f"Léxico {path}: linha {i} deve ter 2 colunas "
                f"(forma, motivo_exclusao): {line!r}"
            )
        forma, motivo = parts[0].strip().lower(), parts[1].strip()
        if not forma or not motivo:
            raise LexicoError(
                f"Léxico {path}: linha {i} com forma/motivo vazio: {line!r}"
            )
        out[forma] = motivo
    return out


@lru_cache(maxsize=1)
def carregar_dominio_taxonomia() -> Dict[str, str]:
    """dominio -> fonte (path|janela|ambos|sistema)."""
    path = dir_lexicos() / "dominio_taxonomia.tsv"
    linhas = _ler_linhas(path)
    if linhas[0].startswith("dominio"):
        linhas = linhas[1:]
    if not linhas:
        raise LexicoError(f"Léxico sem entradas após o cabeçalho: {path}")
    out: Dict[str, str] = {}
    fontes_ok = {"path", "janela", "ambos", "sistema"}
    for i, line in enumerate(linhas, start=1):
        parts = line.split("\t")
        if len(parts) != 2:
            raise LexicoError(
                f"Léxico {path}: linha {i} deve ter 2 colunas "
                f"(dominio, fonte): {line!r}"
            )
        dom, fonte = parts[0].strip(), parts[1].strip()
        if not dom or fonte not in fontes_ok:
            raise LexicoError(
                f"Léxico {path}: linha {i} dominio/fonte inválidos: {line!r}"
            )
        out[dom] = fonte
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
    for dom, pistas in carregar_dominio_janela().items():
        for p in pistas:
            if p in c:
                return dom
    return ""
