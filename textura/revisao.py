#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Vocabulário canónico de ``revisao_sugerida`` (Phase 5).

Todos os emissores do pipeline devem construir etiquetas via as funções
deste módulo. Formas prefixadas levam um detalhe após ``:``; o teste de
adesão valida cada segmento separado por ``;``.
"""

from __future__ import annotations

import re

# Etiquetas exactas (sem sufixo).
REVISAO_EXACTAS: frozenset[str] = frozenset({
    "genitiva_por_complemento",
    "atributiva_via_conj",
    "atributiva_coordenada",
})

# Prefixos (incluem o ':' final). O detalhe é livre mas obrigatório.
REVISAO_PREFIXOS: frozenset[str] = frozenset({
    "coordenacao_heterogenea:",
    "associativa_com_nao_textural:",
    "dominio_janela:",
})

# União legível para documentação / gerador.
REVISAO_VOCABULARIO: frozenset[str] = REVISAO_EXACTAS | REVISAO_PREFIXOS

_RX_SEP = re.compile(r"\s*;\s*")


class RevisaoError(ValueError):
    """Etiqueta de revisão fora do vocabulário canónico."""


def etiqueta_exacta(nome: str) -> str:
    if nome not in REVISAO_EXACTAS:
        raise RevisaoError(
            f"revisao_sugerida exacta desconhecida: {nome!r}; "
            f"válidas: {sorted(REVISAO_EXACTAS)}"
        )
    return nome


def etiqueta_prefixada(prefixo: str, detalhe: str) -> str:
    """``prefixo`` com ou sem ':' final; ``detalhe`` não vazio."""
    p = prefixo if prefixo.endswith(":") else f"{prefixo}:"
    if p not in REVISAO_PREFIXOS:
        raise RevisaoError(
            f"prefixo de revisao_sugerida desconhecido: {p!r}; "
            f"válidos: {sorted(REVISAO_PREFIXOS)}"
        )
    d = str(detalhe or "").strip()
    if not d:
        raise RevisaoError(f"detalhe vazio para prefixo {p!r}")
    return f"{p}{d}"


def juntar_etiquetas(*partes: str) -> str:
    """Concatena etiquetas já canónicas, omitindo vazias."""
    limpas = [p.strip() for p in partes if p and str(p).strip()]
    return "; ".join(limpas)


def segmentos(blob: str) -> list[str]:
    if blob is None or (isinstance(blob, float) and blob != blob):
        return []
    s = str(blob).strip()
    if not s:
        return []
    return [p for p in _RX_SEP.split(s) if p]


def e_etiqueta_valida(tag: str) -> bool:
    t = (tag or "").strip()
    if not t:
        return False
    if t in REVISAO_EXACTAS:
        return True
    for pref in REVISAO_PREFIXOS:
        if t.startswith(pref) and len(t) > len(pref):
            return True
    return False


def validar_revisao_sugerida(blob: str) -> list[str]:
    """Devolve lista de segmentos inválidos (vazia = OK)."""
    return [seg for seg in segmentos(blob) if not e_etiqueta_valida(seg)]
