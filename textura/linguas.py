#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Registo de línguas (EN/PT/FR/DE) — Phase 3.

A língua da execução é escolhida ao nível da corrida (``--lingua``), não por
detecção automática nem por linha da matriz. O modelo spaCy, as preposições
genitivas e o estado de validação vêm deste registo.

A omissão da corrida é ``todas`` (união dos paradigmas NOS). EN continua
o âncora de modelo spaCy e preposições quando ``--lingua todas``. Os
goldens EN/PT/FR fixam ``--lingua`` explicitamente. DE permanece
«não validado» (ausente do corpus adjudicado).
"""

from __future__ import annotations

LINGUA_OMISSAO = "todas"

from dataclasses import dataclass
from functools import lru_cache
from typing import Mapping

from textura.lexico import COPULAS, NEGACAO, NOS
from textura.tokenizacao import sem_diacriticos


# Preposições usadas por ``_tem_complemento_genitivo`` antes da Phase 3
# (mistura EN+PT já presente no código — não alterar o conjunto EN).
_PREPS_GENITIVO_EN = frozenset({"of", "de", "in", "on"})
_PREPS_ASSOC_EN = frozenset({"with"})


@dataclass(frozen=True)
class LinguaConfig:
    codigo: str
    modelo_spacy: str
    preps_genitivo: frozenset[str]
    preps_associativa: frozenset[str]
    status: str  # "validado" | "nao_validado"
    nota: str = ""

    @property
    def nos(self) -> frozenset[str]:
        return frozenset(NOS.get(self.codigo, frozenset()))

    @property
    def copulas(self) -> frozenset[str]:
        # EN: léxico partilhado. Outras línguas: união com formas locais
        # (não substitui o TSV global usado pela heurística).
        extra = _COPULAS_EXTRA.get(self.codigo, frozenset())
        if self.codigo == "en":
            return frozenset(COPULAS)
        return frozenset(COPULAS) | extra

    @property
    def negadores(self) -> frozenset[str]:
        extra = _NEGADORES_EXTRA.get(self.codigo, frozenset())
        if self.codigo == "en":
            return frozenset(NEGACAO)
        return frozenset(NEGACAO) | extra


_COPULAS_EXTRA: Mapping[str, frozenset[str]] = {
    "pt": frozenset({
        "é", "são", "era", "eram", "foi", "foram", "ser", "sendo", "sido",
        "está", "estão", "esteve", "estava", "ficar", "fica", "parece",
        "permanece", "torna", "tornou",
    }),
    "fr": frozenset({
        "est", "sont", "était", "étaient", "être", "semble", "semblent",
        "devient", "reste", "demeurent",
    }),
    "de": frozenset({
        "ist", "sind", "war", "waren", "sein", "scheint", "bleibt", "wird",
    }),
}

_NEGADORES_EXTRA: Mapping[str, frozenset[str]] = {
    "pt": frozenset({
        "não", "nao", "nunca", "jamais", "sem", "nem",
    }),
    "fr": frozenset({
        "ne", "pas", "jamais", "sans", "non", "point",
    }),
    "de": frozenset({
        "nicht", "nie", "niemals", "ohne", "kein", "keine",
    }),
}


REGISTO: dict[str, LinguaConfig] = {
    "en": LinguaConfig(
        codigo="en",
        modelo_spacy="en_core_web_sm",
        preps_genitivo=_PREPS_GENITIVO_EN,
        preps_associativa=_PREPS_ASSOC_EN,
        status="validado",
        nota="Âncora golden; não alterar sem actualizar testes Phase 0.",
    ),
    "pt": LinguaConfig(
        codigo="pt",
        modelo_spacy="pt_core_news_sm",
        preps_genitivo=frozenset({"de", "do", "da", "dos", "das", "em", "no",
                                  "na", "nos", "nas"}),
        preps_associativa=frozenset({"com"}),
        status="validado",
        nota="Classificação dependencial com pt_core_news_sm.",
    ),
    "fr": LinguaConfig(
        codigo="fr",
        modelo_spacy="fr_core_news_sm",
        preps_genitivo=frozenset({"de", "du", "des", "d'", "en", "dans"}),
        preps_associativa=frozenset({"avec"}),
        status="validado",
        nota="Golden fino (atributiva / genitiva / coordenação) com "
             "fr_core_news_sm-3.8.0; corpus contém atestações FR.",
    ),
    "de": LinguaConfig(
        codigo="de",
        modelo_spacy="de_core_news_sm",
        # Caso genitivo alemão: dependências spaCy diferem; ver teste xfail.
        preps_genitivo=frozenset({"von", "vom", "in", "im", "auf"}),
        preps_associativa=frozenset({"mit"}),
        status="nao_validado",
        nota=(
            "Ausente do corpus adjudicado; genitivo morfológico sem prep "
            "ainda não tratado (TODO; teste xfail). Pode permanecer "
            "nao_validado indefinidamente."
        ),
    ),
}

CODIGOS = tuple(REGISTO.keys())
DEFAULT = REGISTO["en"]


@lru_cache(maxsize=1)
def _nos_dobrados() -> dict[str, frozenset[str]]:
    return {
        cod: frozenset(sem_diacriticos(f) for f in forms)
        for cod, forms in NOS.items()
    }


def lingua_do_no(forma: str) -> str:
    """Língua do modelo spaCy a partir da forma do nó (não do contexto).

    Formas exclusivas de um paradigma NOS ganham (``textura`` → pt).
    Formas partilhadas (``texture``, ``textures``, ``textural``) → ``en``.
    Inventário só FR: ``--lingua fr``.
    """
    w = sem_diacriticos(forma)
    if not w:
        return "en"
    donos = [cod for cod, forms in _nos_dobrados().items() if w in forms]
    if len(donos) == 1:
        return donos[0]
    if not donos and w.startswith("textura"):
        return "pt"
    return "en"


def obter(codigo: str) -> LinguaConfig:
    key = (codigo or "en").strip().lower()
    if key not in REGISTO:
        raise KeyError(
            f"Língua desconhecida {codigo!r}; "
            f"registadas: {', '.join(CODIGOS)}"
        )
    return REGISTO[key]


@dataclass(frozen=True)
class ExecucaoLinguistica:
    """Resolução ao nível da corrida (nunca por linha da matriz)."""

    lingua_cli: str
    cfg: LinguaConfig
    modelo_spacy: str
    aviso: str = ""


def resolver_execucao(
    lingua: str,
    modelo_cli: str | None = None,
) -> ExecucaoLinguistica:
    """Mapeia ``--lingua`` / ``--modelo`` → config de uma só língua efectiva.

    ``--lingua todas`` une paradigmas NOS. A classificação spaCy é
    despachada pela forma do nó (``lingua_do_no``): ``textura*`` → PT,
    ``texture*`` exclusivos EN → EN; formas partilhadas (``textural``)
    ficam EN. Não há detecção de língua pelo texto da janela.
    """
    lingua = (lingua or LINGUA_OMISSAO).strip().lower()
    avisos: list[str] = []

    if lingua == "todas":
        cfg = DEFAULT
        avisos.append(
            "lingua=todas: paradigmas NOS unidos; spaCy despachado "
            "pela forma do no (textura*→pt, texture*→en; "
            "textural/texture partilhados→en). "
            "sem detecção de língua pelo texto da janela."
        )
    else:
        cfg = obter(lingua)

    modelo = (modelo_cli or "").strip() or cfg.modelo_spacy
    if cfg.status == "nao_validado":
        avisos.append(
            f"AVISO: língua '{cfg.codigo}' registada como não validado "
            f"({cfg.nota})"
        )

    return ExecucaoLinguistica(
        lingua_cli=lingua,
        cfg=cfg,
        modelo_spacy=modelo,
        aviso=" ".join(avisos),
    )
