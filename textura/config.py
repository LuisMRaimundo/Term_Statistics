#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Constantes e padrões partilhados da extracção NEAR (Phase 1 move)."""

from __future__ import annotations

import re

from textura.lexico import (
    ABREVIATURAS,
    COPULAS,
    GRADUACAO,
    MODALIDADE,
    NEGACAO,
    NOS,
    RELACOES_NUCLEARES,
)

MODALIZACAO = GRADUACAO  # alias de leitura legado


COLUNAS = ["L5", "L4", "L3", "L2", "L1", "NODE",
           "R1", "R2", "R3", "R4", "R5",
           "caminho", "url", "raiz", "contexto", "n"]

# Campo lexical. Chave = etiqueta do tipo lexical; valor = truncaturas.
# O asterisco final funciona como truncatura à direita (prefixo).
CAMPO = {
    # --- núcleo da invariância -------------------------------------------
    "uniform":      ["uniform*"],
    "invariable":   ["invariab*", "invarian*"],
    "unvarying":    ["unvarying"],
    "immutable":    ["immutab*"],
    "unchanging":   ["unchanging", "unchanged"],
    # --- sinónimos e relacionados ----------------------------------------
    "constant":     ["constant*"],
    "consistent":   ["consisten*"],
    "regular":      ["regular*"],
    "stable":       ["stab*"],
    "steady":       ["stead*"],
    "sustained":    ["sustain*"],
    "static":       ["static", "stasis"],
    "monotonous":   ["monoton*"],
    # --- campo contrastante ----------------------------------------------
    "varied":       ["varied", "variety", "variet*"],
    "varying":      ["varying", "variab*"],
    "changing":     ["changing", "changeab*"],
    "irregular":    ["irregular*"],
    "unequal":      ["unequal", "uneven*"],
    "diverse":      ["divers*"],
    "mutable":      ["mutab*"],
    "multiform":    ["multiform*"],
    # --- eixo da homogeneidade (controlo) --------------------------------
    "homogeneous":  ["homogene*", "homogenous"],
    "heterogeneous": ["heterogene*"],
    # --- controlo negativo -----------------------------------------------
    "loud":         ["loud*"],
    "orchestral":   ["orchestral"],
}

# ---------------------------------------------------------------------------
# 2. SEGMENTAÇÃO E TOKENIZAÇÃO
# ---------------------------------------------------------------------------

# Hífen interno preservado (di-uniform, rotation-invariant); hífen de
# fim de linha continua a ser colado por RE_HIFEN_QUEBRA em normaliza().
RE_TOKEN = re.compile(
    r"[A-Za-zÀ-ÖØ-öø-ÿ](?:[A-Za-zÀ-ÖØ-öø-ÿ'’]|-(?=[A-Za-zÀ-ÖØ-öø-ÿ]))*"
)
RE_FIM_FRASE = re.compile(r"[.!?;]+[\"'”’\)\]]*\s")
RE_HIFEN_QUEBRA = re.compile(r"(\w)-\s+(\w)")


SCHEMA_NEAR = 2
