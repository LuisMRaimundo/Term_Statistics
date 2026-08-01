#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Constantes e padrões partilhados da extracção NEAR (Phase 1 move)."""

from __future__ import annotations

import re
from typing import FrozenSet


COLUNAS = ["L5", "L4", "L3", "L2", "L1", "NODE",
           "R1", "R2", "R3", "R4", "R5",
           "caminho", "url", "raiz", "contexto", "n"]

# Formas do nó por língua (col. 6). Ajustar conforme necessário.
NOS = {
    "en": {"texture", "textures", "textural", "textured", "texturally", "texturing"},
    "pt": {"textura", "texturas", "textural", "texturais"},
    "de": {"textur", "texturen"},
}

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

# Negação: "no" NÃO entra aqui (colisões com «no. 1» — ver _nega_no_opus).
NEGACAO = {"not", "never", "nor", "neither", "without", "lacking",
           "hardly", "scarcely", "barely", "rarely", "seldom",
           "n't", "cannot", "isn", "aren", "wasn", "weren", "doesn", "don",
           "absence", "lack", "devoid"}

# Graduação (antes chamada incorrectamente «modalização»).
GRADUACAO = {"less", "more", "quite", "rather", "fairly", "somewhat",
             "relatively", "certain", "some", "largely", "broadly",
             "mostly", "generally", "almost", "nearly", "increasingly",
             "essentially", "virtually", "comparatively", "slightly",
             "very", "highly", "fully", "completely", "entirely"}
MODALIZACAO = GRADUACAO  # alias de leitura legado

# Operadores modais / evidenciais (coluna modalizado).
MODALIDADE = {
    "may", "might", "could", "would", "can", "should", "must",
    "seems", "seem", "appears", "appear", "apparently", "perhaps",
    "possibly", "presumably", "arguably", "suggests", "tends",
    "likely", "unlikely", "pode", "poderia", "parece", "talvez",
    "seemingly",
}

RELACOES_NUCLEARES = frozenset({
    "atributiva", "predicativa", "predicativa_secundaria",
    "nominal_composto", "nominal_genitiva", "adverbial",
})

# Cópulas: presença entre nó e termo indicia construção predicativa.
COPULAS = {"is", "are", "was", "were", "be", "been", "being",
           "remains", "remain", "remained", "becomes", "become", "became",
           "seems", "seem", "seemed", "appears", "appear", "appeared",
           "stays", "stayed"}

# Abreviaturas que terminam em ponto sem fechar frase.
ABREVIATURAS = {
    "p", "pp", "vol", "vols", "no", "nos", "ed", "eds", "cf", "ibid", "op",
    "cit", "et", "al", "e.g", "i.e", "fig", "figs", "ex", "exx", "ms", "mss",
    "mr", "mrs", "ms", "dr", "prof", "st", "ca", "c", "n", "trans", "rev",
    "repr", "diss", "univ", "publ", "chap", "chaps", "sec", "secs", "bk",
    "jan", "feb", "mar", "apr", "jun", "jul", "aug", "sep", "sept", "oct",
    "nov", "dec", "esp", "viz", "etc", "bwv", "kv", "hob", "d", "k",
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
