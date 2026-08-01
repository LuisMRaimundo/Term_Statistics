#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
textura_near.py — mineração de co-ocorrências NEAR/x sobre a matriz KWIC
=======================================================================

Compatibility shim (Phase 1). Implementation lives in the ``textura`` package.
Public and test-facing names are re-exported here so ``import textura_near as tn``
keeps working unchanged.

Uso:
    python textura_near.py --xlsx CAMINHO.xlsx --near 4 --lingua en
    python textura_near.py --xlsx CAMINHO.xlsx --near 4 --limite 20000
"""

from __future__ import annotations

# --- config / constants ---------------------------------------------------
from textura.config import (
    ABREVIATURAS,
    CAMPO,
    COLUNAS,
    COPULAS,
    GRADUACAO,
    MODALIDADE,
    MODALIZACAO,
    NEGACAO,
    NOS,
    RE_FIM_FRASE,
    RE_HIFEN_QUEBRA,
    RE_TOKEN,
    RELACOES_NUCLEARES,
    SCHEMA_NEAR,
)

# --- tokenisation / NEAR --------------------------------------------------
from textura.tokenizacao import (
    _Consulta,
    _casa_em,
    _rx_palavra,
    anota_polaridade_linear,
    anota_sintaxe,
    compila_campo,
    emparelha_contexto,
    encontra_termos,
    fronteiras_frase,
    indice_frase,
    indices_no,
    melhor_par_tokens,
    normaliza,
    procura_near,
    recalcular_distancia_lado,
    tokeniza,
)

# --- syntax ---------------------------------------------------------------
from textura.relacoes import (
    _RX_TEXTURAL,
    _amod_coordenado_do_no,
    _associativa_heterogenea,
    _coordenacao_heterogenea,
    _e_token_textural,
    _escopo_negacao,
    _gov_efectivo,
    _mesmo_token,
    _no_no_subtree,
    _relacao_dependencia_base,
    _resultado_rel,
    _tem_complemento_genitivo,
    _token_em,
    anota_com_heuristica,
    anota_com_spacy,
    relacao_dependencia,
)

# --- lexicon (window domain) ---------------------------------------------
from textura.lexico import (
    DOMINIO_JANELA_LEXICO,
    dominio_janela,
)

# --- language registry (Phase 3) -----------------------------------------
from textura.linguas import (
    CODIGOS as LINGUAS_CODIGOS,
    REGISTO as LINGUAS,
    obter as lingua_obter,
    resolver_execucao,
)

# --- review vocabulary / column order (Phase 5) --------------------------
from textura.exportacao import (
    COLUNAS_HITS_PRIORIDADE,
    DESCRICAO_COLUNAS_HITS,
)
from textura.revisao import (
    REVISAO_EXACTAS,
    REVISAO_PREFIXOS,
    REVISAO_VOCABULARIO,
    validar_revisao_sugerida,
)

# --- statistics -----------------------------------------------------------
from textura.estatistica import (
    POLO_ESTABILIDADE,
    POLO_VARIABILIDADE,
    benjamini_hochberg,
    bootstrap_proporcao,
    conta_tokens,
    dispersao_gries_dp,
    eixo_semantico,
    juilland_d,
    medidas_associacao,
    pielou,
    polaridade,
    shannon,
    simpson_inverso,
)

# --- duplicates / IDs -----------------------------------------------------
from textura.duplicados import (
    _grupos_citacao_entre_docs,
    _partilha_ngrama,
    _score_citacao,
    _score_sobrevivente_janela,
    agregar_ocorrencias,
    atribuir_match_ids,
    deduplicar_hits_exactos,
    fundir_janelas_e_marcar_duplicados,
    hit_key_de,
    occurrence_id_de,
)

# --- export ---------------------------------------------------------------
from textura.exportacao import (
    grafico_distancias,
    grafico_frequencias,
    grafico_polaridade,
    reordenar_colunas_hits,
)

# --- CLI ------------------------------------------------------------------
from textura.pipeline import _configurar_consola, main

# Intentional re-exports for ``import textura_near as tn`` (ruff F401).
__all__ = [
    "ABREVIATURAS",
    "CAMPO",
    "COLUNAS",
    "COLUNAS_HITS_PRIORIDADE",
    "COPULAS",
    "DESCRICAO_COLUNAS_HITS",
    "DOMINIO_JANELA_LEXICO",
    "GRADUACAO",
    "LINGUAS",
    "LINGUAS_CODIGOS",
    "MODALIDADE",
    "MODALIZACAO",
    "NEGACAO",
    "NOS",
    "POLO_ESTABILIDADE",
    "POLO_VARIABILIDADE",
    "REVISAO_EXACTAS",
    "REVISAO_PREFIXOS",
    "REVISAO_VOCABULARIO",
    "RE_FIM_FRASE",
    "RE_HIFEN_QUEBRA",
    "RE_TOKEN",
    "RELACOES_NUCLEARES",
    "SCHEMA_NEAR",
    "_Consulta",
    "_RX_TEXTURAL",
    "_amod_coordenado_do_no",
    "_associativa_heterogenea",
    "_casa_em",
    "_configurar_consola",
    "_coordenacao_heterogenea",
    "_e_token_textural",
    "_escopo_negacao",
    "_gov_efectivo",
    "_grupos_citacao_entre_docs",
    "_mesmo_token",
    "_no_no_subtree",
    "_partilha_ngrama",
    "_relacao_dependencia_base",
    "_resultado_rel",
    "_rx_palavra",
    "_score_citacao",
    "_score_sobrevivente_janela",
    "_tem_complemento_genitivo",
    "_token_em",
    "agregar_ocorrencias",
    "anota_com_heuristica",
    "anota_com_spacy",
    "anota_polaridade_linear",
    "anota_sintaxe",
    "atribuir_match_ids",
    "benjamini_hochberg",
    "bootstrap_proporcao",
    "compila_campo",
    "conta_tokens",
    "deduplicar_hits_exactos",
    "dispersao_gries_dp",
    "dominio_janela",
    "eixo_semantico",
    "emparelha_contexto",
    "encontra_termos",
    "fronteiras_frase",
    "fundir_janelas_e_marcar_duplicados",
    "grafico_distancias",
    "grafico_frequencias",
    "grafico_polaridade",
    "hit_key_de",
    "indice_frase",
    "indices_no",
    "juilland_d",
    "lingua_obter",
    "main",
    "medidas_associacao",
    "melhor_par_tokens",
    "normaliza",
    "occurrence_id_de",
    "pielou",
    "polaridade",
    "procura_near",
    "recalcular_distancia_lado",
    "relacao_dependencia",
    "reordenar_colunas_hits",
    "resolver_execucao",
    "shannon",
    "simpson_inverso",
    "tokeniza",
    "validar_revisao_sugerida",
]


if __name__ == "__main__":
    raise SystemExit(main())
