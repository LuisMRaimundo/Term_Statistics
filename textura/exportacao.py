#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Gráficos auxiliares e ordenação/escrita Excel da concordância."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


def grafico_frequencias(df, destino: Path):
    col = "doc_id" if "doc_id" in df.columns else "caminho"
    cont = df.groupby("termo_tipo")[col].nunique().sort_values(ascending=False)
    fig, ax = plt.subplots(figsize=(9, max(4, 0.32 * len(cont))))
    ax.barh(cont.index[::-1], cont.values[::-1], color="#4a5a6a")
    ax.set_xlabel("Obras únicas com co-ocorrência")
    ax.set_ylabel("")
    ax.set_title("Dispersão do campo lexical (obras únicas)")
    ax.grid(axis="x", linewidth=0.4, alpha=0.5)
    fig.tight_layout()
    fig.savefig(destino, dpi=150)
    plt.close(fig)


def grafico_distancias(df, destino: Path):
    fig, ax = plt.subplots(figsize=(7, 4))
    for lado, cor in (("esq", "#4a5a6a"), ("dir", "#a8763e")):
        sub = df[df["lado"] == lado]["distancia"]
        ax.hist(sub, bins=np.arange(0.5, df["distancia"].max() + 1.5),
                alpha=0.65, label=f"{lado}erda" if lado == "esq" else "direita",
                color=cor)
    ax.set_xlabel("Distância em tokens ao nó")
    ax.set_ylabel("Co-ocorrências")
    ax.set_title("Distribuição da distância por lado")
    ax.legend()
    ax.grid(axis="y", linewidth=0.4, alpha=0.5)
    fig.tight_layout()
    fig.savefig(destino, dpi=150)
    plt.close(fig)


def grafico_polaridade(df, destino: Path):
    sub = df.replace({"polaridade": {"": np.nan}}).dropna(subset=["polaridade"])
    col = ("relacao_sintactica" if "relacao_sintactica" in sub.columns
           else "relacao")
    tab = pd.crosstab(sub[col], sub["polaridade"])
    fig, ax = plt.subplots(figsize=(7, 4))
    tab.plot(kind="bar", stacked=True, ax=ax,
             color=["#4a5a6a", "#a8763e"])
    ax.set_xlabel("Relação sintáctica")
    ax.set_ylabel("Co-ocorrências")
    ax.set_title("Polaridade por relação sintáctica")
    ax.grid(axis="y", linewidth=0.4, alpha=0.5)
    fig.tight_layout()
    fig.savefig(destino, dpi=150)
    plt.close(fig)
# Ordem canónica das colunas de ``8_Concordancia`` (fonte viva do dicionário).
COLUNAS_HITS_PRIORIDADE: tuple[str, ...] = (
    "source_matrix_row", "texture_occurrence_id", "match_id", "hit_key",
    "grupo_passagem_id", "candidato_duplicado",
    "no", "termo_tipo", "canonical_term", "query_pattern",
    "termo_forma", "matched_form", "n_palavras", "distancia", "lado",
    "negado", "graduado", "modalizado", "relacao_sintactica",
    "polaridade_base", "polaridade", "eixo",
    "censurado_esq", "censurado_dir",
    "idx_no", "idx_termo", "off_no", "off_termo",
    "n_nos_janela", "forma_em_composto",
    "caminho_ficheiro", "doc_id", "url", "contexto",
    "motivo_exclusao", "nuclear", "fonte_classificacao",
    "n_janelas_fundidas", "revisao_sugerida",
    "nucleo_da_propriedade", "orientacao", "governante", "percurso_dep",
    "dominio", "dominio_janela", "revisto_por_humano", "nota_revisao",
)

# Descrições curtas (PT) — consumidas por ``utilitarios/gera_dicionario_colunas.py``.
DESCRICAO_COLUNAS_HITS: dict[str, str] = {
    "source_matrix_row": "N.º de linha na matriz KWIC de origem",
    "texture_occurrence_id": "ID estável da ocorrência mestra de textura",
    "match_id": "ID do hit NEAR dentro da ocorrência",
    "hit_key": "Chave de deduplicação exacta do hit",
    "grupo_passagem_id": "Grupo de passagem/janela sobreposta",
    "candidato_duplicado": "Etiquetas C#### / P#### / J#### de possível duplicado",
    "no": "Forma do nó (texture / textura / …) na janela",
    "termo_tipo": "Etiqueta do tipo lexical do campo (canonical bucket)",
    "canonical_term": "Termo canónico adjudicado do campo lexical",
    "query_pattern": "Padrão (truncatura) que casou o termo",
    "termo_forma": "Forma lexical do termo na janela",
    "matched_form": "Forma exactamente casada no contexto",
    "n_palavras": "N.º de tokens do match multiword (se aplicável)",
    "distancia": "Distância em tokens entre termo e nó",
    "lado": "Lado do termo relativamente ao nó (esq/dir)",
    "negado": "Negação no escopo (nao / directo / indirecto / vazio)",
    "graduado": "Presença de graduação (more/rather/…)",
    "modalizado": "Presença de modalidade/evidencialidade",
    "relacao_sintactica": "Classe sintáctica (taxonomia nuclear/não-nuclear)",
    "polaridade_base": "Polaridade antes de inversão por negação",
    "polaridade": "Polaridade efectiva (estabilidade / variabilidade)",
    "eixo": "Eixo semântico (homogeneidade_sincronica / …)",
    "censurado_esq": "Contexto truncado à esquerda na matriz",
    "censurado_dir": "Contexto truncado à direita na matriz",
    "idx_no": "Índice token do nó na janela",
    "idx_termo": "Índice token do termo na janela",
    "off_no": "Offset de carácter do nó no contexto",
    "off_termo": "Offset de carácter do termo no contexto",
    "n_nos_janela": "N.º de formas do nó na janela",
    "forma_em_composto": "Match dentro de composto hifenizado",
    "caminho_ficheiro": "Caminho/fonte documental (valor de dados, não path OS)",
    "doc_id": "Identificador estável do documento",
    "url": "URL se presente na matriz",
    "contexto": "Janela textual exportada (evidência)",
    "motivo_exclusao": "Motivo de exclusão / não-nuclear",
    "nuclear": "TRUE = entra na análise pós-revisão",
    "fonte_classificacao": "dependencias (spaCy) ou heuristica",
    "n_janelas_fundidas": "Janelas fundidas por sobreposição",
    "revisao_sugerida": "Etiquetas de revisão automática (ver vocabulário)",
    "nucleo_da_propriedade": "Núcleo nominal da propriedade na árvore",
    "orientacao": "Direcção da relação (termo_sobre_no / …)",
    "governante": "Governante sintáctico reportado",
    "percurso_dep": "Percurso de dependência spaCy",
    "dominio": "Domínio documental (triagem path / revisão)",
    "dominio_janela": "Domínio sugerido por pistas na janela",
    "revisto_por_humano": "Marca de revisão humana",
    "nota_revisao": "Nota livre do revisor",
}


def reordenar_colunas_hits(res: pd.DataFrame) -> pd.DataFrame:
    """Coloca identificadores à frente sem perder colunas extra."""
    frente = [c for c in COLUNAS_HITS_PRIORIDADE if c in res.columns]
    resto = [c for c in res.columns if c not in frente]
    return res[frente + resto]
