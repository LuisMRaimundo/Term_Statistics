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
def reordenar_colunas_hits(res: pd.DataFrame) -> pd.DataFrame:
    """Coloca identificadores à frente sem perder colunas extra."""
    prioridade = [
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
    ]
    frente = [c for c in prioridade if c in res.columns]
    resto = [c for c in res.columns if c not in frente]
    return res[frente + resto]
