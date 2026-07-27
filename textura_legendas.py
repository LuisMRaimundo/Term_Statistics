#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Legendas editáveis dos gráficos (JSON) — padrão em português."""

from __future__ import annotations

import json
from pathlib import Path

PADRAO = {
    "rodape": "TEXTURA  ·  pesquisa bibliográfica de termos",
    "sankey": {
        "titulo": "Sankey: formas casadas → documentos",
        "subtitulo": "Largura da faixa = nº de ocorrências (estilo ATLAS.ti)",
        "eixo_esq": "Formas casadas",
        "eixo_dir": "Documentos",
    },
    "nuvem": {
        "titulo": "Nuvem de formas casadas",
        "subtitulo": "Tamanho proporcional à frequência",
    },
    "docs": {
        "titulo": "Dispersão lexical pelos documentos",
        "subtitulo": "Ordenado por nº de hits (após desduplicar contextos)",
        "xlabel": "Ocorrências",
    },
    "formas": {
        "titulo": "Formas gráficas casadas",
        "subtitulo": "Formas de superfície recuperadas pela consulta booleana",
        "xlabel": "Ocorrências",
    },
    "near": {
        "titulo": "Distribuição das distâncias NEAR",
        "subtitulo": "Distância em tokens entre termos com relação sintáctica",
        "xlabel": "Distância (tokens)",
        "ylabel": "Nº de pares",
        "mediana": "Mediana",
        "media": "Média",
    },
}


def carregar(caminho: Path | str | None, consulta: str = "") -> dict:
    """Carrega legendas; funde com o padrão; injeta a consulta nos subtítulos."""
    leg = json.loads(json.dumps(PADRAO))  # deep copy
    if caminho:
        p = Path(caminho)
        if p.exists():
            extra = json.loads(p.read_text(encoding="utf-8"))
            for k, v in extra.items():
                if isinstance(v, dict) and isinstance(leg.get(k), dict):
                    leg[k].update({kk: vv for kk, vv in v.items() if vv is not None})
                elif v is not None:
                    leg[k] = v
    if consulta:
        for chave in ("sankey", "nuvem", "docs", "formas", "near"):
            sub = str(leg[chave].get("subtitulo", "") or "")
            if "Consulta:" in sub or "Query:" in sub:
                continue
            leg[chave]["subtitulo"] = (
                f"Consulta: {consulta}" + (f"  ·  {sub}" if sub else "")
            )
    return leg


def guardar(caminho: Path | str, legendas: dict) -> None:
    Path(caminho).write_text(
        json.dumps(legendas, ensure_ascii=False, indent=2), encoding="utf-8")


def resumo_titulos(leg: dict) -> str:
    """Uma linha para o registo CLI/GUI."""
    partes = []
    for k in ("sankey", "nuvem", "docs", "formas", "near"):
        t = (leg.get(k) or {}).get("titulo", "")
        if t:
            partes.append(f"{k}={t!r}")
    return "; ".join(partes)
