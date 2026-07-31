#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Legendas editáveis dos gráficos (JSON) — padrão em português."""

from __future__ import annotations

import json
from pathlib import Path

PADRAO = {
    "rodape": "TEXTURA  ·  pesquisa bibliográfica de termos",
    "sankey": {
        "titulo": "Sankey: termos → documentos",
        "subtitulo": "Largura da faixa = nº de ocorrências (estilo ATLAS.ti)",
        "eixo_esq": "Termos",
        "eixo_dir": "Documentos",
    },
    "nuvem": {
        "titulo": "Nuvem de palavras",
        "subtitulo": "Tamanho proporcional à frequência",
    },
    "docs": {
        "titulo": "Dispersão lexical pelos documentos",
        "subtitulo": "Ordenado por nº de hits (após desduplicar contextos)",
        "xlabel": "Ocorrências",
    },
    "formas": {
        "titulo": "Frequência por termo canónico",
        "subtitulo": "Hits nucleares em 8_Concordancia (sem desduplicar contexto)",
        "xlabel": "Ocorrências (N_hits)",
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

# Preferências do utilizador (sobrevive entre sessões da GUI)
DEFEITO_UTILIZADOR = Path(__file__).resolve().parent / "legendas_defeito.json"


def carregar_defeito_utilizador() -> dict:
    """PADRAO fundido com ``legendas_defeito.json`` (se existir)."""
    return carregar(DEFEITO_UTILIZADOR if DEFEITO_UTILIZADOR.exists() else None)


def guardar_defeito_utilizador(legendas: dict) -> Path:
    """Grava as legendas actuais como defeito da próxima sessão."""
    # Não persistir subtítulos com «Consulta: …» injectada pelo motor
    limpo = json.loads(json.dumps(legendas))
    for chave in ("sankey", "nuvem", "docs", "formas", "near"):
        bloco = limpo.get(chave)
        if not isinstance(bloco, dict):
            continue
        sub = str(bloco.get("subtitulo") or "")
        if sub.startswith("Consulta:"):
            # manter só a parte após « · » se existir; senão limpar
            if "  ·  " in sub:
                bloco["subtitulo"] = sub.split("  ·  ", 1)[1].strip()
            else:
                bloco["subtitulo"] = PADRAO.get(chave, {}).get("subtitulo", "")
    guardar(DEFEITO_UTILIZADOR, limpo)
    return DEFEITO_UTILIZADOR


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
