#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Validações automáticas da análise (req. 3, 9, 10)."""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

import textura_lexico as tlex


class ErroValidacaoAnalise(Exception):
    """Erro bloqueante da análise (configuração / polaridade)."""


@dataclass
class ContagemExclusoes:
    """Rastreio de elegíveis vs excluídos por teste."""
    teste: str
    unidade: str
    n_elegiveis: int
    n_excluidos: int
    padroes_incluidos: list[str] = field(default_factory=list)
    regras_dedupe: str = ""
    parametros: str = ""

    @property
    def pct_excluidos(self) -> float:
        tot = self.n_elegiveis + self.n_excluidos
        return round(100.0 * self.n_excluidos / tot, 2) if tot else 0.0

    def como_dict(self) -> dict:
        return {
            "teste": self.teste,
            "unidade_analise": self.unidade,
            "padroes_incluidos": "; ".join(self.padroes_incluidos),
            "n_ocorrencias_elegiveis": self.n_elegiveis,
            "n_ocorrencias_excluidas": self.n_excluidos,
            "pct_excluidas": self.pct_excluidos,
            "regras_deduplicacao": self.regras_dedupe,
            "parametros": self.parametros,
        }


def validar_configuracao(termos: dict[str, tlex.TermoConfig],
                         search_patterns_extra: list[str] | None = None) -> tlex.RelatorioMapeamento:
    """Garante que todo o padrão de pesquisa tem categoria (req. 3)."""
    rel = tlex.RelatorioMapeamento()
    conhecidos = set(tlex.todos_search_patterns(termos))
    for p in search_patterns_extra or []:
        if p not in conhecidos and p not in tlex.NODE_PATTERNS:
            rel.padroes_sem_categoria.append(p)

    # padrões na config sem categoria canónica (não deveria acontecer)
    for t in termos.values():
        if not t.canonical_term:
            rel.padroes_sem_categoria.extend(t.patterns)

    if rel.padroes_sem_categoria:
        raise ErroValidacaoAnalise(
            "Padrões de pesquisa sem categoria na configuração: "
            + ", ".join(sorted(set(rel.padroes_sem_categoria)))
        )
    return rel


def validar_polaridade_ocorrencias(
        res: pd.DataFrame,
        termos: dict[str, tlex.TermoConfig],
) -> tlex.RelatorioMapeamento:
    """Erro se termo do campo lexical (orientação) tiver polaridade nula (req. 3, 9)."""
    rel = tlex.RelatorioMapeamento()
    # Todos os tipos de dominio campo_lexical DEVEM ter polaridade
    campo = {
        k: v for k, v in termos.items()
        if v.dominio == "campo_lexical"
    }
    sem_pol = [t.canonical_term for t in campo.values() if t.polaridade is None]
    for t in campo.values():
        if t.polaridade is None:
            rel.padroes_sem_polaridade.extend(t.patterns)

    if sem_pol or rel.padroes_sem_polaridade:
        raise ErroValidacaoAnalise(
            "Padrões/tipos dos dois campos sem valor de polaridade "
            "(análise de orientação abortada): "
            + ", ".join(sorted(set(sem_pol + rel.padroes_sem_polaridade)))
        )

    if res is not None and not res.empty and "canonical_term" in res.columns:
        sub = res[res["canonical_term"].isin(campo)]
        nulos = sub[sub["polaridade"].isna() | (sub["polaridade"].astype(str) == "")]
        if len(nulos):
            formas = sorted(nulos["canonical_term"].astype(str).unique())
            raise ErroValidacaoAnalise(
                "Ocorrências com polaridade nula para tipos dos dois campos: "
                + ", ".join(formas)
            )
    return rel


def validar_formas_mapeadas(
        res: pd.DataFrame,
        termos: dict[str, tlex.TermoConfig],
) -> tlex.RelatorioMapeamento:
    """Lista matched_form / query_pattern sem categoria (relatório, req. 3)."""
    import re

    def _rx(p: str) -> re.Pattern:
        esq, dir_ = p.startswith("*"), p.endswith("*")
        nuc = re.escape(p.strip("*"))
        star = r"\w{0,14}"
        return re.compile(
            "^" + (star if esq else "") + nuc + (star if dir_ else "") + "$")

    rel = tlex.RelatorioMapeamento()
    if res is None or res.empty:
        return rel
    pad_map = tlex.mapa_padrao_para_canonico(termos)
    canonicos = set(termos)
    if "query_pattern" in res.columns:
        for p in res["query_pattern"].dropna().astype(str).unique():
            if p and p not in pad_map:
                rel.padroes_sem_categoria.append(p)
    if "canonical_term" in res.columns:
        for c in res["canonical_term"].dropna().astype(str).unique():
            if c and c not in canonicos:
                rel.formas_sem_categoria.append(c)
    if "matched_form" in res.columns and "canonical_term" in res.columns:
        for _, row in res[["matched_form", "canonical_term"]].dropna().iterrows():
            forma = str(row["matched_form"])
            can = str(row["canonical_term"])
            pads = termos[can].patterns if can in termos else ()
            toks = forma.split()
            ok = any(
                len(p.split()) == len(toks) and all(
                    _rx(w).match(tok) for w, tok in zip(p.split(), toks))
                for p in pads
            ) if pads else False
            if not ok and forma:
                rel.formas_sem_categoria.append(forma)
    rel.padroes_sem_categoria = sorted(set(rel.padroes_sem_categoria))
    rel.formas_sem_categoria = sorted(set(rel.formas_sem_categoria))
    return rel


def metadados_teste_dataframe(contagens: list[ContagemExclusoes]) -> pd.DataFrame:
    if not contagens:
        return pd.DataFrame(columns=[
            "teste", "unidade_analise", "padroes_incluidos",
            "n_ocorrencias_elegiveis", "n_ocorrencias_excluidas",
            "pct_excluidas", "regras_deduplicacao", "parametros",
        ])
    return pd.DataFrame([c.como_dict() for c in contagens])
