#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Medidas de associação, dispersão, bootstrap e correcção BH."""

from __future__ import annotations

import math

import numpy as np
from scipy import stats

from textura.lexico import (
    POLO_ESTABILIDADE as _POLO_E,
    POLO_VARIABILIDADE as _POLO_V,
)
from textura.tokenizacao import indice_frase

# Cópias mutáveis: --termos pode registar etiquetas novas em runtime.
POLO_ESTABILIDADE = set(_POLO_E)
POLO_VARIABILIDADE = set(_POLO_V)


def shannon(contagens) -> float:
    total = sum(contagens)
    if total == 0:
        return float("nan")
    p = np.array([c / total for c in contagens if c > 0])
    return float(-(p * np.log(p)).sum())


def pielou(contagens) -> float:
    S = sum(1 for c in contagens if c > 0)
    return float(shannon(contagens) / math.log(S)) if S > 1 else float("nan")


def simpson_inverso(contagens) -> float:
    total = sum(contagens)
    if total == 0:
        return float("nan")
    p = np.array([c / total for c in contagens if c > 0])
    return float(1.0 / (p ** 2).sum())


def benjamini_hochberg(pvals) -> np.ndarray:
    """Correcção BH; devolve p ajustados na ordem de entrada."""
    p = np.asarray(pvals, dtype=float)
    m = len(p)
    ordem = np.argsort(p)
    ajust = np.empty(m, dtype=float)
    anterior = 1.0
    for rank in range(m - 1, -1, -1):
        i = ordem[rank]
        val = min(anterior, p[i] * m / (rank + 1))
        ajust[i] = val
        anterior = val
    return np.clip(ajust, 0, 1)
def conta_tokens(tokens, no_idx, limites, lo, hi, mesma_frase=True) -> int:
    """Nº de tokens cuja distância ao nó está em ]lo, hi], na mesma frase."""
    frase_no = indice_frase(tokens[no_idx][1], limites)
    n = 0
    for k in range(max(0, no_idx - hi), min(len(tokens), no_idx + hi + 1)):
        d = abs(k - no_idx)
        if d <= lo or d > hi:
            continue
        if mesma_frase and indice_frase(tokens[k][1], limites) != frase_no:
            continue
        n += 1
    return n


def medidas_associacao(o11, o12, r1, r2, n_janelas):
    """Medidas de associação sobre a tabela 2x2 'janela vs. banda de referência'.

        o11 = ocorrências do termo dentro de NEAR/x
        o12 = ocorrências do termo na banda de referência (mais distante)
        r1  = total de tokens dentro de NEAR/x
        r2  = total de tokens na banda de referência

    O termo de comparação é a banda distante dos MESMOS contextos, e não o
    corpus integral — o qual não está disponível na matriz KWIC. Trata-se,
    portanto, de uma associação posicional: mede se o termo se concentra
    junto do nó mais do que na sua vizinhança alargada. É esta a leitura
    que deve constar do texto.
    """
    o21, o22 = r1 - o11, r2 - o12
    n = r1 + r2
    if n == 0 or (o11 + o12) == 0:
        return {}
    e11 = r1 * (o11 + o12) / n
    e12 = r2 * (o11 + o12) / n
    e21 = r1 * (o21 + o22) / n
    e22 = r2 * (o21 + o22) / n

    def _g(o, e):
        return 2 * o * math.log(o / e) if o > 0 and e > 0 else 0.0

    ll = _g(o11, e11) + _g(o12, e12) + _g(o21, e21) + _g(o22, e22)
    ll = ll if o11 >= e11 else -ll          # sinal: + atracção, - repulsão

    mi = math.log2(o11 / e11) if o11 > 0 and e11 > 0 else float("nan")
    mi3 = math.log2(o11 ** 3 / e11) if o11 > 0 and e11 > 0 else float("nan")
    tscore = (o11 - e11) / math.sqrt(o11) if o11 > 0 else float("nan")
    zscore = (o11 - e11) / math.sqrt(e11) if e11 > 0 else float("nan")
    dice = 2 * o11 / (n_janelas + o11 + o12) if (n_janelas + o11 + o12) else float("nan")
    logdice = 14 + math.log2(dice) if dice > 0 else float("nan")
    dp_termo = (o11 / r1 - o12 / r2) if r1 and r2 else float("nan")

    # razão de possibilidades com correcção de Haldane e IC de Woolf
    a, b, c, d = o11 + 0.5, o12 + 0.5, o21 + 0.5, o22 + 0.5
    orr = (a * d) / (b * c)
    se = math.sqrt(1 / a + 1 / b + 1 / c + 1 / d)
    ic_inf, ic_sup = math.exp(math.log(orr) - 1.96 * se), math.exp(math.log(orr) + 1.96 * se)

    p_fisher = stats.fisher_exact([[o11, o12], [o21, o22]]).pvalue

    return {
        "O11_janela": o11, "E11_esperado": round(e11, 3),
        "O12_banda_ref": o12,
        "log_likelihood_G2": round(ll, 3),
        "MI": round(mi, 3), "MI3": round(mi3, 3),
        "t_score": round(tscore, 3), "z_score": round(zscore, 3),
        "logDice": round(logdice, 3),
        "DeltaP": round(dp_termo, 5),
        "razao_possib": round(orr, 3),
        "IC95_inf": round(ic_inf, 3), "IC95_sup": round(ic_sup, 3),
        "p_fisher": p_fisher,
    }


def dispersao_gries_dp(cont_por_parte: dict, tam_por_parte: dict) -> float:
    """Desvio de proporções (DP) de Gries. 0 = uniforme; 1 = concentrado."""
    total = sum(cont_por_parte.values())
    tam_tot = sum(tam_por_parte.values())
    if total == 0 or tam_tot == 0:
        return float("nan")
    s = 0.0
    for parte, tam in tam_por_parte.items():
        esperado = tam / tam_tot
        observado = cont_por_parte.get(parte, 0) / total
        s += abs(observado - esperado)
    return round(s / 2, 4)


def juilland_d(cont_por_parte: dict, n_partes: int) -> float:
    """D de Juilland. 1 = distribuição uniforme pelas partes."""
    v = list(cont_por_parte.values()) + [0] * (n_partes - len(cont_por_parte))
    m = float(np.mean(v))
    if m == 0 or n_partes < 2:
        return float("nan")
    cv = float(np.std(v, ddof=0)) / m
    return round(1 - cv / math.sqrt(n_partes - 1), 4)


def bootstrap_proporcao(v, n_rep=5000, semente=20260724):
    """IC95% percentílico para uma proporção binária, por reamostragem."""
    rng = np.random.default_rng(semente)
    v = np.asarray(v, dtype=float)
    if v.size == 0:
        return (float("nan"), float("nan"))
    amostras = rng.choice(v, size=(n_rep, v.size), replace=True).mean(axis=1)
    return (round(float(np.percentile(amostras, 2.5)), 4),
            round(float(np.percentile(amostras, 97.5)), 4))


def polaridade(tipo: str, negado: bool | None = False,
               *, inverter_negada: bool = False) -> str | None:
    """Polaridade de base; inversão sob negação só se inverter_negada=True."""
    if tipo in POLO_ESTABILIDADE:
        base = "estabilidade"
    elif tipo in POLO_VARIABILIDADE:
        base = "variabilidade"
    else:
        return None
    if inverter_negada and negado:
        base = "variabilidade" if base == "estabilidade" else "estabilidade"
    return base


def eixo_semantico(tipo: str) -> str:
    """Eixo: homogeneidade_sincronica | invariancia_diacronica | ambos."""
    sinc = {"uniform", "homogeneous", "heterogeneous", "diverse",
            "varied", "unequal", "irregular", "multiform", "mutable"}
    diac = {"static", "invariable", "unvarying", "immutable", "unchanging",
            "constant", "varying", "changing", "stable", "steady", "sustained",
            "monotonous", "consistent", "regular"}
    if tipo in sinc and tipo in diac:
        return "ambos"
    if tipo in sinc:
        return "homogeneidade_sincronica"
    if tipo in diac:
        return "invariancia_diacronica"
    return "ambos"
