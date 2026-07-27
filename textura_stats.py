#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
textura_stats.py — camada estatística avançada
==============================================

Importado por textura_near.py. Reúne os procedimentos correntes na
linguística de corpus contemporânea que ultrapassam a estatística
descritiva: medidas de efeito com intervalo, inferência por reamostragem,
modelação multivariável, análise de correspondências e classificação
hierárquica dos colocados.

Dependências obrigatórias: numpy, scipy, matplotlib, pandas.
Opcional: statsmodels (regressão logística; sem ela, é usada uma
implementação própria por máxima verosimilhança com IRLS).
"""

from __future__ import annotations

import math

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy import stats
from scipy.cluster.hierarchy import dendrogram, linkage
from scipy.spatial.distance import pdist


# ---------------------------------------------------------------------------
# 1. MEDIDAS DE EFEITO
# ---------------------------------------------------------------------------

def log_ratio(o11, o12, r1, r2, k=0.5):
    """Log Ratio de Hardie (2014): medida de efeito para colocação/keyness.

    Ao contrário de G² e de p, não cresce com a dimensão da amostra:
    exprime quantas vezes o termo é mais provável na janela do que na
    banda de referência, em log2. Suavização de +k evita divisão por zero.
    LR = 1 significa duas vezes mais provável; LR = 2, quatro vezes.
    """
    if r1 <= 0 or r2 <= 0:
        return float("nan")
    p1 = (o11 + k) / (r1 + 2 * k)
    p2 = (o12 + k) / (r2 + 2 * k)
    return round(math.log2(p1 / p2), 4)


def bayes_factor_proporcao(k, n, p0=0.5):
    """Factor de Bayes (Jeffreys, a priori Beta(1,1)) de H1 contra H0: p = p0.

    BF10 > 3 constitui evidência moderada a favor de H1; BF10 < 1/3,
    evidência a favor de H0. Ao contrário do valor de p, quantifica
    também a evidência a favor da hipótese nula — o que importa quando o
    resultado esperado é a ausência de inclinação.
    """
    if n == 0:
        return float("nan")
    # verosimilhança marginal sob H1 (a priori uniforme) = Beta(k+1, n-k+1)
    log_m1 = (math.lgamma(k + 1) + math.lgamma(n - k + 1) - math.lgamma(n + 2))
    log_m0 = k * math.log(p0) + (n - k) * math.log(1 - p0)
    return round(math.exp(log_m1 - log_m0), 4)


def permutacao_diferenca(grupo_a, grupo_b, n_rep=10000, semente=20260724):
    """Teste de permutação para a diferença de proporções entre dois grupos.

    Não pressupõe normalidade nem frequências esperadas mínimas, pelo que
    é preferível ao χ² quando alguma célula é pouco povoada.
    """
    rng = np.random.default_rng(semente)
    a, b = np.asarray(grupo_a, float), np.asarray(grupo_b, float)
    if a.size == 0 or b.size == 0:
        return float("nan"), float("nan")
    obs = a.mean() - b.mean()
    junto = np.concatenate([a, b])
    na = a.size
    conta = 0
    for _ in range(n_rep):
        rng.shuffle(junto)
        if abs(junto[:na].mean() - junto[na:].mean()) >= abs(obs):
            conta += 1
    return round(float(obs), 4), (conta + 1) / (n_rep + 1)


# ---------------------------------------------------------------------------
# 2. REGRESSÃO LOGÍSTICA
# ---------------------------------------------------------------------------

def _irls(X, y, max_iter=60, tol=1e-9, firth: bool = False):
    """Máxima verosimilhança por IRLS; opcionalmente penalização de Firth."""
    n, p = X.shape
    beta = np.zeros(p)
    for _ in range(max_iter):
        eta = X @ beta
        mu = 1.0 / (1.0 + np.exp(-np.clip(eta, -35, 35)))
        w = np.clip(mu * (1 - mu), 1e-9, None)
        z = eta + (y - mu) / w
        XtW = X.T * w
        Hinv = None
        try:
            XtWX = XtW @ X
            novo = np.linalg.solve(XtWX, XtW @ z)
            if firth:
                Hinv = np.linalg.pinv(XtWX)
        except np.linalg.LinAlgError:
            XtWX = XtW @ X
            novo = np.linalg.lstsq(XtWX, XtW @ z, rcond=None)[0]
            if firth:
                Hinv = np.linalg.pinv(XtWX)
        if firth and Hinv is not None:
            # score de Firth: 0.5 * diag(H) * X  (aproximação)
            h = np.clip(np.sum((X @ Hinv) * X, axis=1), 0, 1)
            score = X.T @ (y - mu + h * (0.5 - mu))
            try:
                novo = beta + np.linalg.solve(XtWX, score)
            except np.linalg.LinAlgError:
                novo = beta + np.linalg.lstsq(XtWX, score, rcond=None)[0]
        if np.max(np.abs(novo - beta)) < tol:
            beta = novo
            break
        beta = novo
    eta = X @ beta
    mu = 1.0 / (1.0 + np.exp(-np.clip(eta, -35, 35)))
    w = np.clip(mu * (1 - mu), 1e-9, None)
    cov = np.linalg.pinv((X.T * w) @ X)
    return beta, cov, mu


def _se_cluster_cr1(X, y, mu, beta, clusters) -> np.ndarray:
    """Erros-padrão sanduíche CR1 (agrupados por obra)."""
    clusters = np.asarray(clusters)
    n, p = X.shape
    meat = np.zeros((p, p))
    g_unique = pd.unique(clusters)
    G = len(g_unique)
    for g in g_unique:
        m = clusters == g
        Xg = X[m]
        ug = (y[m] - mu[m]).reshape(-1, 1)
        Sg = Xg.T @ ug
        meat += Sg @ Sg.T
    bread = np.linalg.pinv((X.T * np.clip(mu * (1 - mu), 1e-9, None)) @ X)
    # correcção CR1
    fator = (G / max(G - 1, 1)) * ((n - 1) / max(n - p, 1))
    cov = fator * (bread @ meat @ bread)
    return np.sqrt(np.clip(np.diag(cov), 0, None))


def regressao_logistica(df, alvo="polaridade", positivo="estabilidade",
                        preditores=("relacao", "lado", "distancia"),
                        cluster: str | None = None,
                        firth_se_separacao: bool = True,
                        min_nivel: int = 10):
    """Regressão logística com referência = categoria mais frequente,
    agrupamento de níveis raros, Firth se separação, SE cluster CR1.
    """
    d = df.dropna(subset=[alvo]).copy()
    preditores = [p for p in preditores if p in d.columns]
    if d.empty or not preditores:
        return pd.DataFrame(), {}

    y = (d[alvo] == positivo).astype(float).values
    if y.min() == y.max():
        return pd.DataFrame(), {"aviso": "resposta constante"}

    niveis_agrupados = []
    partes, nomes = [np.ones((len(d), 1))], ["(intercepção)"]
    for p in preditores:
        if pd.api.types.is_numeric_dtype(d[p]):
            v = d[p].astype(float).values
            partes.append(((v - v.mean()) / (v.std() or 1)).reshape(-1, 1))
            nomes.append(f"{p} (padronizada)")
        else:
            vc = d[p].astype(str).value_counts()
            raros = set(vc[vc < min_nivel].index)
            serie = d[p].astype(str).where(~d[p].astype(str).isin(raros),
                                           other="outras")
            if raros:
                niveis_agrupados.extend(f"{p}:{r}" for r in sorted(raros))
            # referência = categoria mais frequente (nao alfabetica)
            ref = serie.value_counts().index[0]
            cats = [c for c in serie.value_counts().index if c != ref]
            for c in cats:
                partes.append((serie == c).astype(float).values.reshape(-1, 1))
                nomes.append(f"{p}: {c} vs {ref}")
    X = np.hstack(partes)

    beta, cov, mu = _irls(X, y, firth=False)
    se = np.sqrt(np.clip(np.diag(cov), 0, None))
    usou_firth = False
    if firth_se_separacao and (
            (np.abs(beta) > 10).any() or (se > 100).any()):
        beta, cov, mu = _irls(X, y, firth=True)
        se = np.sqrt(np.clip(np.diag(cov), 0, None))
        usou_firth = True

    se_tipo = "modelo"
    if cluster and cluster in d.columns:
        se = _se_cluster_cr1(X, y, mu, beta, d[cluster].astype(str).values)
        se_tipo = f"cluster CR1 ({cluster})"

    z = np.divide(beta, se, out=np.zeros_like(beta), where=se > 0)
    pvals = 2 * (1 - stats.norm.cdf(np.abs(z)))

    lo = np.clip(beta - 1.96 * se, -50, 50)
    hi = np.clip(beta + 1.96 * se, -50, 50)
    tab = pd.DataFrame({
        "termo": nomes,
        "coef_log_odds": np.round(beta, 4),
        "erro_padrao": np.round(se, 4),
        "razao_possib": np.round(np.exp(np.clip(beta, -50, 50)), 4),
        "IC95_inf": np.round(np.exp(lo), 4),
        "IC95_sup": np.round(np.exp(hi), 4),
        "z": np.round(z, 3),
        "p": [f"{v:.4g}" for v in pvals],
    })

    ll = float(np.sum(y * np.log(np.clip(mu, 1e-12, 1)) +
                      (1 - y) * np.log(np.clip(1 - mu, 1e-12, 1))))
    pbar = y.mean()
    ll0 = float(len(y) * (pbar * math.log(max(pbar, 1e-12))
                          + (1 - pbar) * math.log(max(1 - pbar, 1e-12))))
    ajuste = {
        "n": len(y),
        "log-verosimilhança": round(ll, 3),
        "R² de McFadden": round(1 - ll / ll0, 4) if ll0 else float("nan"),
        "AIC": round(-2 * ll + 2 * X.shape[1], 3),
        "exactidão de classificação": round(
            float(((mu > .5) == (y > .5)).mean()), 4),
        "taxa de base": round(float(max(pbar, 1 - pbar)), 4),
        "erros_padrao": se_tipo,
        "firth": "sim" if usou_firth else "nao",
        "niveis_agrupados": niveis_agrupados,
    }
    if usou_firth:
        ajuste["aviso"] = "Firth aplicado (separacao / coeficientes extremos)"
    return tab, ajuste


# ---------------------------------------------------------------------------
# 3. ANÁLISE DE CORRESPONDÊNCIAS
# ---------------------------------------------------------------------------

def analise_correspondencias(tabela: pd.DataFrame, destino=None):
    """AC simples sobre uma tabela de contingência (decomposição em valores
    singulares da matriz de resíduos padronizados de Pearson).

    Projecta tipos lexicais e categorias sintácticas no mesmo plano,
    revelando associações que a leitura linha a linha de um χ² não expõe.
    """
    N = tabela.values.astype(float)
    total = N.sum()
    if total == 0 or min(N.shape) < 2:
        return pd.DataFrame(), pd.DataFrame(), []
    P = N / total
    r, c = P.sum(1), P.sum(0)
    Dr, Dc = np.diag(1 / np.sqrt(np.clip(r, 1e-12, None))), \
             np.diag(1 / np.sqrt(np.clip(c, 1e-12, None)))
    S = Dr @ (P - np.outer(r, c)) @ Dc
    U, s, Vt = np.linalg.svd(S, full_matrices=False)
    inercia = s ** 2
    prop = np.round(100 * inercia / inercia.sum(), 2)

    k = min(2, len(s))
    lin = pd.DataFrame(Dr @ U[:, :k] * s[:k], index=tabela.index,
                       columns=[f"dim{i+1}" for i in range(k)]).round(4)
    col = pd.DataFrame(Dc @ Vt[:k].T * s[:k], index=tabela.columns,
                       columns=[f"dim{i+1}" for i in range(k)]).round(4)

    if destino is not None and k == 2:
        fig, ax = plt.subplots(figsize=(7.5, 6))
        ax.axhline(0, lw=.5, color="#999"); ax.axvline(0, lw=.5, color="#999")
        ax.scatter(lin["dim1"], lin["dim2"], s=28, color="#4a5a6a")
        for n, (x, y) in lin.iterrows():
            ax.annotate(str(n), (x, y), fontsize=8, xytext=(3, 3),
                        textcoords="offset points", color="#2b3844")
        ax.scatter(col["dim1"], col["dim2"], s=52, marker="^", color="#a8763e")
        for n, (x, y) in col.iterrows():
            ax.annotate(str(n), (x, y), fontsize=9, weight="bold",
                        xytext=(3, -9), textcoords="offset points", color="#7a5324")
        ax.set_xlabel(f"dim 1 ({prop[0]}% da inércia)")
        ax.set_ylabel(f"dim 2 ({prop[1] if len(prop) > 1 else 0}%)")
        ax.set_title("Análise de correspondências: tipos lexicais × relação sintáctica")
        fig.tight_layout(); fig.savefig(destino, dpi=150); plt.close(fig)

    return lin, col, list(prop)


# ---------------------------------------------------------------------------
# 4. CLASSIFICAÇÃO HIERÁRQUICA DOS COLOCADOS
# ---------------------------------------------------------------------------

def _negado_como_fraccao(v) -> float:
    """Converte negado (bool legado ou directo/indirecto/nao) em 0/1."""
    if v is True or v == 1:
        return 1.0
    if v is False or v == 0 or v is None:
        return 0.0
    s = str(v).strip().lower()
    if s in {"directo", "true", "1", "sim"}:
        return 1.0
    return 0.0


def perfis_e_dendrograma(res: pd.DataFrame, destino=None, min_ocorr=5):
    """Agrupa os tipos lexicais pelo seu perfil contextual.

    O perfil de cada tipo é o vector de proporções sobre lado da
    co-ocorrência, distância média, relação sintáctica, negação e
    modalização. A distância é a de Ward sobre perfis padronizados.
    Tipos que partilham perfil comportam-se de modo semelhante no
    discurso, independentemente da frequência.
    """
    cont = res["termo_tipo"].value_counts()
    manter = cont[cont >= min_ocorr].index
    d = res[res["termo_tipo"].isin(manter)]
    if d["termo_tipo"].nunique() < 3:
        return pd.DataFrame(), None

    perfil = pd.DataFrame(index=sorted(d["termo_tipo"].unique()))
    perfil["prop_esquerda"] = d.groupby("termo_tipo")["lado"].apply(
        lambda s: (s == "esq").mean()).round(4)
    perfil["distancia_media"] = d.groupby("termo_tipo")["distancia"].mean().round(4)
    # negado ∈ {directo, indirecto, nao} ou bool legado — so «directo»/True conta
    neg_num = d["negado"].map(_negado_como_fraccao)
    perfil["prop_negado"] = neg_num.groupby(d["termo_tipo"]).mean().round(4)
    mod_num = d["modalizado"].map(
        lambda v: 1.0 if v is True or v == 1 or str(v).lower() in {
            "true", "1"} else 0.0)
    perfil["prop_modalizado"] = mod_num.groupby(d["termo_tipo"]).mean().round(4)
    if "relacao_sintactica" in d.columns:
        col_rel = "relacao_sintactica"
    elif "relacao_dep" in d.columns:
        col_rel = "relacao_dep"
    else:
        col_rel = "relacao"
    prop_rel = (pd.crosstab(d["termo_tipo"], d[col_rel], normalize="index")
                .round(4).add_prefix("rel_"))
    perfil = perfil.join(prop_rel).fillna(0)

    M = perfil.values.astype(float)
    M = (M - M.mean(0)) / np.where(M.std(0) == 0, 1, M.std(0))
    Z = linkage(pdist(M, metric="euclidean"), method="ward")

    if destino is not None:
        fig, ax = plt.subplots(figsize=(9, max(4, 0.3 * len(perfil))))
        dendrogram(Z, labels=list(perfil.index), orientation="right",
                   ax=ax, color_threshold=0.7 * max(Z[:, 2]),
                   above_threshold_color="#999")
        ax.set_title("Agrupamento dos tipos lexicais por perfil contextual (Ward)")
        ax.set_xlabel("distância")
        fig.tight_layout(); fig.savefig(destino, dpi=150); plt.close(fig)

    return perfil.reset_index().rename(columns={"index": "termo_tipo"}), Z


# ---------------------------------------------------------------------------
# 5. CURVA DE RIQUEZA LEXICAL
# ---------------------------------------------------------------------------

def riqueza_lexical(res: pd.DataFrame):
    """Índices independentes da dimensão da amostra.

    A razão tipo/ocorrência varia com o tamanho do corpus e não é
    comparável entre subcorpora; C de Herdan e índice de Guiraud
    corrigem essa dependência.
    """
    formas = res["termo_forma"].astype(str)
    N, V = len(formas), formas.nunique()
    if N == 0:
        return {}
    return {
        "ocorrências (N)": N,
        "formas distintas (V)": V,
        "razão tipo/ocorrência": round(V / N, 4),
        "C de Herdan": round(math.log(V) / math.log(N), 4) if N > 1 else float("nan"),
        "índice de Guiraud": round(V / math.sqrt(N), 4),
        "hapax (formas com 1 ocorrência)": int((formas.value_counts() == 1).sum()),
    }
