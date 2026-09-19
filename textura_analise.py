#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
textura_analise.py — fase 2: estatística e gráficos sobre Excel revisto
======================================================================

Não reextrai. Exige 0_Instrucoes e colunas de revisão da fase 1.

Uso:
    python textura_analise.py --xlsx resultado_near.xlsx
    python textura_analise.py --xlsx resultado_near.xlsx --desduplicacao nenhuma
    python textura_analise.py --xlsx resultado_near.xlsx --legendas legendas.json
"""

from __future__ import annotations

import argparse
import json
import math
import re
import shutil
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font
from scipy import stats

import textura_freq as tfreq
import textura_legendas as tleg
import textura_lexico as tlex
import textura_near as tn
import textura_plots as tplot
import textura_triagem as ttri

try:
    import textura_stats as tst
except ImportError:
    tst = None

MODOS_DEDUPE = (
    "nenhuma", "candidatos", "contexto", "obra_termo",
    "ocorrencia", "ocorrencia_termo",
)

RELACOES_VALIDAS = sorted(
    set(tn.RELACOES_NUCLEARES) | set(ttri.RELACOES_NAO_NUCLEARES) | {"obliqua"}
)
COLS_EDITAVEIS = [
    "relacao_sintactica", "nuclear", "polaridade", "eixo",
    "dominio", "motivo_exclusao",
]
# Folhas que a fase 2 (re)escreve. Qualquer outra folha do Excel de
# entrada é copiada tal qual (valores, formatação, validações).
FOLHAS_GERADAS_FASE2 = frozenset({
    "1_Resumo", "2_Frequencias", "3_Testes", "4_Sintaxe", "5_Coocorrencia",
    "8_Concordancia", "8_Concordancia_Hits",
    "9_Associacao", "12_Formas", "13_Regressao", "13_Ajuste_modelo",
    "15_Perfis", "0_Avisos", "6_Graficos",
    "6_Graficos_nuvem", "6_Graficos_barras", "6_Graficos_familias",
    "16_Kappa", "curadoria_fluxo",
})
NUVEM_RANDOM_STATE = 20260725
ROTULO_NAO_ADJUDICADA = "não adjudicada"
ROTULO_NAO_ADJUDICADO = "não adjudicado"
CAMPOS_LEXICO_OPCIONAIS = (
    ("polaridade", ROTULO_NAO_ADJUDICADA, "polaridade"),
    ("eixo", ROTULO_NAO_ADJUDICADO, "eixo"),
)
_TRUE = {"true", "1", "sim", "yes"}
_FALSE = {"false", "0", "nao", "não", "no"}
FRASE_LEGENDA_EIXOS = (
    "A frequência não equivale à força de associação ao núcleo; as medidas de "
    "associação constam do Apêndice."
)


def escrever_legenda_eixos(
        destino: Path, *, n: int, cauda: dict, conjunto: str) -> None:
    """Legenda pronta para o Word (UTF-8, português)."""
    linhas = [
        "Frequência das famílias lexicais realizadas junto de textur*, "
        "agrupadas por eixo semântico de curadoria.",
        "",
        "Unidade de contagem: ocorrência nuclear (atribuição genuína) e "
        "documento distinto (n_documentos).",
        f"Conjunto: atribuições nucleares. {conjunto} N={n}.",
        "",
        "Termos colapsados em «outros» por eixo:",
    ]
    if cauda:
        for eixo, termos in cauda.items():
            linhas.append(f"  {eixo}: {', '.join(termos)}")
    else:
        linhas.append("  (nenhum)")
    linhas.extend(["", FRASE_LEGENDA_EIXOS, ""])
    Path(destino).write_text("\n".join(linhas), encoding="utf-8")


def tabela_curadoria_fluxo(brutas, nuc, col_doc, cura_map) -> tuple[pd.DataFrame, str]:
    """Fluxo bruto → genuínas para o apêndice. Não adivinha contagens."""
    cols = (
        "canonical_term", "eixo", "coocorrencias_brutas",
        "atribuicoes_genuinas", "n_documentos", "decisao", "motivo",
    )
    aviso = ""
    termos = (
        sorted(nuc["canonical_term"].astype(str).unique())
        if len(nuc) and "canonical_term" in nuc.columns else []
    )
    brutas_n: dict | None
    if brutas is None or "canonical_term" not in getattr(brutas, "columns", []):
        aviso = (
            "coocorrencias_brutas: frame pré-adjudicação sem coluna "
            "canonical_term; coluna deixada vazia."
        )
        brutas_n = None
    else:
        brutas_n = (
            brutas.groupby(brutas["canonical_term"].astype(str))
            .size()
            .to_dict()
        )
    genu = (
        nuc.groupby(nuc["canonical_term"].astype(str)).size().to_dict()
        if len(nuc) and "canonical_term" in nuc.columns else {}
    )
    docs = (
        nuc.groupby(nuc["canonical_term"].astype(str))[col_doc].nunique().to_dict()
        if len(nuc) and col_doc in getattr(nuc, "columns", []) else {}
    )
    rows = []
    for t in termos:
        info = cura_map.get(t) or {
            "eixo": "por_classificar",
            "decisao": "por_classificar",
            "motivo": "",
        }
        rows.append({
            "canonical_term": t,
            "eixo": info.get("eixo", "por_classificar"),
            "coocorrencias_brutas": (
                "" if brutas_n is None else int(brutas_n.get(t, 0))
            ),
            "atribuicoes_genuinas": int(genu.get(t, 0)),
            "n_documentos": int(docs.get(t, 0)),
            "decisao": info.get("decisao", "por_classificar"),
            "motivo": info.get("motivo", ""),
        })
    return pd.DataFrame(rows, columns=list(cols)), aviso


def _cfg_consola():
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _meta_linha(unidade: str, n: int) -> pd.DataFrame:
    return pd.DataFrame({"unidade": [unidade], "N": [n]})


def _escrever_folha(
    xw,
    nome: str,
    df: pd.DataFrame,
    unidade: str,
    n: int,
    *,
    com_meta: bool = True,
):
    """Escreve folha estatística.

    ``com_meta=False`` para folhas de concordância reutilizáveis (sem banner
    ``unidade``/``N`` nas primeiras linhas — esse banner partia
    ``pd.read_excel`` em re-análises).
    """
    if not com_meta:
        if df is not None and len(df):
            df.to_excel(xw, sheet_name=nome, index=False)
        return
    meta = _meta_linha(unidade, n)
    meta.to_excel(xw, sheet_name=nome, index=False, startrow=0)
    if df is not None and len(df):
        df.to_excel(xw, sheet_name=nome, index=False, startrow=3)


def _parece_banner_meta(cols) -> bool:
    c = [str(x).strip().lower() for x in list(cols)]
    return len(c) >= 2 and c[0] == "unidade" and c[1] == "n"


def ler_folha_concordancia(xlsx, sheet: str = "8_Concordancia") -> pd.DataFrame:
    """Lê concordância fase 1 ou export de análise (com banner meta legado)."""
    conc = pd.read_excel(xlsx, sheet_name=sheet)
    if "relacao_sintactica" in conc.columns:
        return conc
    if _parece_banner_meta(conc.columns):
        # Export antigo de textura_analise: meta nas linhas 0–2, cabeçalho na 3
        conc2 = pd.read_excel(xlsx, sheet_name=sheet, header=3)
        if "relacao_sintactica" in conc2.columns:
            print(
                f"  [aviso] {sheet}: formato de saída de análise (banner meta); "
                "a ler cabeçalho na linha 4. Prefira o Excel *_near* / "
                "*_revisto* como entrada da fase 2.",
                flush=True,
            )
            return conc2
    return conc


def _sugerir_entrada_revisao(xlsx: Path) -> str:
    """Se o utilizador passou *_analise.xlsx, apontar o Excel de revisão."""
    nome = xlsx.name
    low = nome.lower()
    if "_analise" not in low:
        return ""
    stem = re.sub(r"_analise(\.xlsx)?$", "", nome, flags=re.I)
    if not stem.lower().endswith(".xlsx"):
        cand = xlsx.with_name(stem + ".xlsx")
    else:
        cand = xlsx.with_name(stem)
    if cand.exists():
        return str(cand)
    # Compósitaa_near_revisto_LR_v2_analise → …_v2.xlsx
    alt = xlsx.with_name(re.sub(r"_analise\.xlsx$", ".xlsx", nome, flags=re.I))
    if alt.exists() and alt != xlsx:
        return str(alt)
    return ""


def _nuclear_true_count(serie: pd.Series) -> int:
    return int(serie.map(
        lambda v: v is True or str(v).lower() in {"true", "1", "sim"}
    ).sum())


def comparar_espelho_hits(xlsx: Path) -> dict:
    """Compara ``8_Concordancia`` (edição) com o espelho ``8_Concordancia_Hits``."""
    xlsx = Path(xlsx)
    with pd.ExcelFile(xlsx) as xl:
        sheets = set(xl.sheet_names)
    if "8_Concordancia" not in sheets:
        return {"ok": False, "divergiu": False, "mensagem": "falta 8_Concordancia"}
    conc = ler_folha_concordancia(xlsx, "8_Concordancia")
    n_conc = len(conc)
    n_conc_nuc = (_nuclear_true_count(conc["nuclear"])
                  if "nuclear" in conc.columns else 0)
    n_hits = n_hits_nuc = None
    if "8_Concordancia_Hits" in sheets:
        hits = pd.read_excel(xlsx, sheet_name="8_Concordancia_Hits")
        n_hits = len(hits)
        n_hits_nuc = (_nuclear_true_count(hits["nuclear"])
                      if "nuclear" in hits.columns else 0)
    divergiu = (
        n_hits is not None
        and (n_hits != n_conc or n_hits_nuc != n_conc_nuc)
    )
    msg = (
        f"8_Concordancia linhas={n_conc} nuclear={n_conc_nuc}; "
        f"Hits linhas={n_hits} nuclear={n_hits_nuc}"
    )
    return {
        "ok": True, "divergiu": divergiu, "mensagem": msg,
        "n_conc": n_conc, "n_hits": n_hits,
        "n_conc_nuclear": n_conc_nuc, "n_hits_nuclear": n_hits_nuc,
        "conc": conc,
    }


def sincronizar_hits_com_concordancia(xlsx: Path, *, escrever: bool = True) -> dict:
    """Copia ``8_Concordancia`` → ``8_Concordancia_Hits`` (arquivo alinhado).

    A revisão edita só Concordancia; Hits é espelho e pode ficar velha.
    ``escrever=False`` só compara (relatório / ``--so-leitura``).
    """
    est = comparar_espelho_hits(xlsx)
    if not est.get("ok"):
        return est
    if est.get("divergiu"):
        print(
            "ATENCAO: 8_Concordancia e 8_Concordancia_Hits divergem "
            f"({est['mensagem']}). A folha editada (Concordancia) e a fonte. "
            + ("Hits vai ser substituida." if escrever else "Modo so-leitura: Hits nao e escrita."),
            flush=True,
        )
    if not escrever:
        est["n_hits_antes"] = est.get("n_hits_nuclear")
        est["n_hits_depois"] = est.get("n_conc_nuclear")
        return est
    conc = est["conc"]
    n_conc_nuc = est["n_conc_nuclear"]
    n_antes = est.get("n_hits_nuclear")
    with pd.ExcelWriter(
        xlsx, engine="openpyxl", mode="a", if_sheet_exists="replace",
    ) as xw:
        conc.to_excel(xw, sheet_name="8_Concordancia_Hits", index=False)
    return {
        "ok": True,
        "divergiu": est["divergiu"],
        "n_conc": n_conc_nuc,
        "n_hits_antes": n_antes,
        "n_hits_depois": n_conc_nuc,
        "mensagem": (
            f"8_Concordancia_Hits <- 8_Concordancia "
            f"(nuclear True: {n_antes} -> {n_conc_nuc})"
            if n_antes is not None else
            f"8_Concordancia_Hits criada (nuclear True={n_conc_nuc})"
        ),
    }


def ler_plano_a_priori(caminho: Path) -> dict:
    """JSON ou YAML mínimo (``chave: valor`` por linha)."""
    caminho = Path(caminho)
    texto = caminho.read_text(encoding="utf-8")
    bruto = texto.lstrip()
    if caminho.suffix.lower() == ".json" or bruto.startswith("{"):
        dados = json.loads(texto)
        if not isinstance(dados, dict):
            raise SystemExit("--plano-a-priori: JSON tem de ser um objecto")
        return {str(k): v for k, v in dados.items()}
    out: dict = {}
    for linha in texto.splitlines():
        linha = linha.split("#", 1)[0].strip()
        if not linha or ":" not in linha:
            continue
        k, v = linha.split(":", 1)
        out[k.strip()] = v.strip().strip("'\"")
    return out


def aplicar_plano_a_priori(
    plano: dict,
    *,
    desduplicacao: str,
    nulo_polaridade: str,
    cooc_unidade: str,
    relacoes: list[str] | None,
) -> tuple[str, str, str, list[str] | None, list[str]]:
    """O plano prevalece. Devolve valores + logs de divergência CLI vs plano."""
    logs = []
    mapa = {
        "desduplicacao": "desduplicacao",
        "nulo_polaridade": "nulo_polaridade",
        "nulo-polaridade": "nulo_polaridade",
        "cooc_unidade": "cooc_unidade",
        "cooc-unidade": "cooc_unidade",
    }
    actual = {
        "desduplicacao": desduplicacao,
        "nulo_polaridade": nulo_polaridade,
        "cooc_unidade": cooc_unidade,
    }
    for k_plano, k_arg in mapa.items():
        if k_plano not in plano:
            continue
        novo = str(plano[k_plano]).strip()
        if not novo:
            continue
        if str(actual[k_arg]) != novo:
            logs.append(
                f"plano a priori: {k_arg} {actual[k_arg]!r} -> {novo!r} (plano ganha)"
            )
        actual[k_arg] = novo
    rel_plano = plano.get("relacao", plano.get("relacoes"))
    rels = relacoes
    if rel_plano not in (None, ""):
        if isinstance(rel_plano, str):
            novos = [p.strip() for p in rel_plano.split(",") if p.strip()]
        else:
            novos = [str(p).strip() for p in rel_plano if str(p).strip()]
        if list(rels or []) != novos:
            logs.append(
                f"plano a priori: relacoes {list(rels or [])} -> {novos} (plano ganha)"
            )
        rels = novos or None
    return (
        actual["desduplicacao"], actual["nulo_polaridade"],
        actual["cooc_unidade"], rels, logs,
    )


def cohen_kappa(y1, y2) -> float:
    """κ de Cohen (rótulos categóricos; sem dependências extra)."""
    a = pd.Series(list(y1)).astype(str)
    b = pd.Series(list(y2)).astype(str)
    if len(a) != len(b) or not len(a):
        return float("nan")
    cats = sorted(set(a) | set(b))
    tab = pd.crosstab(a, b).reindex(index=cats, columns=cats, fill_value=0)
    n = float(tab.values.sum())
    if n <= 0:
        return float("nan")
    po = float(np.trace(tab.values)) / n
    p1 = tab.sum(axis=1).to_numpy(dtype=float) / n
    p2 = tab.sum(axis=0).to_numpy(dtype=float) / n
    pe = float(p1 @ p2)
    if pe >= 1.0:
        return 1.0 if po >= 1.0 else 0.0
    return float((po - pe) / (1.0 - pe))


def avaliar_kappa_cego(conc: pd.DataFrame, caminho: Path) -> tuple[pd.DataFrame, list[str]]:
    """Cruza a concordância com um segundo revisor (Excel ou JSON) por ``hit_key``."""
    caminho = Path(caminho)
    logs = []
    if not caminho.is_file():
        raise SystemExit(f"--kappa-cego: ficheiro inexistente: {caminho}")
    if caminho.suffix.lower() == ".json":
        bruto = json.loads(caminho.read_text(encoding="utf-8"))
        cego = pd.DataFrame(bruto if isinstance(bruto, list) else bruto.get("hits", []))
    else:
        try:
            cego = pd.read_excel(caminho, sheet_name="8_Concordancia")
        except ValueError:
            cego = pd.read_excel(caminho, sheet_name=0)
    if "hit_key" not in conc.columns or "hit_key" not in cego.columns:
        logs.append("kappa cego: falta hit_key — cruzamento impossivel")
        return pd.DataFrame(), logs
    a = conc[["hit_key"]].copy()
    a["nuclear_a"] = (conc["nuclear"].map(
        lambda v: v is True or str(v).lower() in _TRUE)
        if "nuclear" in conc.columns else False)
    a["relacao_a"] = (conc["relacao_sintactica"].astype(str)
                      if "relacao_sintactica" in conc.columns else "")
    b = cego[["hit_key"]].copy()
    b["nuclear_b"] = (cego["nuclear"].map(
        lambda v: v is True or str(v).lower() in _TRUE)
        if "nuclear" in cego.columns else False)
    b["relacao_b"] = (cego["relacao_sintactica"].astype(str)
                      if "relacao_sintactica" in cego.columns else "")
    m = a.merge(b, on="hit_key", how="inner")
    n_par = len(m)
    logs.append(
        f"kappa cego: {n_par} hits emparelhados "
        f"(A={len(a)} B={len(b)} ficheiro={caminho.name})"
    )
    if not n_par:
        return pd.DataFrame(), logs
    kn = cohen_kappa(m["nuclear_a"], m["nuclear_b"])
    kr = cohen_kappa(m["relacao_a"], m["relacao_b"])
    tab = pd.DataFrame({
        "campo": ["nuclear", "relacao_sintactica"],
        "n_emparelhados": [n_par, n_par],
        "kappa": [round(kn, 4) if kn == kn else "",
                  round(kr, 4) if kr == kr else ""],
        "acordo": [
            float((m["nuclear_a"] == m["nuclear_b"]).mean()),
            float((m["relacao_a"] == m["relacao_b"]).mean()),
        ],
    })
    return tab, logs


def cramers_v(tab: pd.DataFrame) -> float:
    if tab.size == 0 or min(tab.shape) < 2:
        return float("nan")
    chi2 = stats.chi2_contingency(tab.values)[0]
    n = tab.values.sum()
    r, k = tab.shape
    return float(np.sqrt(chi2 / (n * min(r - 1, k - 1)))) if n else float("nan")


def colinear_deterministica(df, a: str, b: str) -> bool:
    """True se V de Cramer ~= 1 (uma variavel funcao da outra)."""
    if a not in df.columns or b not in df.columns:
        return False
    tab = pd.crosstab(df[a], df[b])
    if tab.empty:
        return False
    # uma unica celula nao-nula por linha
    if (tab.gt(0).sum(axis=1) <= 1).all():
        return True
    v = cramers_v(tab)
    return v == v and v >= 0.999


def _serie_nao_vazia(serie: pd.Series) -> pd.Series:
    s = serie.astype(str).str.strip()
    return ~(
        serie.isna()
        | s.eq("")
        | s.str.lower().isin({"nan", "none", "nat"})
    )


def dv_deterministica_de_canonical(df: pd.DataFrame, col: str) -> tuple[bool, float]:
    """True se ``col`` é função de ``canonical_term`` (V=1 ou 1 valor)."""
    if col not in df.columns or "canonical_term" not in df.columns:
        return False, float("nan")
    mask = _serie_nao_vazia(df[col])
    sub = df.loc[mask, ["canonical_term", col]]
    if sub.empty:
        return False, float("nan")
    n_obs = int(sub[col].astype(str).str.strip().str.lower().nunique())
    if n_obs <= 1:
        return True, 1.0
    if colinear_deterministica(sub, "canonical_term", col):
        tab = pd.crosstab(sub["canonical_term"], sub[col])
        v = cramers_v(tab)
        return True, (1.0 if v != v else float(v))
    return False, float("nan")


def _linha_teste_omitido(familia: str, col: str, v: float) -> dict:
    nome = "polarity" if col == "polaridade" else col
    vtxt = "1" if (v != v or v >= 0.999) else f"{v:.3f}"
    return {
        "familia": familia,
        "teste": (
            f"omitted: {nome} is a deterministic function of "
            f"canonical_term (Cramér's V = {vtxt})"
        ),
        "estatistica": "",
        "dimensao_efeito": "",
        "p": float("nan"),
        "metodo": "",
    }


def _aviso_teste_omitido(familia: str, col: str) -> str:
    return (
        f"teste {familia} omitido: {col} é função determinística de "
        f"canonical_term (Cramér's V = 1). Não é falha; o teste não se aplica."
    )


def _parse_bool_override(v):
    """TRUE/FALSE humano; None se vazio ou marca de revisor (ex. LR)."""
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, bool):
        return v
    s = str(v).strip().lower()
    if s in _TRUE:
        return True
    if s in _FALSE:
        return False
    return None


def aplicar_override_revisto(conc: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """``revisto_por_humano`` TRUE/FALSE prevalece sobre ``nuclear``."""
    if "revisto_por_humano" not in conc.columns or "nuclear" not in conc.columns:
        return conc, []
    out = conc.copy()
    logs = []
    for i, row in out.iterrows():
        humano = _parse_bool_override(row.get("revisto_por_humano"))
        if humano is None:
            continue
        computado = bool(
            row.get("nuclear") is True
            or str(row.get("nuclear")).lower() in _TRUE
        )
        out.at[i, "nuclear"] = humano
        hit = row.get("hit_key", i)
        logs.append(
            f"override revisto_por_humano: hit_key={hit} "
            f"nuclear_computado={computado} nuclear_humano={humano}"
        )
    return out, logs


def gini_index(values) -> float:
    x = np.asarray(list(values), dtype=float)
    x = x[np.isfinite(x)]
    if x.size == 0 or float(x.sum()) <= 0:
        return float("nan")
    x = np.sort(x)
    n = x.size
    return float((2.0 * np.sum(np.arange(1, n + 1) * x)) / (n * x.sum())
                 - (n + 1) / n)


def _mascara_lexico_vazio(serie: pd.Series) -> pd.Series:
    return ~_serie_nao_vazia(serie)


def _aviso_campo_nao_adjudicado(df: pd.DataFrame, col: str) -> str | None:
    if col not in df.columns or "canonical_term" not in df.columns:
        return None
    vazias = _mascara_lexico_vazio(df[col])
    if not bool(vazias.any()):
        return None
    cnt = (df.loc[vazias, "canonical_term"].astype(str)
           .value_counts())
    partes = " · ".join(f"{t} {int(n)}" for t, n in cnt.items())
    return (
        f"{col} não adjudicada no léxico ({int(vazias.sum())} hits; "
        f"{int(cnt.size)} termos): {partes}"
    )


def _preencher_lexico_vazio(df: pd.DataFrame, col: str, rotulo: str) -> pd.DataFrame:
    d = df.copy()
    if col not in d.columns:
        return d
    d.loc[_mascara_lexico_vazio(d[col]), col] = rotulo
    return d


def _dispersao_por_chave(df: pd.DataFrame, keys: list[str],
                         col_doc: str) -> pd.DataFrame:
    if df.empty or col_doc not in df.columns:
        return pd.DataFrame()
    rows = []
    for key_vals, g in df.groupby(keys, dropna=False):
        if not isinstance(key_vals, tuple):
            key_vals = (key_vals,)
        n = len(g)
        por_doc = g[col_doc].astype(str).value_counts()
        n_docs = int(por_doc.size)
        top_n = int(por_doc.iloc[0]) if n_docs else 0
        rec = {k: v for k, v in zip(keys, key_vals)}
        rec.update({
            "n_documentos": n_docs,
            "hits_por_documento": (
                round(n / n_docs, 4) if n_docs else float("nan")),
            "share_doc_maximo": (
                round(top_n / n, 4) if n else float("nan")),
            "doc_dominante": (por_doc.index[0] if n_docs else ""),
        })
        rows.append(rec)
    return pd.DataFrame(rows)


def _df_pesos_grafico(freq: dict, *, col_forma: str = "forma",
                      col_peso: str = "peso",
                      ponderacao: str | None = None,
                      familia: dict | None = None) -> pd.DataFrame:
    """Tabela de auditoria: pares exactos passados ao renderer."""
    rows = []
    for k, v in (freq or {}).items():
        rec = {col_forma: k, col_peso: int(v)}
        if ponderacao is not None:
            rec["ponderacao"] = ponderacao
        if familia is not None:
            rec["canonical_term"] = familia.get(k, "")
        rows.append(rec)
    if not rows:
        cols = [c for c in ("ponderacao", col_forma, "canonical_term", col_peso)
                if c == col_forma or c == col_peso
                or (c == "ponderacao" and ponderacao is not None)
                or (c == "canonical_term" and familia is not None)]
        return pd.DataFrame(columns=cols)
    df = pd.DataFrame(rows)
    by = [c for c in ("ponderacao", col_peso, col_forma) if c in df.columns]
    asc = [c != col_peso for c in by]
    return df.sort_values(by, ascending=asc).reset_index(drop=True)


def mapa_forma_familia(formas: pd.DataFrame) -> dict:
    """``matched_form`` → ``canonical_term`` (primeira ocorrência se houver conflito)."""
    if formas is None or len(formas) == 0:
        return {}
    if "matched_form" not in formas.columns or "canonical_term" not in formas.columns:
        return {}
    out = {}
    for form, g in formas.groupby("matched_form", dropna=False):
        out[form] = str(g["canonical_term"].iloc[0])
    return out


def freq_nuvem_de_formas(formas: pd.DataFrame,
                         ponderacao: str = "hits") -> dict:
    """Pesos da nuvem a partir de ``12_Formas``.

    ``ponderacao='hits'`` → ``n``; ``ponderacao='docs'`` → ``n_documentos``.
    """
    if formas is None or len(formas) == 0 or "matched_form" not in formas.columns:
        return {}
    modo = (ponderacao or "hits").strip().lower()
    if modo in {"docs", "documentos", "n_documentos"}:
        col = "n_documentos" if "n_documentos" in formas.columns else (
            "obras" if "obras" in formas.columns else None)
        if col is None:
            return {}
        return (formas.groupby("matched_form", dropna=False)[col]
                .max().fillna(0).astype(int).to_dict())
    col_n = "n" if "n" in formas.columns else None
    if col_n is None:
        return {}
    return (formas.groupby("matched_form", dropna=False)[col_n]
            .sum().astype(int).to_dict())


def _nota_teto(mostrados: int, total: int, *,
               extra: str = "", cap=None) -> str:
    if cap is None or mostrados >= total:
        return ""
    base = f"showing {mostrados} of {total}"
    if extra:
        base = f"{base} ({extra})"
    return f"{base}; link cap = {cap}"


def _anotar_instrucoes_adjudicacao(wb) -> None:
    if "0_Instrucoes" not in wb.sheetnames:
        return
    ws = wb["0_Instrucoes"]
    chaves = {
        str(ws.cell(r, 1).value).strip()
        for r in range(2, ws.max_row + 1)
        if ws.cell(r, 1).value is not None
    }
    pares = [
        ("adjudicacao_nuclear",
         "A coluna nuclear é o veredicto automático da fase 1. "
         "revisto_por_humano, quando TRUE ou FALSE, prevalece sobre nuclear; "
         "vazio = usar nuclear. Cada override é registado em 0_Avisos "
         "(hit_key, valor computado, valor humano)."),
        ("coluna_adjudicacao",
         "Adjudicação humana: revisto_por_humano (TRUE/FALSE). "
         "nuclear permanece o juízo automático."),
    ]
    for chave, valor in pares:
        if chave in chaves:
            continue
        r = ws.max_row + 1
        ws.cell(r, 1).value = chave
        ws.cell(r, 2).value = valor


def teste_contingencia(tab: pd.DataFrame, n_perm: int = 20000, semente=20260725,
                       min_validas: int = 200):
    """χ² ou Monte Carlo se alguma celula esperada < 5.

    Simulações degeneradas (expected zero) são saltadas; o denominador de p
    usa só o número efectivo de draws válidos. Se esse número ficar abaixo
    de ``min_validas``, cai-se para o teste exacto de Fisher (2×2) com aviso,
    em vez de reportar um p com resolução muito inferior ao ``n_perm`` nominal.
    """
    if tab.shape[0] < 2 or tab.shape[1] < 2:
        return {"metodo": "inaplicavel", "p": float("nan"),
                "estatistica": "", "cramer_v": float("nan"),
                "n_validas": 0}
    chi2, p_asym, gl, esperado = stats.chi2_contingency(tab.values)
    v = cramers_v(tab)
    if (esperado < 5).any():
        rng = np.random.default_rng(semente)
        obs = tab.values
        n = int(obs.sum())
        rprop = obs.sum(1) / n
        cprop = obs.sum(0) / n
        estat = chi2
        maiores = 0
        validas = 0
        max_tentativas = max(n_perm * 20, n_perm + 100)
        tentativas = 0
        while validas < n_perm and tentativas < max_tentativas:
            tentativas += 1
            flat = rng.multinomial(n, np.outer(rprop, cprop).ravel())
            sim = flat.reshape(obs.shape)
            try:
                c2 = stats.chi2_contingency(sim, correction=False)[0]
            except ValueError:
                continue
            validas += 1
            if c2 >= estat - 1e-12:
                maiores += 1

        # Poucas sims válidas → Fisher exacto em tabelas 2×2 (resolução honesta)
        if validas < min_validas and obs.shape == (2, 2):
            p_fish = float(stats.fisher_exact(obs).pvalue)
            print(
                f"AVISO: teste_contingencia — só {validas} sims MC válidas "
                f"(pedido {n_perm}); a usar Fisher exacto.",
                flush=True,
            )
            return {
                "metodo": (
                    f"Fisher exacto (fallback; {validas} MC validas/"
                    f"{tentativas} tentativas < {min_validas})"
                ),
                "p": p_fish,
                "estatistica": f"χ²({gl}) = {chi2:.3f}; Fisher",
                "cramer_v": round(v, 4),
                "n_validas": validas,
            }
        if validas == 0:
            return {"metodo": "Monte Carlo (inaplicavel: 0 sims validas)",
                    "p": float("nan"),
                    "estatistica": f"χ²({gl}) = {chi2:.3f}",
                    "cramer_v": round(v, 4),
                    "n_validas": 0}
        p = (maiores + 1) / (validas + 1)
        metodo = (
            f"Monte Carlo ({validas} perm. validas/{tentativas} tentativas; "
            f"celula esperada < 5)"
        )
        return {"metodo": metodo, "p": p,
                "estatistica": f"χ²({gl}) = {chi2:.3f}",
                "cramer_v": round(v, 4),
                "n_validas": validas}
    return {"metodo": "χ² assintotico", "p": p_asym,
            "estatistica": f"χ²({gl}) = {chi2:.3f}",
            "cramer_v": round(v, 4),
            "n_validas": None}


def logdice(o11: int, o12: int, n_janelas_no: int) -> float:
    """logDice com denominador = janelas do no (A5)."""
    dice = 2 * o11 / (n_janelas_no + o11 + o12) if (
        n_janelas_no + o11 + o12) else float("nan")
    return 14 + math.log2(dice) if dice and dice > 0 else float("nan")


def polaridade_nulo_banda(res_nuc: pd.DataFrame, hits_banda: dict | None,
                          campo: dict) -> float:
    """Proporcao esperada de estabilidade na banda (opcao 2, omissao)."""
    if not hits_banda:
        # fallback: proporcao de padroes E no lexico
        n_e = sum(1 for t in campo if t in tlex.POLO_ESTABILIDADE)
        return n_e / max(len(campo), 1)
    tot_e = sum(hits_banda.get(t, 0) for t in campo
                if t in tlex.POLO_ESTABILIDADE)
    tot = sum(hits_banda.get(t, 0) for t in campo)
    if tot == 0:
        n_e = sum(1 for t in campo if t in tlex.POLO_ESTABILIDADE)
        return n_e / max(len(campo), 1)
    return tot_e / tot


def validar_fase1(xlsx: Path, *, estrito: bool = False) -> dict:
    xlsx = Path(xlsx)
    with pd.ExcelFile(xlsx) as xl:
        sheets = list(xl.sheet_names)
        if "0_Instrucoes" not in sheets:
            raise SystemExit(
                "Ficheiro sem folha 0_Instrucoes — nao passou pela fase 1 "
                "(extraccao para revisao). Corra textura_near.py primeiro.")
        if "8_Concordancia" not in sheets:
            raise SystemExit("Falta a folha 8_Concordancia.")
        eh_saida_analise = (
            "1_Resumo" in sheets
            or "3_Testes" in sheets
            or "_analise" in xlsx.name.lower()
        )
        if eh_saida_analise:
            alt = _sugerir_entrada_revisao(xlsx)
            print(
                "  [aviso] Este Excel parece saída da fase 2 (*_analise / "
                "folhas 1_Resumo…). A entrada correcta é o Excel de revisão "
                "(*_near* / *_revisto*) com 8_Concordancia editável.",
                flush=True,
            )
            if alt:
                print(f"  [aviso] Use em vez disso:\n    {alt}", flush=True)
        conc = ler_folha_concordancia(xl, "8_Concordancia")
        meta = {}
        try:
            inst = pd.read_excel(xl, sheet_name="0_Instrucoes")
            if "chave" in inst.columns and "valor" in inst.columns:
                meta = dict(zip(inst["chave"].astype(str), inst["valor"]))
        except Exception:
            pass
    for c in ("relacao_sintactica", "nuclear", "canonical_term"):
        if c not in conc.columns:
            alt = _sugerir_entrada_revisao(xlsx)
            tip = (
                f"\nProvável causa: seleccionou um *_analise.xlsx (saída da "
                f"fase 2), não o Excel revisto.\n"
                f"Use: {alt or xlsx.with_name(xlsx.stem.replace('_analise', '') + '.xlsx')}"
            )
            raise SystemExit(f"Coluna obrigatoria em falta: {c}.{tip}")
    # validar taxonomia
    for i, row in conc.iterrows():
        rel = str(row.get("relacao_sintactica", ""))
        if rel and rel not in RELACOES_VALIDAS and rel != "nan":
            raise SystemExit(
                f"Valor invalido em relacao_sintactica linha {i+2}: {rel!r}. "
                f"Admissiveis: {RELACOES_VALIDAS}")

    # R95 — checklist de revisão (avisos; erros só com --estrito)
    checklist = None
    try:
        import textura_triagem as ttri
        checklist = ttri.checklist_revisao(conc)
        for a in checklist.get("avisos") or []:
            print(f"  [revisão] aviso: {a}", flush=True)
        for e in checklist.get("erros") or []:
            print(f"  [revisão] ERRO: {e}", flush=True)
        print(
            f"  [revisão] checklist score={checklist.get('score')} "
            f"nuclear={checklist.get('n_nuclear')}/{checklist.get('n')}",
            flush=True,
        )
        if estrito and not checklist.get("ok"):
            raise SystemExit(
                "Checklist de revisão falhou (--estrito). "
                "Corrija 8_Concordancia ou use textura_doctor.py."
            )
    except SystemExit:
        raise
    except Exception as exc:  # noqa: BLE001
        print(f"  [revisão] checklist indisponível: {exc}", flush=True)

    return {
        "conc": conc, "meta": meta, "sheets": sheets, "checklist": checklist,
    }


def comparar_revisao(conc: pd.DataFrame, meta: dict) -> dict:
    """Taxa de concordancia automatico vs humano (se houver snapshot)."""
    n = len(conc)
    alteradas = 0
    if "revisto_por_humano" in conc.columns:
        alteradas = int(conc["revisto_por_humano"].astype(str).str.strip()
                        .replace("", np.nan).notna().sum())
    return {
        "linhas": n,
        "marcadas_revisto_por_humano": alteradas,
        "taxa_marcacao": round(alteradas / n, 4) if n else 0,
    }


def _flag_candidato(serie: pd.Series) -> pd.Series:
    s = serie.astype(str).str.strip()
    return s.ne("") & ~s.str.lower().isin({"nan", "none", "nat", "false", "0"})


def aplicar_desduplicacao(nuc: pd.DataFrame, modo: str,
                          col_doc: str) -> tuple[pd.DataFrame, str]:
    """Desduplicação opcional na fase 2 (a revisão manual é a fonte de verdade)."""
    modo = (modo or "nenhuma").strip().lower()
    if modo not in MODOS_DEDUPE:
        raise ValueError(
            f"desduplicacao desconhecida: {modo!r}; "
            f"escolha entre {', '.join(MODOS_DEDUPE)}")
    if modo == "nenhuma" or len(nuc) == 0:
        return nuc.copy(), "nenhuma"
    if modo == "ocorrencia":
        if "texture_occurrence_id" not in nuc.columns:
            return nuc.copy(), "ocorrencia(col em falta→nenhuma)"
        return (nuc.drop_duplicates(subset=["texture_occurrence_id"])
                .copy(), "ocorrencia")
    if modo == "ocorrencia_termo":
        cols = [c for c in ("texture_occurrence_id", "canonical_term")
                if c in nuc.columns]
        if len(cols) < 2:
            return nuc.copy(), "ocorrencia_termo(cols em falta→nenhuma)"
        return nuc.drop_duplicates(subset=cols).copy(), "ocorrencia_termo"
    if modo == "obra_termo":
        cols = [c for c in (col_doc, "canonical_term") if c in nuc.columns]
        if len(cols) < 2:
            return nuc.copy(), "obra_termo(cols em falta→nenhuma)"
        return nuc.drop_duplicates(subset=cols).copy(), "obra_termo"
    if modo == "contexto":
        if "contexto" not in nuc.columns:
            return nuc.copy(), "contexto(indisponivel→nenhuma)"
        chave = nuc["contexto"].astype(str).str.strip()
        return (nuc.assign(_k=chave).drop_duplicates(subset=["_k"])
                .drop(columns=["_k"]).copy(), "contexto")
    # candidatos: entre linhas assinaladas na fase 1, fica 1 por contexto;
    # se não houver flag, equivale a desduplicar por contexto exacto.
    if "candidato_duplicado" in nuc.columns and "contexto" in nuc.columns:
        flagged = _flag_candidato(nuc["candidato_duplicado"])
        if flagged.any():
            keep = []
            visto: set[str] = set()
            for i, row in nuc.iterrows():
                if not bool(flagged.loc[i]):
                    keep.append(i)
                    continue
                k = str(row.get("contexto", "")).strip()
                if k in visto:
                    continue
                visto.add(k)
                keep.append(i)
            return nuc.loc[keep].copy(), "candidatos"
    return aplicar_desduplicacao(nuc, "contexto", col_doc)[0], "candidatos→contexto"


def analisar(xlsx: Path, saida: Path | None = None,
             nulo_polaridade: str = "banda",
             cooc_unidade: str = "obra",
             desduplicacao: str = "nenhuma",
             legendas: Path | None = None,
             relacoes: list[str] | None = None,
             estrito: bool = False,
             plano_a_priori: Path | None = None,
             kappa_cego: Path | None = None,
             lexico: Path | None = None,
             modo_tese: bool = False) -> int:
    _cfg_consola()
    logs_plano = []
    if plano_a_priori:
        plano = ler_plano_a_priori(plano_a_priori)
        (desduplicacao, nulo_polaridade, cooc_unidade, relacoes,
         logs_plano) = aplicar_plano_a_priori(
            plano, desduplicacao=desduplicacao,
            nulo_polaridade=nulo_polaridade, cooc_unidade=cooc_unidade,
            relacoes=relacoes)
        for lg in logs_plano:
            print(f"  [plano] {lg}", flush=True)
    saida = saida or xlsx
    info = validar_fase1(xlsx, estrito=estrito)
    conc = info["conc"]
    meta = info["meta"]
    consulta = str(meta.get("consulta", meta.get("query", "")) or "")
    leg = tleg.carregar(legendas, consulta=consulta)

    # conjuntos
    conc["nuclear"] = conc["nuclear"].map(
        lambda v: v is True or str(v).lower() in {"true", "1", "sim"})
    conc, logs_override = aplicar_override_revisto(conc)
    rev = comparar_revisao(conc, meta)
    brutas = conc
    nuc = conc.loc[conc["nuclear"]].copy()
    if relacoes:
        rels = {r.strip().lower() for r in relacoes if r and r.strip()}
        if "relacao_sintactica" in nuc.columns and rels:
            antes = len(nuc)
            nuc = nuc.loc[
                nuc["relacao_sintactica"].astype(str).str.strip().str.lower()
                .isin(rels)
            ].copy()
            print(f"Filtro --relacao {sorted(rels)}: {antes} → {len(nuc)}",
                  flush=True)
    col_doc = "doc_id" if "doc_id" in nuc.columns else "caminho_ficheiro"
    dedup, modo_dedupe = aplicar_desduplicacao(nuc, desduplicacao, col_doc)

    n_bruto = len(brutas)
    n_nuc = len(nuc)
    n_dedup = len(dedup)
    n_janelas_no = int(float(meta.get("janelas_kwic_processadas", n_bruto) or n_bruto))
    tot_near = float(meta.get("tot_near", 0) or 0)
    tot_banda = float(meta.get("tot_banda", 0) or 0)

    # Sempre recontar a partir das nucleares *actuais* (após revisão).
    # Os n_* em 0_Instrucoes são da extracção fase 1 e ficam obsoletos
    # quando o utilizador muda nuclear TRUE↔FALSE.
    if "texture_occurrence_id" in dedup.columns:
        n_occ_nuc = int(dedup["texture_occurrence_id"].nunique())
    elif "texture_occurrence_id" in nuc.columns:
        n_occ_nuc = int(nuc["texture_occurrence_id"].nunique())
    else:
        n_occ_nuc = n_dedup
    n_occ_meta = meta.get("n_ocorrencias_nucleares", meta.get("n_ocorrencias"))
    try:
        n_occ_meta_i = int(float(n_occ_meta)) if n_occ_meta not in (None, "") else None
    except (TypeError, ValueError):
        n_occ_meta_i = None
    if n_occ_meta_i is not None and n_occ_meta_i != n_occ_nuc:
        avisos_pre = (
            f"meta fase 1 tinha n_ocorrencias_nucleares={n_occ_meta_i}; "
            f"recontado após revisão = {n_occ_nuc} (usado nos resultados)."
        )
    else:
        avisos_pre = ""
    print(f"Fase 2 | brutas={n_bruto} nucleares(hits)={n_nuc} "
          f"dedup({modo_dedupe})={n_dedup} "
          f"ocorrencias_nucleares={n_occ_nuc} janelas_no={n_janelas_no}",
          flush=True)
    if avisos_pre:
        print(f"  [aviso] {avisos_pre}", flush=True)
    if legendas:
        print(f"Legendas: {legendas} | {tleg.resumo_titulos(leg)}", flush=True)

    # --- A9 colinearidade -------------------------------------------------
    avisos = []
    if avisos_pre:
        avisos.append(avisos_pre)
    avisos.extend(logs_override)
    avisos.extend(logs_plano)
    tab_kappa = pd.DataFrame()
    if kappa_cego:
        tab_kappa, logs_k = avaliar_kappa_cego(conc, Path(kappa_cego))
        avisos.extend(logs_k)
        for lg in logs_k:
            print(f"  [kappa] {lg}", flush=True)
    # Folha Hits desactualizada → sincronizar espelho no Excel de revisão
    if "8_Concordancia_Hits" in info.get("sheets", []):
        try:
            hits_df = pd.read_excel(xlsx, sheet_name="8_Concordancia_Hits")
            if "nuclear" in hits_df.columns:
                h_true = _nuclear_true_count(hits_df["nuclear"])
                if h_true != int(n_nuc):
                    sync = sincronizar_hits_com_concordancia(xlsx)
                    msg = sync.get("mensagem") or (
                        f"Hits sincronizada ({h_true} → {n_nuc})"
                    )
                    print(f"  [sync] {msg}", flush=True)
                    avisos.append(f"sincronizado: {msg}")
                    # folha Hits no workbook de entrada já está alinhada
                    info["sheets"] = list(pd.ExcelFile(xlsx).sheet_names)
        except Exception as exc:  # noqa: BLE001
            avisos.append(
                f"8_Concordancia_Hits desactualizada e sync falhou: {exc}"
            )
            print(f"  [aviso] sync Hits falhou: {exc}", flush=True)
    if colinear_deterministica(nuc, "canonical_term", "eixo"):
        avisos.append(
            "BLOQUEADO: eixo e funcao deterministica de canonical_term "
            "(V de Cramer = 1). Testes que cruzam eixo foram omitidos. "
            "Proposta de revisao do eixo: varied -> invariancia_diacronica "
            "(ou ambos); uniform -> homogeneidade_sincronica; "
            "static -> invariancia_diacronica.")
    if colinear_deterministica(nuc, "canonical_term", "polaridade"):
        avisos.append(
            "polaridade é função de canonical_term no léxico adjudicado "
            "(cada termo tem uma só polaridade). Esperado em classes "
            "mono-polares; não é falha da análise.")
    for col, _rotulo, _fam in CAMPOS_LEXICO_OPCIONAIS:
        msg_na = _aviso_campo_nao_adjudicado(nuc, col)
        if msg_na:
            avisos.append(msg_na)

    # --- 2_Frequencias (A2) ----------------------------------------------
    freq_tok = (nuc.groupby("canonical_term")
                .agg(ocorrencias_token=("canonical_term", "size"),
                     obras_com_ocorrencia=(col_doc, "nunique"))
                .sort_values("ocorrencias_token", ascending=False)
                .reset_index())
    disp_term = _dispersao_por_chave(nuc, ["canonical_term"], col_doc)
    if len(disp_term):
        freq_tok = freq_tok.merge(disp_term, on="canonical_term", how="left")
    # H, J, 1/D sobre ocorrencias_token
    cont_tok = nuc["canonical_term"].value_counts()
    S = int((cont_tok > 0).sum())
    H = tn.shannon(cont_tok.values) if S else float("nan")
    J = tn.pielou(cont_tok.values) if S else float("nan")
    invD = tn.simpson_inverso(cont_tok.values) if S else float("nan")

    # --- 3_Testes (A4, A8, A9) -------------------------------------------
    testes = []
    pol = dedup.replace({"polaridade": {"": np.nan}}).dropna(subset=["polaridade"])
    det_pol, v_pol = dv_deterministica_de_canonical(nuc, "polaridade")
    if det_pol:
        testes.append(_linha_teste_omitido("polaridade", "polaridade", v_pol))
        avisos.append(_aviso_teste_omitido("polaridade", "polaridade"))
    elif len(pol):
        n_est = int((pol["polaridade"] == "estabilidade").sum())
        n_tot = len(pol)
        # nulo
        hits_banda = {}
        if "hits_banda" in meta:
            try:
                hits_banda = eval(meta["hits_banda"], {"__builtins__": {}})
            except Exception:
                hits_banda = {}
        campo = {}
        if "campo_tipos" in meta:
            campo = {t: [] for t in str(meta["campo_tipos"]).split(",") if t}
        if nulo_polaridade == "lexico":
            p0 = sum(1 for t in campo if t in tlex.POLO_ESTABILIDADE) / max(
                len(campo), 1)
            rotulo_nulo = f"proporcional ao lexico (p0={p0:.3f})"
        else:
            p0 = polaridade_nulo_banda(nuc, hits_banda, campo or {
                t: [] for t in nuc["canonical_term"].unique()})
            rotulo_nulo = f"calibrado pela banda de referencia (p0={p0:.3f})"
        bt = stats.binomtest(n_est, n_tot, p0, alternative="two-sided")
        if tst:
            ic = tn.bootstrap_proporcao(
                (pol["polaridade"] == "estabilidade").values)
        else:
            ic = (float("nan"), float("nan"))
        testes.append({
            "familia": "polaridade",
            "teste": f"Binomial estabilidade vs nulo ({rotulo_nulo})",
            "estatistica": f"{n_est}/{n_tot} = {n_est/n_tot:.3f} "
                           f"[IC95% {ic[0]:.3f}-{ic[1]:.3f}]",
            "dimensao_efeito": f"prop={n_est/n_tot:.3f}; nulo={p0:.3f}",
            "p": bt.pvalue,
            "metodo": "binomial exacto",
        })
        # Bayes -> 3_Testes (A3)
        if tst:
            bf = tst.bayes_factor_proporcao(n_est, n_tot)
            testes.append({
                "familia": "polaridade",
                "teste": "Factor de Bayes (proporcao vs 0,5; prior Beta(1,1))",
                "estatistica": f"BF10={bf}",
                "dimensao_efeito": f"BF={bf}",
                "p": float("nan"),
                "metodo": "Bayes factor proporcao",
            })

    col_rel = "relacao_sintactica"
    if len(pol) and col_rel in pol.columns:
        if not colinear_deterministica(pol, col_rel, "polaridade"):
            tab = pd.crosstab(pol[col_rel], pol["polaridade"])
            r = teste_contingencia(tab)
            testes.append({
                "familia": "estrutura_sintactica",
                "teste": "relacao_sintactica × polaridade",
                "estatistica": r["estatistica"],
                "dimensao_efeito": f"V_Cramer={r['cramer_v']}",
                "p": r["p"],
                "metodo": r["metodo"],
            })
        else:
            avisos.append(
                "teste relação×polaridade omitido: tabela degenerada "
                "(só uma polaridade nas nucleares → colinearidade). "
                "Não é falha; o teste χ² não se aplica.")

    det_eixo, v_eixo = dv_deterministica_de_canonical(dedup, "eixo")
    if det_eixo:
        testes.append(_linha_teste_omitido("eixo", "eixo", v_eixo))
        avisos.append(_aviso_teste_omitido("eixo", "eixo"))
    elif "eixo" in dedup.columns and col_rel in dedup.columns:
        tab = pd.crosstab(dedup[col_rel], dedup["eixo"])
        r = teste_contingencia(tab)
        testes.append({
            "familia": "estrutura_sintactica",
            "teste": "relacao_sintactica × eixo",
            "estatistica": r["estatistica"],
            "dimensao_efeito": f"V_Cramer={r['cramer_v']}",
            "p": r["p"],
            "metodo": r["metodo"],
        })

    tst_df = pd.DataFrame(testes)
    if len(tst_df):
        # BH por familia (A14)
        tst_df["p_ajustado_BH"] = np.nan
        for fam, idx in tst_df.groupby("familia").groups.items():
            sub = tst_df.loc[idx, "p"].astype(float)
            mask = sub.notna()
            if mask.any():
                adj = tn.benjamini_hochberg(sub[mask].values)
                tst_df.loc[sub[mask].index, "p_ajustado_BH"] = adj

    # --- 4_Sintaxe sobre dedup -------------------------------------------
    sintaxe = (pd.crosstab(dedup["canonical_term"], dedup[col_rel])
               if len(dedup) else pd.DataFrame())

    # --- 5_Coocorrencia por obra (A12) -----------------------------------
    cooc = pd.DataFrame()
    if cooc_unidade == "obra" and len(nuc):
        tipos = sorted(nuc["canonical_term"].unique())
        mat = pd.DataFrame(0, index=tipos, columns=tipos, dtype=int)
        for _, g in nuc.groupby(col_doc):
            presentes = sorted(g["canonical_term"].unique())
            for a in presentes:
                for b in presentes:
                    mat.at[a, b] += 1
        # so escrever se houver off-diagonal
        if (mat.values.sum() - np.trace(mat.values)) > 0:
            cooc = mat
        else:
            avisos.append(
                "5_Coocorrencia omitida: matriz diagonal "
                f"(unidade={cooc_unidade}; nenhum par de tipos na mesma obra).")

    # --- 9_Associacao (A5, A6) -------------------------------------------
    # O11 sobre nucleares; O12 da meta se existir
    hits_banda = {}
    if "hits_banda" in meta:
        try:
            hits_banda = eval(meta["hits_banda"], {"__builtins__": {}})
        except Exception:
            hits_banda = {}
    filas = []
    for etq, o11 in Counter(nuc["canonical_term"]).items():
        o12 = int(hits_banda.get(etq, 0))
        r1 = tot_near or max(n_janelas_no * 3, 1)  # fallback
        r2 = tot_banda or r1
        m = tn.medidas_associacao(o11, o12, r1, r2, n_janelas_no)
        if not m:
            continue
        # corrigir logDice (A5)
        m["logDice"] = round(logdice(o11, o12, n_janelas_no), 3)
        m["canonical_term"] = etq
        m["obras"] = int(nuc.loc[nuc["canonical_term"] == etq, col_doc].nunique())
        # leitura (A6)
        orr = m.get("razao_possib", float("nan"))
        dp = m.get("DeltaP", float("nan"))
        m["leitura"] = (
            f"OR={orr}: o termo e {orr:.2f}x mais provavel na janela NEAR "
            f"do que na banda; ΔP={dp:.5f} (diferenca de probabilidades)."
        )
        m["r1_tokens_near"] = r1
        m["r2_tokens_banda"] = r2
        m["n_janelas_no"] = n_janelas_no
        m["formula_logDice"] = "14+log2(2*O11/(n_janelas_no+O11+O12))"
        # dispersao so se obras >= 5
        if m["obras"] < 5:
            m["DP_Gries"] = float("nan")
            m["D_Juilland"] = float("nan")
            m["partes_nao_nulas"] = m["obras"]
        else:
            partes = Counter(nuc.loc[nuc["canonical_term"] == etq, col_doc])
            tam = Counter(nuc[col_doc])
            m["DP_Gries"] = tn.dispersao_gries_dp(partes, tam)
            m["D_Juilland"] = tn.juilland_d(partes, max(len(tam), 1))
            m["partes_nao_nulas"] = int((np.array(list(partes.values())) > 0).sum())
        filas.append(m)
    assoc = pd.DataFrame(filas)
    if len(assoc):
        # ordenar por dimensao de efeito (OR), nao por G2
        assoc = assoc.sort_values("razao_possib", ascending=False)
        if "MI3" in assoc.columns:
            assoc = assoc.rename(columns={"MI3": "MI3_heuristica"})

    # --- 15_Perfis sobre dedup (A1, A11) ---------------------------------
    perfil = pd.DataFrame()
    if tst and len(dedup) >= 3:
        d = dedup.copy()
        # Excel fase 1 pode já ter termo_tipo e canonical_term — evitar
        # colunas duplicadas no rename (value_counts falha com DataFrame).
        if "canonical_term" in d.columns:
            d["termo_tipo"] = d["canonical_term"]
        elif "termo_tipo" not in d.columns:
            d["termo_tipo"] = d.get("matched_form", pd.Series(dtype=str))
        try:
            perfil, _ = tst.perfis_e_dendrograma(d, None, min_ocorr=1)
        except Exception as exc:
            avisos.append(f"15_Perfis omitido: {exc}")
            perfil = pd.DataFrame()

    # --- 13_Regressao: prever nuclear (A10) ------------------------------
    reg_tab, reg_aj = pd.DataFrame(), {}
    if tst and len(brutas) > 20:
        dreg = brutas.copy()
        dreg["polaridade"] = np.where(dreg["nuclear"], "estabilidade", "variabilidade")
        # abuso do alvo: usamos polaridade como proxy binario nuclear
        # melhor: criar coluna alvo_nuclear
        dreg = dreg.rename(columns={"polaridade": "_pol_old"})
        dreg["alvo_nuclear"] = np.where(dreg["nuclear"], "nuclear", "incidental")
        # so preditores nao circulares
        preds = [c for c in ("distancia", "lado", "censurado_esq", "censurado_dir")
                 if c in dreg.columns]
        if preds:
            # adaptar regressao para alvo custom
            tab, aj = tst.regressao_logistica(
                dreg.rename(columns={"alvo_nuclear": "polaridade"}),
                alvo="polaridade", positivo="nuclear",
                preditores=tuple(preds))
            # bloquear coeficientes absurdos (separacao)
            if len(tab) and (
                    (tab["coef_log_odds"].abs() > 10).any()
                    or (tab["erro_padrao"] > 100).any()):
                avisos.append(
                    "13_Regressao abortada: separacao completa "
                    "(|coef|>10 ou EP>100).")
                reg_tab, reg_aj = pd.DataFrame(), {"aviso": "separacao"}
            else:
                reg_tab, reg_aj = tab, aj

    # --- A3 formas por tipo (sem indices de riqueza) ---------------------
    formas = (nuc.groupby(["canonical_term", "matched_form"])
              .agg(n=("matched_form", "size"),
                   obras=(col_doc, "nunique"))
              .reset_index())
    disp_form = _dispersao_por_chave(
        nuc, ["canonical_term", "matched_form"], col_doc)
    if len(disp_form):
        formas = formas.merge(
            disp_form, on=["canonical_term", "matched_form"], how="left")

    # --- Graficos (B) ----------------------------------------------------
    base = saida.parent
    g_msgs = []
    leg_formas = leg.get("formas") or {}
    leg_nuvem = leg.get("nuvem") or {}
    leg_sankey = leg.get("sankey") or {}
    rodape = str(leg.get("rodape") or "")
    # Gráficos usam *todas* as nucleares (N_hits), não o conjunto deduplicado
    # (N_dedup). Os testes χ²/BF usam dedup — daí Resumo ter 1171 e 1156.
    nota_n = (
        f"N_hits={n_nuc}"
        + (f" · N_dedup({modo_dedupe})={n_dedup}" if modo_dedupe != "nenhuma"
           else "")
    )
    g1 = base / "_g_freq_token.png"
    g1x = base / "_g_freq_token.xlsx"
    docs_bar = (list(freq_tok["n_documentos"])
                if len(freq_tok) and "n_documentos" in freq_tok.columns
                else list(freq_tok["obras_com_ocorrencia"]) if len(freq_tok)
                else [])
    tit_f = str(leg_formas.get("titulo")
                or "Frequência por termo canónico")
    if f"N={n_nuc}" not in tit_f and "N_hits=" not in tit_f:
        tit_f = f"{tit_f}  ({nota_n})"
    sub_f = str(leg_formas.get("subtitulo") or "")
    base_sub = (
        "unidade=canonical_term · conjunto=nuclear "
        f"(após revisto_por_humano) · sem desduplicação de contexto · N={n_nuc}"
    )
    sub_f = f"{sub_f} · {base_sub}".strip(" ·") if sub_f else base_sub
    xlab_f = str(leg_formas.get("xlabel")
                 or "Ocorrencias (N_hits) e documentos")
    labs_f = list(freq_tok["canonical_term"]) if len(freq_tok) else []
    hits_f = list(freq_tok["ocorrencias_token"]) if len(freq_tok) else []
    cores_f = [tplot.cor_canonical(t) for t in labs_f]
    kw_barras = dict(
        titulo=tit_f, subtitulo=sub_f, nome_a="N_hits",
        nome_b="n_documentos", xlabel=xlab_f, rodape=rodape,
        max_n=None, cores=cores_f,
    )
    try:
        tplot.barras_horizontais_agrupadas(
            labs_f, hits_f, docs_bar, g1, **kw_barras)
        g_msgs.append(f"OK freq token -> {g1.name}")
    except Exception as exc:
        g_msgs.append(f"FALHA freq: {exc}")
    try:
        tplot.barras_horizontais_agrupadas_xlsx(
            labs_f, hits_f, docs_bar, g1x, **kw_barras)
        g_msgs.append(f"OK freq token Excel -> {g1x.name}")
    except Exception as exc:
        g_msgs.append(f"FALHA freq Excel: {exc}")
    tab_barras = pd.DataFrame({
        "canonical_term": list(freq_tok["canonical_term"]) if len(freq_tok) else [],
        "N_hits": list(freq_tok["ocorrencias_token"]) if len(freq_tok) else [],
        "n_documentos": docs_bar if len(freq_tok) else [],
    })
    if len(tab_barras):
        tab_barras = tab_barras.sort_values(
            ["N_hits", "canonical_term"], ascending=[False, True]
        ).reset_index(drop=True)

    g_eixos = base / "_g_freq_eixos.png"
    g_eixos_leg = base / "_g_freq_eixos_legenda.txt"
    g_fluxo = base / "curadoria_fluxo.tsv"
    tab_fluxo = pd.DataFrame()
    try:
        from textura.lexico import carregar_eixos_curadoria
        cura_map = carregar_eixos_curadoria()
        df_eixos = tab_barras.copy()
        if len(df_eixos):
            aplicada, _em_falta = tlex.aplicar_curadoria(
                list(df_eixos["canonical_term"]), cura_map)
            df_eixos["eixo"] = [
                aplicada[str(t)]["eixo"] for t in df_eixos["canonical_term"]]
            df_eixos["decisao"] = [
                aplicada[str(t)]["decisao"] for t in df_eixos["canonical_term"]]
            df_eixos["motivo"] = [
                aplicada[str(t)]["motivo"] for t in df_eixos["canonical_term"]]
        else:
            for _c in ("eixo", "decisao", "motivo"):
                df_eixos[_c] = pd.Series(dtype=str)
        campo_rot = None
        if lexico:
            try:
                campo_rot = tlex.carregar_campo_termos(Path(lexico))
            except Exception:
                campo_rot = None
        rotulos_eixos = tlex.rotulos_exibicao_lexico(
            list(df_eixos["canonical_term"]) if len(df_eixos) else [],
            campo=campo_rot,
        )
        leg_eixos = leg.get("formas_eixos") or {}
        cauda_eixos, _fig_eixos = tplot.barras_agrupadas_por_eixo(
            df_eixos, g_eixos,
            modo_tese=modo_tese,
            legendas=leg_eixos,
            rotulos_exibicao=rotulos_eixos,
            rodape=rodape,
            titulo=str(leg_eixos.get("titulo") or ""),
            subtitulo=str(leg_eixos.get("subtitulo") or ""),
        )
        g_msgs.append(f"OK freq eixos -> {g_eixos.name}")
        escrever_legenda_eixos(
            g_eixos_leg, n=n_nuc, cauda=cauda_eixos,
            conjunto="(após revisto_por_humano) ·",
        )
        g_msgs.append(f"OK legenda eixos -> {g_eixos_leg.name}")
        tab_fluxo, aviso_brutas = tabela_curadoria_fluxo(
            brutas, nuc, col_doc, cura_map)
        if aviso_brutas:
            print(f"  [aviso] {aviso_brutas}", flush=True)
            g_msgs.append(f"AVISO fluxo: {aviso_brutas}")
        tab_fluxo.to_csv(
            g_fluxo, sep="\t", index=False, encoding="utf-8",
            lineterminator="\n",
        )
        g_msgs.append(f"OK curadoria_fluxo -> {g_fluxo.name}")
    except Exception as exc:
        g_msgs.append(f"FALHA freq eixos: {exc}")

    lex_path = Path(lexico) if lexico else tfreq.lexico_ao_lado(xlsx)
    mapa_fam = tfreq.ler_mapa_familia_lexico(lex_path) if lex_path else {}
    if mapa_fam:
        print(f"Lexico familias: {lex_path.name} ({len(mapa_fam)} etiquetas)",
              flush=True)
    tab_familias = tfreq.tabela_frequencia_familia(nuc, lexico=mapa_fam)
    g_fam = base / "_g_freq_familia.png"
    g_famx = base / "_g_freq_familia.xlsx"
    tit_fam = f"Frequência por família (substantivo)  ({nota_n})"
    sub_fam = (
        "unidade=familia_substantivo · conjunto=nuclear "
        f"(após revisto_por_humano) · sem desduplicação de contexto · N={n_nuc}"
    )
    if len(tab_familias):
        try:
            tfreq.escrever_barras(
                list(tab_familias["familia"]),
                list(tab_familias["N_hits"]),
                list(tab_familias["n_documentos"]),
                g_fam, g_famx,
                titulo=tit_fam, subtitulo=sub_fam, rodape=rodape,
                chaves_cor=list(tab_familias["familia"]),
                nome_cat="familia",
            )
            g_msgs.append(f"OK freq familia -> {g_fam.name}")
        except Exception as exc:
            g_msgs.append(f"FALHA freq familia: {exc}")

    familia_nuvem = mapa_forma_familia(formas)
    cores_nuvem = tplot.cores_por_forma(familia_nuvem)
    freq_nuvem_hits = freq_nuvem_de_formas(formas, ponderacao="hits")
    freq_nuvem_docs = freq_nuvem_de_formas(formas, ponderacao="docs")
    tab_nuvem = pd.concat([
        _df_pesos_grafico(freq_nuvem_hits, ponderacao="hits",
                          familia=familia_nuvem),
        _df_pesos_grafico(freq_nuvem_docs, ponderacao="docs",
                          familia=familia_nuvem),
    ], ignore_index=True)
    g_nuvem_hits = base / "_g_nuvem_hits.png"
    g_nuvem_docs = base / "_g_nuvem_docs.png"
    g_nuvem = g_nuvem_hits  # alias: a nuvem por hits é a figura de referência
    tit_n = str(leg_nuvem.get("titulo") or "Nuvem de palavras")
    for dest, freq, pond, fonte in (
        (g_nuvem_hits, freq_nuvem_hits, "N_hits", "12_Formas.n"),
        (g_nuvem_docs, freq_nuvem_docs, "n_documentos", "12_Formas.n_documentos"),
    ):
        try:
            sub_n = (
                f"ponderação={pond} ({fonte}) · unidade=matched_form · "
                "cor=canonical_term · conjunto=nuclear "
                f"(após revisto_por_humano) · sem desduplicação de contexto · N={n_nuc}"
            )
            tplot.nuvem_palavras(
                freq, dest,
                titulo=tit_n,
                subtitulo=sub_n,
                rodape=rodape,
                max_words=None,
                random_state=NUVEM_RANDOM_STATE,
                cores_por_palavra=cores_nuvem)
            g_msgs.append(f"OK nuvem {pond} -> {dest.name} ({len(freq)} formas)")
        except Exception as exc:
            g_msgs.append(f"FALHA nuvem {pond}: {exc}")
    if g_nuvem_hits.exists():
        shutil.copy2(g_nuvem_hits, base / "_g_nuvem.png")

    g_sankey1 = base / "_g_sankey_forma_obra.html"
    g_sankey2 = base / "_g_sankey_termo_rel.html"
    try:
        cap_s1 = None
        p1 = tplot.pares_forma_obra(nuc, max_pares=cap_s1)
        n_lig1 = int(sum(w for *_, w in p1)) if p1 else 0
        n_form_tot = int(nuc["matched_form"].nunique()) if len(nuc) else 0
        n_doc_tot = int(nuc[col_doc].nunique()) if len(nuc) else 0
        n_form_show = len({a for a, _, _ in p1})
        n_doc_show = len({b for _, b, _ in p1})
        tit_s1 = str(leg_sankey.get("titulo")
                     or "Sankey: forma → documento")
        if "N=" not in tit_s1 and "N_hits=" not in tit_s1:
            tit_s1 = f"{tit_s1}  ({nota_n})"
        sub_s1 = ""
        if cap_s1 is not None and n_lig1 < n_nuc:
            sub_s1 = (
                f"showing {n_lig1} of {n_nuc} occurrences "
                f"({n_form_show} of {n_form_tot} forms, "
                f"{n_doc_show} of {n_doc_tot} documents); "
                f"link cap = {cap_s1}"
            )
            avisos.append(f"_g_sankey_forma_obra.html: {sub_s1}")
        tplot.sankey_html(p1, g_sankey1, titulo=tit_s1, subtitulo=sub_s1)

        nuc_graf = nuc
        for col, rotulo, _fam in CAMPOS_LEXICO_OPCIONAIS:
            nuc_graf = _preencher_lexico_vazio(nuc_graf, col, rotulo)
        cap_s2 = None
        p2 = tplot.pares_termo_rel_pol(nuc_graf, max_pares=cap_s2)
        n_lig2_a = int(sum(
            w for a, b, w in p2
            if a in set(nuc_graf["canonical_term"].astype(str))
        )) if p2 else 0
        tit_s2 = f"termo → relação → polaridade  ({nota_n})"
        sub_s2 = ""
        if cap_s2 is not None:
            n_pares2 = len(p2)
            sub_s2 = f"showing {n_pares2} links; link cap = {cap_s2}"
            avisos.append(f"_g_sankey_termo_rel.html: {sub_s2}")
        tplot.sankey_html(p2, g_sankey2, titulo=tit_s2, subtitulo=sub_s2)
        g_msgs.append("OK sankey HTML")
    except Exception as exc:
        g_msgs.append(f"FALHA sankey: {exc}")

    for m in g_msgs:
        if str(m).startswith("FALHA"):
            avisos.append(str(m))

    # --- Resumo ----------------------------------------------------------
    hits_por_doc = (
        nuc[col_doc].astype(str).value_counts()
        if len(nuc) and col_doc in nuc.columns else pd.Series(dtype=int)
    )
    n_pol_adj = int(_serie_nao_vazia(nuc["polaridade"]).sum()) if (
        len(nuc) and "polaridade" in nuc.columns) else 0
    n_pol_sem = n_nuc - n_pol_adj
    n_eixo_adj = int(_serie_nao_vazia(nuc["eixo"]).sum()) if (
        len(nuc) and "eixo" in nuc.columns) else 0
    n_eixo_sem = n_nuc - n_eixo_adj
    mediana_hpd = (
        float(hits_por_doc.median()) if len(hits_por_doc) else float("nan"))
    n_doc_1 = int((hits_por_doc == 1).sum()) if len(hits_por_doc) else 0
    gini_hpd = gini_index(hits_por_doc.values) if len(hits_por_doc) else float("nan")

    resumo = pd.DataFrame({
        "indicador": [
            "Fase", "Data fase 2", "Comando fase 1 (meta)",
            "Schema near (meta)",
            "Linhas brutas (hits)", "Linhas nucleares (N_hits)",
            f"Linhas apos desduplicacao ({modo_dedupe})",
            "Ocorrencias nucleares (N_ocorrencias)",
            "Janelas KWIC do no (denominador logDice)",
            "Ficheiros", "Obras (doc_id)",
            "Tipos (S)", "H (ocorrencias_token)", "J", "1/D",
            "Revisao: linhas marcadas", "Revisao: taxa marcacao",
            "Avisos", "Nulo polaridade", "Unidade co-ocorrencia",
            "Modo desduplicacao", "Analise sem revisao humana",
            "n_hits_com_polaridade_adjudicada",
            "n_hits_sem_polaridade_adjudicada",
            "n_hits_com_eixo_adjudicado",
            "n_hits_sem_eixo_adjudicado",
            "mediana_hits_por_documento",
            "n_documentos_com_1_hit",
            "gini_hits_por_documento",
            "plano_a_priori", "kappa_nuclear", "kappa_relacao",
        ],
        "valor": [
            "2 - analise",
            datetime.now().isoformat(timespec="seconds"),
            meta.get("comando", "—"),
            meta.get("schema_near", "—"),
            n_bruto, n_nuc, n_dedup, n_occ_nuc, n_janelas_no,
            brutas["caminho_ficheiro"].nunique() if "caminho_ficheiro" in brutas else "—",
            brutas[col_doc].nunique() if col_doc in brutas.columns else "—",
            S, round(H, 4) if H == H else "—",
            round(J, 4) if J == J else "—",
            round(invD, 4) if invD == invD else "—",
            rev["marcadas_revisto_por_humano"], rev["taxa_marcacao"],
            " | ".join(avisos) if avisos else "—",
            nulo_polaridade, cooc_unidade, modo_dedupe,
            meta.get("sem_revisao", "nao"),
            n_pol_adj, n_pol_sem, n_eixo_adj, n_eixo_sem,
            mediana_hpd if mediana_hpd == mediana_hpd else "—",
            n_doc_1,
            round(gini_hpd, 4) if gini_hpd == gini_hpd else "—",
            str(plano_a_priori) if plano_a_priori else "—",
            (tab_kappa.loc[tab_kappa["campo"] == "nuclear", "kappa"].iloc[0]
             if len(tab_kappa) and "campo" in tab_kappa.columns else "—"),
            (tab_kappa.loc[tab_kappa["campo"] == "relacao_sintactica", "kappa"].iloc[0]
             if len(tab_kappa) and "campo" in tab_kappa.columns else "—"),
        ],
    })

    # --- Export ----------------------------------------------------------
    print(f"A escrever {saida.name} ...", flush=True)
    saida = Path(saida)
    xlsx = Path(xlsx)
    if saida.resolve() != xlsx.resolve():
        saida.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(xlsx, saida)
    folhas_escritas: set[str] = set()
    with pd.ExcelWriter(
        saida, engine="openpyxl", mode="a", if_sheet_exists="replace",
    ) as xw:
        resumo.to_excel(xw, sheet_name="1_Resumo", index=False)
        folhas_escritas.add("1_Resumo")
        _escrever_folha(xw, "2_Frequencias", freq_tok,
                        "linhas nucleares (ocorrencias_token) / doc_id",
                        n_nuc)
        folhas_escritas.add("2_Frequencias")
        if len(tst_df):
            out_t = tst_df.copy()
            for c in ("p", "p_ajustado_BH"):
                if c in out_t.columns:
                    out_t[c] = out_t[c].map(
                        lambda v: f"{v:.5g}" if v == v else "")
            _escrever_folha(xw, "3_Testes", out_t,
                            f"nucleares apos desduplicacao ({modo_dedupe})",
                            n_dedup)
            folhas_escritas.add("3_Testes")
        if len(sintaxe):
            _escrever_folha(xw, "4_Sintaxe", sintaxe.reset_index(),
                            f"nucleares apos desduplicacao ({modo_dedupe})",
                            n_dedup)
            folhas_escritas.add("4_Sintaxe")
        if len(cooc):
            _escrever_folha(xw, "5_Coocorrencia", cooc.reset_index(),
                            f"obras (co-presenca; unidade={cooc_unidade})",
                            int(nuc[col_doc].nunique()))
            folhas_escritas.add("5_Coocorrencia")
        # Concordância sem banner meta — reutilizável / legível por pandas
        _escrever_folha(
            xw, "8_Concordancia", brutas,
            "todas as linhas (brutas=hits)", n_bruto, com_meta=False,
        )
        folhas_escritas.add("8_Concordancia")
        _escrever_folha(
            xw, "8_Concordancia_Hits", brutas,
            "hits NEAR (= 8_Concordancia)", n_bruto, com_meta=False,
        )
        folhas_escritas.add("8_Concordancia_Hits")
        if len(assoc):
            cols_a = [c for c in (
                "canonical_term", "O11_janela", "O12_banda_ref", "obras",
                "razao_possib", "IC95_inf", "IC95_sup", "DeltaP",
                "logDice", "log_likelihood_G2", "MI", "MI3_heuristica",
                "p_fisher", "leitura", "r1_tokens_near", "r2_tokens_banda",
                "n_janelas_no", "formula_logDice", "DP_Gries", "D_Juilland",
                "partes_nao_nulas",
            ) if c in assoc.columns]
            _escrever_folha(xw, "9_Associacao", assoc[cols_a],
                            "linhas nucleares (O11); logDice usa janelas do no",
                            n_nuc)
            folhas_escritas.add("9_Associacao")
        _escrever_folha(xw, "12_Formas", formas,
                        "linhas nucleares; inventario fechado pela consulta",
                        n_nuc)
        folhas_escritas.add("12_Formas")
        if len(reg_tab):
            _escrever_folha(xw, "13_Regressao", reg_tab,
                            "linhas brutas; DV=nuclear vs incidental",
                            n_bruto)
            folhas_escritas.add("13_Regressao")
            if reg_aj:
                pd.DataFrame({"indicador": list(reg_aj),
                              "valor": list(reg_aj.values())}).to_excel(
                    xw, sheet_name="13_Ajuste_modelo", index=False)
                folhas_escritas.add("13_Ajuste_modelo")
        if len(perfil):
            _escrever_folha(xw, "15_Perfis", perfil,
                            f"nucleares apos desduplicacao ({modo_dedupe})",
                            n_dedup)
            folhas_escritas.add("15_Perfis")
        pd.DataFrame({"aviso": avisos or ["—"]}).to_excel(
            xw, sheet_name="0_Avisos", index=False)
        folhas_escritas.add("0_Avisos")
        tab_nuvem.to_excel(xw, sheet_name="6_Graficos_nuvem", index=False)
        folhas_escritas.add("6_Graficos_nuvem")
        tab_barras.to_excel(xw, sheet_name="6_Graficos_barras", index=False)
        folhas_escritas.add("6_Graficos_barras")
        tab_familias.to_excel(xw, sheet_name="6_Graficos_familias", index=False)
        folhas_escritas.add("6_Graficos_familias")
        if len(tab_kappa):
            tab_kappa.to_excel(xw, sheet_name="16_Kappa", index=False)
            folhas_escritas.add("16_Kappa")
        if len(tab_fluxo):
            _escrever_folha(
                xw, "curadoria_fluxo", tab_fluxo,
                "fluxo de curadoria (brutas vs nucleares)",
                n_nuc, com_meta=False,
            )
            folhas_escritas.add("curadoria_fluxo")

    # embutir graficos; preservar folhas do investigador
    wb = load_workbook(saida)
    for nome in list(wb.sheetnames):
        if nome in FOLHAS_GERADAS_FASE2 and nome not in folhas_escritas | {"6_Graficos"}:
            del wb[nome]
    if "6_Graficos" in wb.sheetnames:
        del wb["6_Graficos"]
    if "6_Graficos_barras" in wb.sheetnames and len(tab_barras):
        try:
            tplot.adicionar_grafico_barras_agrupadas(
                wb["6_Graficos_barras"],
                n=len(tab_barras),
                titulo=tit_f,
                xlabel=xlab_f,
                cores=[tplot.cor_canonical(t)
                       for t in tab_barras["canonical_term"]],
                ancora="E2",
            )
        except Exception as exc:
            g_msgs.append(f"FALHA grafico nativo 6_Graficos_barras: {exc}")
    if "6_Graficos_familias" in wb.sheetnames and len(tab_familias):
        try:
            tplot.adicionar_grafico_barras_agrupadas(
                wb["6_Graficos_familias"],
                n=len(tab_familias),
                titulo=tit_fam,
                xlabel=xlab_f,
                cores=[tplot.cor_canonical(t)
                       for t in tab_familias["familia"]],
                ancora="E2",
            )
        except Exception as exc:
            g_msgs.append(f"FALHA grafico nativo 6_Graficos_familias: {exc}")
    ws = wb.create_sheet("6_Graficos")
    ws["A1"] = f"Graficos | N_nuclear={n_nuc} | {datetime.now().date()}"
    ws["A2"] = " | ".join(g_msgs)
    row = 4
    for p, label in (
        (g1, "Frequencia token"),
        (g_eixos, "Frequencia por eixo (curadoria)"),
        (g_fam, "Frequencia familia (substantivo)"),
        (g_nuvem_hits, "Nuvem (N_hits)"),
        (g_nuvem_docs, "Nuvem (n_documentos)"),
    ):
        ws.cell(row=row, column=1, value=label)
        if p.exists():
            try:
                ws.add_image(XLImage(str(p)), f"A{row + 1}")
            except Exception as exc:
                ws.cell(row=row + 1, column=1, value=f"FALHA embutir: {exc}")
        else:
            ws.cell(row=row + 1, column=1, value="figura nao gerada")
        row += 35
    ws.cell(row=row, column=1,
            value="Fonte dos graficos: folhas 6_Graficos_barras e "
                  "6_Graficos_familias (edite ahi). "
                  f"Copias: {g1x.name} / {g_famx.name}")
    row += 2
    ws.cell(row=row, column=1,
            value=f"Sankey HTML: {g_sankey1.name} ; {g_sankey2.name}")
    _anotar_instrucoes_adjudicacao(wb)
    for nome in wb.sheetnames:
        if nome not in FOLHAS_GERADAS_FASE2:
            continue
        s = wb[nome]
        for c in s[1]:
            if c.value is not None:
                c.font = Font(name="Arial", bold=True)
        s.freeze_panes = "A2"
    wb.save(saida)
    wb.close()

    print(resumo.to_string(index=False))
    print()
    print("=== Análise concluída com sucesso ===")
    print(f"Ficheiro de saída (fase Analisar): {saida}")
    print(
        "Próximo passo na GUI: «4. Apêndice DOCX» — pode usar este "
        "*_analise.xlsx ou o Excel revisto (*_v2.xlsx)."
    )
    if avisos:
        print()
        print(
            "Notas estatísticas (salvaguardas — a análise NÃO falhou):"
        )
        for a in avisos:
            txt = a.strip()
            if txt.lower().startswith("aviso:"):
                txt = txt.split(":", 1)[1].strip()
            print(f"  · {txt}")
    return 0


def main() -> int:
    _cfg_consola()
    ap = argparse.ArgumentParser(
        description="Fase 2: estatistica e graficos sobre Excel revisto")
    ap.add_argument("--xlsx", type=Path, required=True)
    ap.add_argument("--saida", type=Path, default=None)
    ap.add_argument("--nulo-polaridade", choices=["banda", "lexico"],
                    default="banda")
    ap.add_argument("--cooc-unidade", choices=["obra", "frase"],
                    default="obra")
    ap.add_argument("--desduplicacao", choices=list(MODOS_DEDUPE),
                    default="nenhuma",
                    help="nenhuma=usar nucleares tal como revistas; "
                         "contexto=1 linha por snippet; "
                         "candidatos=respeita candidato_duplicado; "
                         "obra_termo=sensibilidade doc x termo; "
                         "ocorrencia=1 linha por texture_occurrence_id; "
                         "ocorrencia_termo=1 linha por ocorrencia x termo")
    ap.add_argument("--legendas", type=Path, default=None,
                    help="JSON com titulos/subtitulos editaveis dos graficos")
    ap.add_argument("--relacao", default="",
                    help="subset de relacoes (virgulas); vazio=todas nucleares")
    ap.add_argument("--plano-a-priori", type=Path, default=None,
                    help="JSON ou YAML minimo com desduplicacao / nulo_polaridade "
                         "/ relacao / cooc_unidade (o plano prevalece)")
    ap.add_argument("--kappa-cego", type=Path, default=None,
                    help="Excel (8_Concordancia) ou JSON do segundo revisor; "
                         "calcula kappa de Cohen por hit_key")
    ap.add_argument("--lexico", type=Path, default=None,
                    help="termos adjudicados (etiqueta / familia = padroes)")
    ap.add_argument(
        "--estrito",
        action="store_true",
        help="falhar se o checklist de revisao tiver erros "
             "(nuclear != relacao, 0 nucleares, ...)",
    )
    ap.add_argument(
        "--modo-tese",
        action="store_true",
        help="grafico de eixos sem titulo/rodape na imagem; PNG 300 dpi + SVG",
    )
    args = ap.parse_args()
    rels = [p.strip() for p in (args.relacao or "").split(",") if p.strip()]
    return analisar(
        args.xlsx, args.saida,
        nulo_polaridade=args.nulo_polaridade,
        cooc_unidade=args.cooc_unidade,
        desduplicacao=args.desduplicacao,
        legendas=args.legendas,
        relacoes=rels or None,
        estrito=args.estrito,
        plano_a_priori=args.plano_a_priori,
        kappa_cego=args.kappa_cego,
        lexico=args.lexico,
        modo_tese=args.modo_tese,
    )


if __name__ == "__main__":
    raise SystemExit(main())
