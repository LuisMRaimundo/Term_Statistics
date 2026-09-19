#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Relatório filtrável de frequência por ``canonical_term`` / família substantivo.

Reutiliza ``textura_plots.barras_horizontais_agrupadas[_xlsx]`` — o formato
do gráfico é o de ``_g_freq_token``. Não reimplementa a figura.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

import textura_plots as tplot

COL_DOC_PREFERIDA = ("doc_id", "caminho_ficheiro")
_NUCLEAR_TRUE = {"true", "1", "sim", "nuclear"}
_NUCLEAR_FALSE = {"false", "0", "nao", "não"}
_SLUG_NAO_PALAVRA = re.compile(r"[^\w\-]+", flags=re.UNICODE)
# Sufixos nominais (mais longos primeiro). Adj/adv ficam de fora em
# ``e_substantivo`` — a família é o substantivo do grupo (homogeneity).
_SUFIXOS_SUBSTANTIVO = (
    "ization", "isation", "ilidade", "idade", "ility",
    "ização", "isação", "ness",
    "tion", "sion", "ção",
    "mento", "ment",
    "ance", "ence", "ância", "ência", "ancy", "ency",
    "ismo", "isme", "ism",
    "schaft", "heit", "keit", "ship", "hood", "ung",
    "ity", "ité", "tät",
    "ude", "ura", "ure",
    "eza", "cy",
)


def coluna_documento(df: pd.DataFrame) -> str:
    """Espelha textura_analise.py:751."""
    return "doc_id" if "doc_id" in df.columns else "caminho_ficheiro"


def normalizar_nuclear(serie: pd.Series) -> pd.Series:
    """Espelha textura_analise.py:735-736."""
    return serie.map(
        lambda v: v is True or str(v).strip().lower() in {"true", "1", "sim"}
    )


def filtrar(df: pd.DataFrame, *,
            nuclear: str = "true",
            termos: list[str] | None = None,
            formas: list[str] | None = None,
            ) -> pd.DataFrame:
    d = df.copy()
    d["nuclear"] = normalizar_nuclear(d["nuclear"])
    modo = str(nuclear).strip().lower()
    if modo in _NUCLEAR_TRUE:
        d = d.loc[d["nuclear"]]
    elif modo in _NUCLEAR_FALSE:
        d = d.loc[~d["nuclear"]]
    if termos:
        alvo = {t.strip().lower() for t in termos if t and t.strip()}
        d = d.loc[d["canonical_term"].astype(str).str.strip().str.lower().isin(alvo)]
    if formas:
        alvo = {f.strip().lower() for f in formas if f and f.strip()}
        d = d.loc[d["matched_form"].astype(str).str.strip().str.lower().isin(alvo)]
    return d


def tabela_frequencia(df: pd.DataFrame, chave: str = "canonical_term") -> pd.DataFrame:
    """N_hits + n_documentos por ``chave``. Espelha analise:834-838 + 1136-1138."""
    if not len(df) or chave not in df.columns:
        return pd.DataFrame(columns=[chave, "N_hits", "n_documentos"])
    col_doc = coluna_documento(df)
    freq = (df.groupby(chave)
              .agg(N_hits=(chave, "size"),
                   n_documentos=(col_doc, "nunique"))
              .reset_index()
              .sort_values(["N_hits", chave], ascending=[False, True])
              .reset_index(drop=True))
    return freq


def e_substantivo(forma: str) -> bool:
    """True se a forma parece substantivo (sufixo), não adj/adv/verbo."""
    f = str(forma or "").strip().lower()
    if len(f) < 3:
        return False
    if f.endswith(("ly", "mente", "ously", "ively")):
        return False
    if f.endswith(("eous", "ious", "ous", "ical", "ular")):
        return False
    if f.endswith(("izing", "ising", "ized", "ised")):
        return False
    if f.endswith(("ing", "ed")):
        return False
    return any(f.endswith(s) for s in _SUFIXOS_SUBSTANTIVO)


def familia_do_grupo(formas) -> str:
    """Nome da família = substantivo mais frequente; senão a forma mais comum."""
    cont = pd.Series(list(formas)).astype(str).str.strip()
    cont = cont[cont.ne("") & cont.str.lower().ne("nan")]
    if not len(cont):
        return ""
    freq = cont.value_counts()
    pares = list(freq.items())
    subst = [(n, c) for n, c in pares if e_substantivo(n)]
    pool = subst or pares
    pool.sort(key=lambda x: (-int(x[1]), str(x[0]).lower()))
    return str(pool[0][0])


def ler_mapa_familia_lexico(caminho: Path | None) -> dict[str, str]:
    """Lê ``etiqueta [: polo] [/ familia] = padroes`` → {etiqueta: familia}."""
    if caminho is None:
        return {}
    path = Path(caminho)
    if not path.is_file():
        return {}
    out: dict[str, str] = {}
    for linha in path.read_text(encoding="utf-8").splitlines():
        linha = linha.split("#", 1)[0].strip()
        if not linha or "=" not in linha:
            continue
        esq, _pads = linha.split("=", 1)
        esq = esq.strip()
        familia = ""
        if "/" in esq:
            esq, familia = (x.strip() for x in esq.split("/", 1))
        if ":" in esq:
            esq = esq.split(":", 1)[0].strip()
        if esq and familia:
            out[esq] = familia
            out[esq.lower()] = familia
    return out


def lexico_ao_lado(xlsx: Path) -> Path | None:
    """``resultado_pesquisa_termos_adjudicados.txt`` ao lado do Excel."""
    stem = Path(xlsx).stem
    for suf in ("_near_analise", "_analise", "_near"):
        if stem.endswith(suf):
            stem = stem[: -len(suf)]
            break
    cand = Path(xlsx).with_name(stem + "_termos_adjudicados.txt")
    return cand if cand.is_file() else None


def mapa_familia_substantivo(df: pd.DataFrame,
                             lexico: dict[str, str] | None = None
                             ) -> dict[str, str]:
    """``canonical_term`` → família (léxico primeiro; senão substantivo)."""
    if "canonical_term" not in df.columns:
        return {}
    col_f = "matched_form" if "matched_form" in df.columns else "canonical_term"
    lexico = lexico or {}
    out: dict[str, str] = {}
    for can, g in df.groupby(df["canonical_term"].astype(str), sort=False):
        chave = str(can)
        if chave in lexico:
            out[chave] = lexico[chave]
        elif chave.lower() in lexico:
            out[chave] = lexico[chave.lower()]
        else:
            out[chave] = familia_do_grupo(g[col_f]) or chave
    return out


def atribuir_familia(df: pd.DataFrame,
                     lexico: dict[str, str] | None = None) -> pd.DataFrame:
    d = df.copy()
    mapa = mapa_familia_substantivo(d, lexico=lexico)
    d["familia"] = d["canonical_term"].map(
        lambda t: mapa.get(str(t), str(t)))
    return d


def tabela_frequencia_familia(df: pd.DataFrame,
                              lexico: dict[str, str] | None = None
                              ) -> pd.DataFrame:
    """N_hits + n_documentos por família (rótulo = substantivo)."""
    if not len(df):
        return pd.DataFrame(columns=[
            "familia", "N_hits", "n_documentos", "canonical_term"])
    d = atribuir_familia(df, lexico=lexico)
    col_doc = coluna_documento(d)
    freq = (d.groupby("familia")
              .agg(N_hits=("familia", "size"),
                   n_documentos=(col_doc, "nunique"),
                   canonical_term=("canonical_term",
                                   lambda s: ", ".join(sorted({str(x) for x in s}))))
              .reset_index()
              .sort_values(["N_hits", "familia"], ascending=[False, True])
              .reset_index(drop=True))
    return freq


def _slug(texto: str) -> str:
    s = _SLUG_NAO_PALAVRA.sub("_", str(texto).strip())
    return (s.strip("_")[:80] or "consulta")


def carregar_base(xlsx: Path, *, sincronizar: bool = True) -> pd.DataFrame:
    """Lê a concordância após override humano.

    ``sincronizar=True`` (omissão): alinha Hits com Concordancia (escreve).
    ``sincronizar=False`` (``--so-leitura``): lê só Concordancia, sem gravar.
    """
    import textura_analise as ta

    xlsx = Path(xlsx)
    ta.sincronizar_hits_com_concordancia(xlsx, escrever=sincronizar)
    if sincronizar:
        df = pd.read_excel(xlsx, sheet_name="8_Concordancia_Hits")
    else:
        df = ta.ler_folha_concordancia(xlsx, "8_Concordancia")
    if "nuclear" in df.columns:
        df = df.copy()
        df["nuclear"] = normalizar_nuclear(df["nuclear"])
    df, _logs = ta.aplicar_override_revisto(df)
    return df


def escrever_barras(labs, hits, docs, png: Path, xlsx: Path, *,
                    titulo: str, subtitulo: str, rodape: str,
                    chaves_cor, nome_cat: str = "canonical_term") -> None:
    cores = [tplot.cor_canonical(t) for t in chaves_cor]
    kw = dict(titulo=titulo, subtitulo=subtitulo, nome_a="N_hits",
              nome_b="n_documentos", xlabel="Ocorrências (N_hits) e documentos",
              rodape=rodape, max_n=None, cores=cores)
    if labs:
        tplot.barras_horizontais_agrupadas(labs, hits, docs, png, **kw)
        tplot.barras_horizontais_agrupadas_xlsx(
            labs, hits, docs, xlsx, nome_cat=nome_cat, **kw)


def gerar(df: pd.DataFrame | None, saida_dir: Path, *,
          xlsx: Path | None = None,
          etiqueta_consulta: str = "",
          nuclear: str = "true",
          termos: list[str] | None = None,
          formas: list[str] | None = None,
          unidade: str = "ambas",
          lexico: dict[str, str] | Path | None = None,
          rodape: str = "TEXTURA · pesquisa bibliográfica de termos") -> dict:
    """Filtra, agrega e escreve gráficos REUTILIZANDO tplot.

    ``unidade``: ``canonical`` (curinga), ``familia`` (substantivo) ou ``ambas``.
    """
    if xlsx is not None:
        df = carregar_base(xlsx)
    if isinstance(lexico, (str, Path)):
        lexico = ler_mapa_familia_lexico(Path(lexico))
    lexico = lexico or {}
    if df is None:
        raise ValueError("indique um DataFrame ou xlsx")
    saida_dir = Path(saida_dir)
    saida_dir.mkdir(parents=True, exist_ok=True)
    d = filtrar(df, nuclear=nuclear, termos=termos, formas=formas)
    freq = tabela_frequencia(d)
    labs = list(freq["canonical_term"]) if len(freq) else []
    hits = list(freq["N_hits"]) if len(freq) else []
    docs = list(freq["n_documentos"]) if len(freq) else []
    n_hits_tot = int(freq["N_hits"].sum()) if len(freq) else 0
    partes = [f"consulta={etiqueta_consulta or '(todas)'}",
              f"conjunto=nuclear={nuclear}"]
    if termos:
        partes.append("termos=" + ",".join(termos))
    if formas:
        partes.append("formas=" + ",".join(formas))
    partes.append(f"N={n_hits_tot}")
    g_png = saida_dir / "_g_freq_token.png"
    g_xlsx = saida_dir / "_g_freq_token.xlsx"
    modo = str(unidade).strip().lower()
    if modo in {"canonical", "canonical_term", "ambas", "ambos", "all", ""}:
        escrever_barras(
            labs, hits, docs, g_png, g_xlsx,
            titulo=f"Frequência por termo canónico  (N_hits={n_hits_tot})",
            subtitulo=" · ".join(["unidade=canonical_term"] + partes),
            rodape=rodape, chaves_cor=labs)

    freq_f = tabela_frequencia_familia(d, lexico=lexico) if "matched_form" in d.columns else (
        pd.DataFrame(columns=["familia", "N_hits", "n_documentos",
                              "canonical_term"]))
    labs_f = list(freq_f["familia"]) if len(freq_f) else []
    hits_f = list(freq_f["N_hits"]) if len(freq_f) else []
    docs_f = list(freq_f["n_documentos"]) if len(freq_f) else []
    g_png_f = saida_dir / "_g_freq_familia.png"
    g_xlsx_f = saida_dir / "_g_freq_familia.xlsx"
    if modo in {"familia", "familia_substantivo", "ambas", "ambos", "all", ""}:
        escrever_barras(
            labs_f, hits_f, docs_f, g_png_f, g_xlsx_f,
            titulo=f"Frequência por família (substantivo)  (N_hits={n_hits_tot})",
            subtitulo=" · ".join(["unidade=familia_substantivo"] + partes),
            rodape=rodape, chaves_cor=labs_f, nome_cat="familia")

    return {"tabela": freq, "png": g_png, "xlsx": g_xlsx,
            "n_termos": len(freq), "n_hits": n_hits_tot,
            "tabela_familia": freq_f, "png_familia": g_png_f,
            "xlsx_familia": g_xlsx_f, "n_familias": len(freq_f)}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="Frequência filtrável por canonical_term / família substantivo")
    ap.add_argument("--xlsx", type=Path, required=True,
                    help="livro adjudicado (.xlsx)")
    ap.add_argument("--folha", default="8_Concordancia_Hits",
                    help="folha de hits (omissão: 8_Concordancia_Hits)")
    ap.add_argument("--nuclear", choices=["true", "false", "ambos"],
                    default="true")
    ap.add_argument("--termo", action="append", default=None, dest="termos",
                    help="canonical_term (repetível)")
    ap.add_argument("--forma", action="append", default=None, dest="formas",
                    help="matched_form (repetível)")
    ap.add_argument("--saida", type=Path, default=Path("./saida_freq"),
                    help="pasta de saída (omissão: ./saida_freq)")
    ap.add_argument("--etiqueta", default="",
                    help="rótulo da consulta (subtítulo e sub-pasta)")
    ap.add_argument("--unidade", choices=["canonical", "familia", "ambas"],
                    default="ambas",
                    help="canonical=curinga; familia=substantivo; ambas (omissão)")
    ap.add_argument("--lexico", type=Path, default=None,
                    help="termos adjudicados (etiqueta / familia = padroes)")
    ap.add_argument("--so-leitura", action="store_true",
                    help="nao escrever Hits; ler so 8_Concordancia")
    args = ap.parse_args(argv)

    xlsx = Path(args.xlsx)
    if not xlsx.is_file():
        print(f"Ficheiro inexistente: {xlsx}", file=sys.stderr)
        return 2
    df = carregar_base(xlsx, sincronizar=not args.so_leitura)
    lexico = ler_mapa_familia_lexico(args.lexico) if args.lexico else (
        ler_mapa_familia_lexico(lexico_ao_lado(xlsx)))
    saida_dir = Path(args.saida)
    if str(args.etiqueta).strip():
        saida_dir = saida_dir / _slug(args.etiqueta)
    out = gerar(
        df, saida_dir,
        etiqueta_consulta=args.etiqueta,
        nuclear=args.nuclear,
        termos=args.termos,
        formas=args.formas,
        unidade=args.unidade,
        lexico=lexico,
    )
    print(f"termos={out['n_termos']}  familias={out['n_familias']}  "
          f"hits={out['n_hits']}")
    print(f"png={out['png']}")
    print(f"xlsx={out['xlsx']}")
    if out.get("xlsx_familia"):
        print(f"familia={out['xlsx_familia']}")
    if len(out["tabela"]):
        print(out["tabela"].to_string(index=False))
    if len(out.get("tabela_familia") or []):
        print(out["tabela_familia"].to_string(index=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
