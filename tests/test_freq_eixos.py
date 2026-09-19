# -*- coding: utf-8 -*-
"""Gráfico de frequência por eixo de curadoria (tese)."""

from __future__ import annotations

import warnings
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

import textura_analise as ta
import textura_lexico as tlex
import textura_plots as tplot
from textura.lexico import carregar_eixos_curadoria


def _wb_minimo(path: Path, *, nuclear=None, extra_sheet=True,
               revisto=None, polaridade=None):
    n = 4
    nuclear = [True, True, True, False] if nuclear is None else nuclear
    conc = pd.DataFrame({
        "hit_key": [f"h{i}" for i in range(n)],
        "canonical_term": ["static", "static", "uniform", "varied"],
        "matched_form": ["static", "static", "uniform", "varied"],
        "nuclear": nuclear,
        "relacao_sintactica": [
            "atributiva", "atributiva", "predicativa", "incidental"],
        "polaridade": polaridade or [
            "estabilidade", "estabilidade", "estabilidade", ""],
        "eixo": ["ambos", "ambos", "ambos", "ambos"],
        "doc_id": ["DocA", "DocA", "DocB", "DocC"],
        "contexto": ["ctx a", "ctx a", "ctx b", "ctx c"],
        "revisto_por_humano": revisto or ["", "", "", ""],
        "caminho_ficheiro": ["a.pdf", "a.pdf", "b.pdf", "c.pdf"],
    })
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        pd.DataFrame({
            "chave": ["consulta", "janelas_kwic_processadas"],
            "valor": ["textur* NEAR/4", 100],
        }).to_excel(xw, sheet_name="0_Instrucoes", index=False)
        conc.to_excel(xw, sheet_name="8_Concordancia", index=False)
        conc.to_excel(xw, sheet_name="8_Concordancia_Hits", index=False)
        if extra_sheet:
            pd.DataFrame({"nota": ["legend"], "n": [1]}).to_excel(
                xw, sheet_name="Resumo_ligacao", index=False)

ROOT = Path(__file__).resolve().parents[1]


def _df_eixos():
    return pd.DataFrame({
        "canonical_term": [
            "uniform", "static", "steady", "homogeneous", "varied", "ghost",
        ],
        "N_hits": [12, 8, 1, 5, 9, 3],
        "n_documentos": [6, 4, 1, 3, 5, 2],
        "eixo": [
            "invariancia", "invariancia", "invariancia",
            "semelhanca_interna", "variabilidade", "por_classificar",
        ],
        "decisao": [
            "retido", "retido", "relacionado",
            "excluido", "excluido", "por_classificar",
        ],
    })


def test_loader_termo_em_falta_nao_rebenta():
    mapa = carregar_eixos_curadoria()
    assert mapa["uniform"]["eixo"] == "invariancia"
    assert mapa["homogeneous"]["decisao"] == "excluido"
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        out, falta = tlex.aplicar_curadoria(["uniform", "xyz_ausente"], mapa)
    assert "xyz_ausente" in falta
    assert out["xyz_ausente"]["eixo"] == "por_classificar"
    assert out["xyz_ausente"]["decisao"] == "por_classificar"
    assert out["uniform"]["decisao"] == "retido"
    assert any("por_classificar" in str(w.message) for w in caught)


def test_agrega_ordem_intra_grupo():
    tab, cauda = tplot._agregar_por_eixo(_df_eixos(), limiar_cauda=0)
    assert cauda == {}
    eixos = list(tab["eixo"])
    assert eixos.index("invariancia") < eixos.index("semelhanca_interna")
    assert eixos.index("semelhanca_interna") < eixos.index("variabilidade")
    inv = tab.loc[tab["eixo"] == "invariancia", "canonical_term"].tolist()
    assert inv == ["uniform", "static", "steady"]


def test_colapsa_cauda_e_soma():
    tab, cauda = tplot._agregar_por_eixo(_df_eixos(), limiar_cauda=2)
    assert cauda["invariancia"] == ["steady"]
    outros = tab.loc[tab["canonical_term"].astype(str).str.startswith("outros")]
    inv_out = outros.loc[outros["eixo"] == "invariancia"].iloc[0]
    assert inv_out["canonical_term"] == "outros (1)"
    assert float(inv_out["N_hits"]) == 1
    assert float(inv_out["n_documentos"]) == 1
    assert "steady" not in list(tab.loc[tab["eixo"] == "invariancia", "canonical_term"])


def test_modo_tese_png_svg_sem_titulo(tmp_path):
    dest = tmp_path / "_g_freq_eixos.png"
    cauda, fig = tplot.barras_agrupadas_por_eixo(
        _df_eixos(), dest, modo_tese=True, fechar=False, limiar_cauda=2,
    )
    assert dest.is_file()
    assert dest.with_suffix(".svg").is_file()
    assert fig is not None
    assert fig._suptitle is None
    titulos = [t.get_text() for t in fig.texts]
    assert titulos == []
    assert cauda["invariancia"] == ["steady"]
    plt.close(fig)


def test_analise_ainda_produz_freq_token(tmp_path):
    src = tmp_path / "in.xlsx"
    dst = tmp_path / "out.xlsx"
    _wb_minimo(src)
    ta.analisar(src, dst, desduplicacao="nenhuma")
    assert (tmp_path / "_g_freq_token.png").is_file()
    assert (tmp_path / "_g_freq_eixos.png").is_file()
    assert (tmp_path / "_g_freq_eixos.svg").is_file()
    assert (tmp_path / "_g_freq_eixos_legenda.txt").is_file()
    texto = (tmp_path / "_g_freq_eixos_legenda.txt").read_text(encoding="utf-8")
    assert ta.FRASE_LEGENDA_EIXOS in texto
    assert (tmp_path / "curadoria_fluxo.tsv").is_file()
    with pd.ExcelFile(dst) as xl:
        assert "curadoria_fluxo" in xl.sheet_names


def test_fluxo_brutas_vazias_se_coluna_em_falta():
    nuc = pd.DataFrame({
        "canonical_term": ["static", "static"],
        "doc_id": ["A", "B"],
    })
    brutas = pd.DataFrame({"outra": [1, 2]})
    df, aviso = ta.tabela_curadoria_fluxo(brutas, nuc, "doc_id", {
        "static": {"eixo": "invariancia", "decisao": "retido", "motivo": ""},
    })
    assert aviso
    assert df.loc[0, "coocorrencias_brutas"] == ""
    assert int(df.loc[0, "atribuicoes_genuinas"]) == 2
