#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
textura_plots.py — estilo gráfico partilhado (tom académico / editorial)
=======================================================================
"""

from __future__ import annotations

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
from matplotlib.colors import LinearSegmentedColormap

# Paleta editorial: tinta, ardósia, acento oliva — sem roxo/neon
INK = "#1C2429"
SLATE = "#3D4F5F"
MUTED = "#6B7C8A"
RULE = "#D5D2CB"
PAPER = "#F6F4F0"
PANEL = "#FFFFFF"
ACCENT = "#2F5D50"
ACCENT_WARM = "#8B5E3C"
ACCENT_COOL = "#3A536B"
SERIES = ["#2F5D50", "#8B5E3C", "#3A536B", "#6B7C8A", "#A67C52"]

_CMAP_DOCS = LinearSegmentedColormap.from_list(
    "textura_docs", ["#A8C4B8", "#2F5D50"])
_CMAP_FORMAS = LinearSegmentedColormap.from_list(
    "textura_formas", ["#D4B896", "#8B5E3C"])


def aplicar_estilo() -> None:
    plt.rcParams.update({
        "figure.facecolor": PAPER,
        "axes.facecolor": PANEL,
        "axes.edgecolor": RULE,
        "axes.labelcolor": INK,
        "axes.titlecolor": INK,
        "axes.titlesize": 12.5,
        "axes.titleweight": "semibold",
        "axes.titlepad": 10,
        "axes.labelsize": 9.5,
        "axes.labelpad": 6,
        "axes.spines.top": False,
        "axes.spines.right": False,
        "axes.spines.left": True,
        "axes.spines.bottom": True,
        "xtick.color": MUTED,
        "ytick.color": MUTED,
        "xtick.labelsize": 8.5,
        "ytick.labelsize": 8.2,
        "grid.color": RULE,
        "grid.linewidth": 0.6,
        "grid.alpha": 0.85,
        "font.family": "serif",
        "font.serif": ["Cambria", "Georgia", "DejaVu Serif", "Times New Roman"],
        "font.size": 9,
        "legend.frameon": False,
        "legend.fontsize": 8.5,
        "savefig.facecolor": PAPER,
        "savefig.bbox": "tight",
        "savefig.pad_inches": 0.18,
        "figure.dpi": 110,
    })


# Largura das imagens no Excel (pixels) — legível com muitos níveis
EXCEL_LARGURA_PX = 820
EXCEL_PX_POR_LINHA = 14  # ~altura de linha Excel standard


def _cabecalho(fig, ax, titulo: str, subtitulo: str = "") -> None:
    """Título + subtítulo em zonas separadas (sem sobreposição)."""
    ax.set_title("")  # limpar título do eixo
    # espaço superior reservado
    top = 0.86 if subtitulo else 0.90
    fig.subplots_adjust(top=top, bottom=0.11, left=0.14, right=0.96)
    fig.text(0.14, 0.965, titulo, transform=fig.transFigure,
             fontsize=12.5, fontweight="semibold", color=INK,
             ha="left", va="top", fontfamily="serif")
    if subtitulo:
        # uma linha; truncar se for demasiado longa
        sub = subtitulo if len(subtitulo) <= 110 else subtitulo[:107] + "..."
        fig.text(0.14, 0.915, sub, transform=fig.transFigure,
                 fontsize=8.2, color=MUTED, ha="left", va="top",
                 fontfamily="sans-serif")


def _rodape(fig, texto: str) -> None:
    if not texto:
        return
    fig.text(0.14, 0.02, texto, transform=fig.transFigure,
             ha="left", va="bottom", fontsize=7, color=MUTED,
             style="italic", fontfamily="sans-serif")


def _acabar(fig, ax, destino: Path, dpi: int = 150,
            titulo: str = "", subtitulo: str = "", rodape: str = "") -> None:
    ax.tick_params(length=0)
    for spine in ("left", "bottom"):
        ax.spines[spine].set_color(RULE)
        ax.spines[spine].set_linewidth(0.9)
    if titulo:
        _cabecalho(fig, ax, titulo, subtitulo)
    elif subtitulo:
        _cabecalho(fig, ax, subtitulo, "")
    if rodape:
        _rodape(fig, rodape)
    fig.savefig(destino, dpi=dpi, bbox_inches="tight",
                pad_inches=0.35, facecolor=fig.get_facecolor())
    plt.close(fig)


def embeber_imagens(ws, caminhos, *, largura_px: int = EXCEL_LARGURA_PX,
                    col: str = "A", linha0: int = 1, gap_linhas: int = 2):
    """Empilha imagens no Excel sem sobreposição (via BytesIO — fiável no Windows)."""
    from io import BytesIO
    from openpyxl.drawing.image import Image as XLImage
    from PIL import Image as PILImage

    linha = linha0
    embutidas = 0
    for caminho in caminhos:
        p = Path(caminho)
        if not p.exists():
            print(f"      [grafico em falta] {p.name}", flush=True)
            continue
        try:
            with PILImage.open(p) as pil:
                ow, oh = pil.size
                factor = min(1.0, largura_px / float(ow))
                nw, nh = max(1, int(ow * factor)), max(1, int(oh * factor))
                if factor < 1.0:
                    pil = pil.resize((nw, nh), PILImage.Resampling.LANCZOS)
                buf = BytesIO()
                pil.save(buf, format="PNG")
                buf.seek(0)
            img = XLImage(buf)
            img.width = nw
            img.height = nh
            ws.add_image(img, f"{col}{linha}")
            linhas_ocupadas = max(10, int(nh / EXCEL_PX_POR_LINHA) + gap_linhas)
            linha += linhas_ocupadas
            embutidas += 1
            print(f"      [excel] {p.name} -> linha {linha0 if embutidas==1 else '…'} "
                  f"({nw}x{nh})", flush=True)
        except Exception as exc:  # noqa: BLE001
            print(f"      [erro a embutir {p.name}] {exc}", flush=True)
    return linha


def barras_horizontais(rotulos, valores, destino: Path, *,
                       titulo: str, subtitulo: str = "",
                       xlabel: str = "Ocorrências",
                       cmap="docs", rodape: str = "TEXTURA · análise de corpus",
                       max_n: int | None = 60):
    """Barras horizontais. max_n=None mostra todos os níveis."""
    aplicar_estilo()
    if not rotulos:
        return
    pares = sorted(zip(valores, rotulos), key=lambda t: -t[0])
    if max_n is not None:
        pares = pares[:max_n]
    valores = [p[0] for p in pares]
    rotulos = [p[1] for p in pares]
    n = len(rotulos)
    # altura cresce com o nº de níveis (dispersão lexical profunda)
    alt = max(3.5, min(28.0, 0.32 * n + 1.4))
    fig, ax = plt.subplots(figsize=(8.5, alt))

    vals = np.asarray(valores, dtype=float)
    ordem = np.argsort(vals)  # ascendente para barh
    labs = [rotulos[i] for i in ordem]
    v = vals[ordem]
    norm = (v - v.min()) / (v.max() - v.min() + 1e-9)
    cm = _CMAP_DOCS if cmap == "docs" else _CMAP_FORMAS
    cores = cm(0.35 + 0.65 * norm)

    bars = ax.barh(range(n), v, color=cores, height=0.72,
                   edgecolor="none", zorder=3)
    ax.set_yticks(range(n))
    ax.set_yticklabels(labs)
    ax.set_xlabel(xlabel)

    ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True, nbins=6))
    ax.grid(axis="x", zorder=0)
    ax.set_axisbelow(True)
    ax.set_ylim(-0.6, n - 0.4)

    xmax = float(v.max()) if len(v) else 1
    for bar, val in zip(bars, v):
        ax.text(val + xmax * 0.012, bar.get_y() + bar.get_height() / 2,
                f"{int(val):,}".replace(",", " "),
                va="center", ha="left", fontsize=7.5, color=SLATE,
                fontfamily="sans-serif", zorder=4)

    ax.set_xlim(0, xmax * 1.14)
    _acabar(fig, ax, destino, titulo=titulo, subtitulo=subtitulo, rodape=rodape)


def histograma_near(dists, destino: Path, *,
                    titulo: str = "Distribuição das distâncias NEAR",
                    subtitulo: str = "Distância em tokens entre termos casados",
                    rodape: str = "TEXTURA · análise de corpus",
                    xlabel: str = "Distância (tokens)",
                    ylabel: str = "Nº de pares",
                    rotulo_mediana: str = "Mediana",
                    rotulo_media: str = "Média"):
    if not dists:
        return
    aplicar_estilo()
    fig, ax = plt.subplots(figsize=(8.2, 4.0))
    d = np.asarray(dists, dtype=int)
    bins = np.arange(-0.5, d.max() + 1.5, 1)
    counts, edges, patches = ax.hist(
        d, bins=bins, color=ACCENT_COOL, edgecolor=PAPER,
        linewidth=1.1, zorder=3, alpha=0.92)
    # gradiente suave por altura
    if counts.max() > 0:
        for patch, c in zip(patches, counts):
            patch.set_facecolor(_CMAP_DOCS(0.4 + 0.55 * (c / counts.max())))

    med, media = float(np.median(d)), float(np.mean(d))
    ax.axvline(med, color=ACCENT_WARM, ls="--", lw=1.4, zorder=4,
               label=f"{rotulo_mediana} = {med:.1f}")
    ax.axvline(media, color=ACCENT, ls=":", lw=1.4, zorder=4,
               label=f"{rotulo_media} = {media:.2f}")

    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True, nbins=6))
    ax.grid(axis="y", zorder=0)
    ax.set_axisbelow(True)
    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1.0),
              frameon=True, fancybox=False, edgecolor=RULE,
              facecolor=PANEL, framealpha=0.95, borderaxespad=0.0)
    fig.subplots_adjust(right=0.78)
    _acabar(fig, ax, destino, titulo=titulo, subtitulo=subtitulo, rodape=rodape)


def barras_empilhadas(tabela, destino: Path, *, titulo: str,
                      xlabel: str = "", ylabel: str = "Co-occurrences",
                      rodape: str = "TEXTURA corpus analysis"):
    """tabela: DataFrame index=categorias, columns=séries."""
    aplicar_estilo()
    fig, ax = plt.subplots(figsize=(8.2, 4.6))
    tab = tabela.copy()
    x = np.arange(len(tab.index))
    bottom = np.zeros(len(tab))
    for i, col in enumerate(tab.columns):
        vals = tab[col].values.astype(float)
        ax.bar(x, vals, bottom=bottom, color=SERIES[i % len(SERIES)],
               width=0.68, edgecolor=PAPER, linewidth=0.6, label=str(col),
               zorder=3)
        bottom += vals
    ax.set_xticks(x)
    ax.set_xticklabels([str(i) for i in tab.index], rotation=25, ha="right")
    ax.set_ylabel(ylabel)
    if xlabel:
        ax.set_xlabel(xlabel)
    ax.grid(axis="y", zorder=0)
    ax.set_axisbelow(True)
    ax.legend(ncol=min(4, len(tab.columns)), loc="upper center",
              bbox_to_anchor=(0.5, -0.22), frameon=False)
    fig.subplots_adjust(bottom=0.28)
    _acabar(fig, ax, destino, titulo=titulo, rodape=rodape)


def histograma_lados(esq, dir_, destino: Path, *,
                     titulo: str = "Distance by side of the node",
                     rodape: str = "TEXTURA corpus analysis"):
    aplicar_estilo()
    fig, ax = plt.subplots(figsize=(8.2, 4.0))
    esq, dir_ = np.asarray(esq, float), np.asarray(dir_, float)
    mx = int(max(esq.max() if len(esq) else 0,
                 dir_.max() if len(dir_) else 0, 1))
    bins = np.arange(0.5, mx + 1.5, 1)
    ax.hist(esq, bins=bins, alpha=0.75, color=ACCENT_COOL,
            edgecolor=PAPER, label="Left", zorder=3)
    ax.hist(dir_, bins=bins, alpha=0.65, color=ACCENT_WARM,
            edgecolor=PAPER, label="Right", zorder=3)
    ax.set_xlabel("Distance to node (tokens)")
    ax.set_ylabel("Co-occurrences")
    ax.grid(axis="y", zorder=0)
    ax.set_axisbelow(True)
    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1.0),
              frameon=True, fancybox=False, edgecolor=RULE,
              facecolor=PANEL, framealpha=0.95)
    fig.subplots_adjust(right=0.78)
    _acabar(fig, ax, destino, titulo=titulo, rodape=rodape)


def sankey_formas_docs(ligacoes, destino: Path, *,
                       titulo: str = "Sankey: formas → documentos",
                       subtitulo: str = "Largura da faixa = ocorrências",
                       rodape: str = "TEXTURA · análise de corpus",
                       eixo_esq: str = "Formas casadas",
                       eixo_dir: str = "Documentos",
                       max_formas: int = 8, max_docs: int = 10,
                       max_ligacoes: int = 36):
    """Sankey horizontal (ATLAS.ti style): formas -> documentos."""
    import pandas as pd
    from matplotlib.patches import Rectangle, PathPatch
    from matplotlib.path import Path as MPath

    def _enc(s, n=28):
        s = str(s or "").strip()
        return s if len(s) <= n else s[: n - 1] + "…"

    try:
        if isinstance(ligacoes, pd.DataFrame):
            d = ligacoes.copy()
            ren = {}
            for c in d.columns:
                cl = str(c).lower()
                if cl in {"forma", "form", "match"}:
                    ren[c] = "forma"
                elif cl in {"documento", "doc", "file"}:
                    ren[c] = "documento"
                elif cl in {"peso", "weight", "ocorrencias", "hits"}:
                    ren[c] = "peso"
            d = d.rename(columns=ren)
            if not {"forma", "documento", "peso"} <= set(d.columns):
                d = ligacoes.iloc[:, :3].copy()
                d.columns = ["forma", "documento", "peso"]
            else:
                d = d[["forma", "documento", "peso"]]
        else:
            d = pd.DataFrame(list(ligacoes), columns=["forma", "documento", "peso"])

        d = d.dropna()
        if d.empty:
            print("      [sankey] sem ligacoes", flush=True)
            return
        d["forma"] = d["forma"].astype(str)
        d["documento"] = d["documento"].astype(str).map(lambda x: _enc(x, 30))
        d["peso"] = pd.to_numeric(d["peso"], errors="coerce").fillna(0)
        d = d[d["peso"] > 0]
        if d.empty:
            return

        top_f = (d.groupby("forma")["peso"].sum()
                 .sort_values(ascending=False).head(max_formas).index)
        top_d = (d.groupby("documento")["peso"].sum()
                 .sort_values(ascending=False).head(max_docs).index)
        d = d[d["forma"].isin(top_f) & d["documento"].isin(top_d)]
        if d.empty:
            return

        agreg = (d.groupby(["forma", "documento"], as_index=False)["peso"].sum()
                 .sort_values("peso", ascending=False)
                 .head(max_ligacoes))
        left_tot = agreg.groupby("forma")["peso"].sum().sort_values(ascending=False)
        right_tot = agreg.groupby("documento")["peso"].sum().sort_values(ascending=False)

        n_side = max(len(left_tot), len(right_tot))
        alt = max(6.2, min(14.0, 0.55 * n_side + 2.8))
        aplicar_estilo()
        fig, ax = plt.subplots(figsize=(10.2, alt))
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")

        gap = 0.018 if n_side <= 12 else 0.012
        # etiquetas FORA das barras
        x0, x1 = 0.30, 0.38
        x2, x3 = 0.62, 0.70

        def _slots(tots):
            total = float(tots.sum()) or 1.0
            usable = 0.90 - gap * (len(tots) + 1)
            y = 0.94
            slots = {}
            for lab, w in tots.items():
                h = max(0.045, usable * (float(w) / total))
                y1 = y - h
                slots[lab] = (y1, y, h)
                y = y1 - gap
            return slots

        left_slots = _slots(left_tot)
        right_slots = _slots(right_tot)
        left_cursor = {k: left_slots[k][1] for k in left_slots}
        right_cursor = {k: right_slots[k][1] for k in right_slots}

        max_w = float(agreg["peso"].max()) or 1.0
        for r in agreg.sort_values("peso", ascending=False).itertuples(index=False):
            h = (left_slots[r.forma][2]) * (float(r.peso) / float(left_tot[r.forma]))
            y_l1 = left_cursor[r.forma]
            y_l0 = y_l1 - h
            left_cursor[r.forma] = y_l0
            h_r = (right_slots[r.documento][2]) * (
                float(r.peso) / float(right_tot[r.documento]))
            y_r1 = right_cursor[r.documento]
            y_r0 = y_r1 - h_r
            right_cursor[r.documento] = y_r0
            alpha = 0.20 + 0.50 * (float(r.peso) / max_w)
            mid = (x1 + x2) / 2
            verts = [
                (x1, y_l1), (mid, y_l1), (mid, y_r1), (x2, y_r1),
                (x2, y_r0), (mid, y_r0), (mid, y_l0), (x1, y_l0),
                (x1, y_l1),
            ]
            codes = [
                MPath.MOVETO,
                MPath.CURVE4, MPath.CURVE4, MPath.CURVE4,
                MPath.LINETO,
                MPath.CURVE4, MPath.CURVE4, MPath.CURVE4,
                MPath.CLOSEPOLY,
            ]
            ax.add_patch(PathPatch(
                MPath(verts, codes),
                facecolor=ACCENT_COOL, edgecolor="none", alpha=alpha, zorder=1))

        fs_l = 9.0 if len(left_slots) <= 8 else 7.5
        fs_r = 8.2 if len(right_slots) <= 10 else 7.0
        for lab, (y0, y1, h) in left_slots.items():
            ax.add_patch(Rectangle(
                (x0, y0), x1 - x0, h,
                facecolor=ACCENT, edgecolor=PAPER, linewidth=0.7, zorder=3))
            ax.text(x0 - 0.014, (y0 + y1) / 2, _enc(lab, 18),
                    ha="right", va="center", fontsize=fs_l, color=INK,
                    fontfamily="sans-serif", zorder=4)

        for lab, (y0, y1, h) in right_slots.items():
            ax.add_patch(Rectangle(
                (x2, y0), x3 - x2, h,
                facecolor=ACCENT_COOL, edgecolor=PAPER, linewidth=0.7, zorder=3))
            # fundo claro atrás do rótulo — evita texto perdido sobre as faixas
            ax.text(x3 + 0.014, (y0 + y1) / 2, _enc(lab, 30),
                    ha="left", va="center", fontsize=fs_r, color=INK,
                    fontfamily="sans-serif", zorder=5,
                    bbox=dict(boxstyle="round,pad=0.15",
                              facecolor=PAPER, edgecolor="none", alpha=0.92))

        ax.text(0.34, 0.015, eixo_esq, ha="center", va="bottom",
                fontsize=9, color=MUTED, fontfamily="sans-serif")
        ax.text(0.66, 0.015, eixo_dir, ha="center", va="bottom",
                fontsize=9, color=MUTED, fontfamily="sans-serif")

        fig.suptitle(titulo, fontsize=13, fontweight="semibold", color=INK,
                     x=0.5, y=0.985, ha="center")
        if subtitulo:
            sub = subtitulo if len(subtitulo) <= 100 else subtitulo[:97] + "…"
            # evitar dump de consultas enormes no PNG
            if sub.lower().startswith("consulta:") or sub.lower().startswith("query:"):
                sub = "Largura da faixa = nº de ocorrências"
            fig.text(0.5, 0.955, sub, ha="center", va="top",
                     fontsize=8.5, color=MUTED, fontfamily="sans-serif")
        if rodape:
            fig.text(0.5, 0.005, rodape, ha="center", va="bottom",
                     fontsize=7, color=MUTED, style="italic",
                     fontfamily="sans-serif")
        destino = Path(destino)
        destino.parent.mkdir(parents=True, exist_ok=True)
        fig.savefig(destino, dpi=160, bbox_inches="tight",
                    pad_inches=0.40, facecolor=PAPER)
        plt.close(fig)
        if not destino.exists():
            raise OSError(f"Sankey PNG nao foi criado: {destino}")
        print(f"      [sankey] PNG: {destino.resolve()}  "
              f"({destino.stat().st_size} bytes)", flush=True)
    except Exception as exc:  # noqa: BLE001
        print(f"      [sankey ERRO] {exc}", flush=True)
        # fallback ultra-simples: barras de fluxo
        try:
            _sankey_fallback_barras(ligacoes, Path(destino), titulo=titulo)
        except Exception as exc2:  # noqa: BLE001
            print(f"      [sankey FALLBACK ERRO] {exc2}", flush=True)
            raise


def _sankey_fallback_barras(ligacoes, destino: Path, *, titulo: str = "Sankey"):
    """Fallback se o Sankey curvo falhar: mapa de calor forma x documento."""
    import pandas as pd
    if isinstance(ligacoes, pd.DataFrame):
        d = ligacoes.copy()
        d.columns = ["forma", "documento", "peso"][:len(d.columns)]
    else:
        d = pd.DataFrame(list(ligacoes), columns=["forma", "documento", "peso"])
    piv = d.pivot_table(index="forma", columns="documento",
                        values="peso", aggfunc="sum", fill_value=0)
    aplicar_estilo()
    fig, ax = plt.subplots(figsize=(10, max(3.5, 0.35 * len(piv) + 1.5)))
    im = ax.imshow(piv.values, aspect="auto", cmap=_CMAP_DOCS)
    ax.set_yticks(range(len(piv.index)))
    ax.set_yticklabels(list(piv.index), fontsize=8)
    ax.set_xticks(range(len(piv.columns)))
    ax.set_xticklabels([str(c)[:28] for c in piv.columns],
                       rotation=55, ha="right", fontsize=7)
    ax.set_title(titulo + " (fallback heatmap)", loc="left")
    fig.colorbar(im, ax=ax, fraction=0.03)
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(destino, dpi=150, bbox_inches="tight", facecolor=PAPER)
    plt.close(fig)
    print(f"      [sankey fallback] PNG: {destino.resolve()}", flush=True)


def nuvem_palavras(frequencias: dict, destino: Path, *,
                   titulo: str = "Nuvem de formas casadas",
                   subtitulo: str = "Tamanho proporcional à frequência",
                   rodape: str = "TEXTURA · análise de corpus",
                   max_words: int = 80):
    """Nuvem de palavras a partir de {forma: contagem}."""
    if not frequencias:
        return
    try:
        from wordcloud import WordCloud
    except ImportError:
        # fallback mínimo sem a biblioteca
        aplicar_estilo()
        top = sorted(frequencias.items(), key=lambda kv: -kv[1])[:20]
        fig, ax = plt.subplots(figsize=(6.4, 3.2))
        labs, vals = zip(*top) if top else ([], [])
        ax.barh(list(labs)[::-1], list(vals)[::-1], color=ACCENT)
        _acabar(fig, ax, destino, titulo=titulo + " (fallback)", rodape=rodape)
        return

    aplicar_estilo()
    def _cor_nuvem(*_args, **_kwargs):
        # tons escuros (legibilidade); evita sage claro sobre fundo claro
        import random
        pal = ["#1F3D36", "#2F5D50", "#3A536B", "#5A4634", "#2A4A42", "#4A3A2A"]
        return random.choice(pal)

    wc = WordCloud(
        width=1200, height=560,
        background_color=PAPER,
        prefer_horizontal=0.78,
        max_words=max_words,
        relative_scaling=0.45,
        min_font_size=12,
        max_font_size=96,
        collocations=False,
        color_func=_cor_nuvem,
    ).generate_from_frequencies(frequencias)

    fig, ax = plt.subplots(figsize=(8.5, 4.2))
    ax.imshow(wc, interpolation="bilinear")
    ax.axis("off")
    _acabar(fig, ax, destino, titulo=titulo, subtitulo=subtitulo, rodape=rodape)


def correspondencias(lin, col, prop, destino: Path, *,
                     titulo: str = "Correspondence analysis",
                     rodape: str = "TEXTURA corpus analysis"):
    aplicar_estilo()
    fig, ax = plt.subplots(figsize=(8.2, 6.2))
    ax.axhline(0, color=RULE, lw=0.8, zorder=1)
    ax.axvline(0, color=RULE, lw=0.8, zorder=1)
    ax.scatter(lin["dim1"], lin["dim2"], s=36, color=ACCENT_COOL,
               alpha=0.9, zorder=3, label="Lexical types")
    for n, (x, y) in lin.iterrows():
        ax.annotate(str(n), (x, y), fontsize=8, color=INK,
                    xytext=(4, 3), textcoords="offset points")
    ax.scatter(col["dim1"], col["dim2"], s=70, marker="^",
               color=ACCENT_WARM, zorder=4, label="Syntactic relations")
    for n, (x, y) in col.iterrows():
        ax.annotate(str(n), (x, y), fontsize=8.5, color=ACCENT_WARM,
                    fontweight="semibold",
                    xytext=(4, -10), textcoords="offset points")
    p0 = prop[0] if prop else 0
    p1 = prop[1] if len(prop) > 1 else 0
    ax.set_xlabel(f"Dimension 1 ({p0}% inertia)")
    ax.set_ylabel(f"Dimension 2 ({p1}% inertia)")
    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1.0),
              frameon=True, fancybox=False, edgecolor=RULE, facecolor=PANEL)
    ax.grid(True, zorder=0, alpha=0.5)
    ax.set_axisbelow(True)
    fig.subplots_adjust(right=0.78)
    _acabar(fig, ax, destino, titulo=titulo, rodape=rodape)


def dendrograma(Z, labels, destino: Path, *,
                titulo: str = "Hierarchical clustering of collocates",
                rodape: str = "TEXTURA corpus analysis"):
    from scipy.cluster.hierarchy import dendrogram
    aplicar_estilo()
    fig, ax = plt.subplots(figsize=(6.4, min(6.0, max(3.0, 0.22 * len(labels)))))
    dendrogram(Z, labels=list(labels), orientation="right", ax=ax,
               color_threshold=0.7 * max(Z[:, 2]),
               above_threshold_color=MUTED)
    ax.set_xlabel("Ward distance")
    ax.tick_params(axis="y", labelsize=8)
    _acabar(fig, ax, destino, titulo=titulo, rodape=rodape)


# ---------------------------------------------------------------------------
# Compatibilidade com textura_analise (Prompt 3)
# ---------------------------------------------------------------------------

def barras_freq(freq: dict, destino: Path, *, titulo: str = "Frequencia"):
    """Wrapper: {rotulo: valor} -> barras horizontais."""
    if not freq:
        return
    labs = list(freq.keys())
    vals = [freq[k] for k in labs]
    barras_horizontais(labs, vals, destino, titulo=titulo,
                       xlabel="Ocorrencias", max_n=None)


def pares_forma_obra(df, *, col_forma="matched_form", col_obra="doc_id",
                     max_pares: int = 80):
    """Lista (forma, obra, peso) a partir da concordancia nuclear."""
    import pandas as pd
    d = df.copy()
    if col_obra not in d.columns:
        for alt in ("caminho_ficheiro", "caminho", "File", "Doc"):
            if alt in d.columns:
                col_obra = alt
                break
        else:
            col_obra = None
    if col_forma not in d.columns or col_obra is None:
        return []
    g = (d.groupby([col_forma, col_obra]).size()
         .reset_index(name="peso")
         .sort_values("peso", ascending=False)
         .head(max_pares))
    return list(g.itertuples(index=False, name=None))


def pares_termo_rel_pol(df, *, max_pares: int = 120):
    """Fluxo termo -> relacao -> polaridade (pares consecutivos)."""
    cols = ("canonical_term", "relacao_sintactica", "polaridade")
    if any(c not in df.columns for c in cols):
        return []
    d = df.dropna(subset=list(cols)).copy()
    out = []
    g1 = (d.groupby(["canonical_term", "relacao_sintactica"]).size()
          .reset_index(name="peso"))
    for r in g1.itertuples(index=False):
        out.append((str(r.canonical_term), str(r.relacao_sintactica), int(r.peso)))
    g2 = (d.groupby(["relacao_sintactica", "polaridade"]).size()
          .reset_index(name="peso"))
    for r in g2.itertuples(index=False):
        out.append((str(r.relacao_sintactica), str(r.polaridade), int(r.peso)))
    out.sort(key=lambda t: -t[2])
    return out[:max_pares]


def sankey_html(ligacoes, destino: Path, *, titulo: str = "Sankey"):
    """Exporta Sankey interactivo (Plotly) e tenta PNG via kaleido."""
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    if not ligacoes:
        destino.write_text(f"<html><body><p>Sem ligacoes — {titulo}</p></body></html>",
                           encoding="utf-8")
        return
    try:
        import plotly.graph_objects as go
    except ImportError:
        # fallback: PNG matplotlib
        png = destino.with_suffix(".png")
        sankey_formas_docs(ligacoes, png, titulo=titulo)
        destino.write_text(
            f"<html><body><p>{titulo}</p><p>Plotly indisponivel; "
            f"ver {png.name}</p></body></html>",
            encoding="utf-8")
        return

    nos, idx = [], {}
    src, tgt, val = [], [], []
    for a, b, w in ligacoes:
        for n in (a, b):
            if n not in idx:
                idx[n] = len(nos)
                nos.append(n)
        src.append(idx[a])
        tgt.append(idx[b])
        val.append(float(w))
    fig = go.Figure(data=[go.Sankey(
        node=dict(label=nos, pad=12, thickness=14),
        link=dict(source=src, target=tgt, value=val),
    )])
    fig.update_layout(title_text=titulo, font_size=11, height=560)
    fig.write_html(str(destino), include_plotlyjs="cdn")
    try:
        fig.write_image(str(destino.with_suffix(".png")))
    except Exception:
        pass
