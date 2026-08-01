#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""IDs de hit, fusão de janelas e citações entre documentos."""

from __future__ import annotations

import re
from collections import defaultdict
from typing import Any

import pandas as pd

from textura.tokenizacao import tokeniza


def _partilha_ngrama(a: list[str], b: list[str], n: int = 8) -> bool:
    """True se a e b partilham pelo menos n tokens contiguos."""
    if len(a) < n or len(b) < n:
        return False
    grams = {" ".join(a[i:i + n]) for i in range(len(a) - n + 1)}
    return any(" ".join(b[j:j + n]) in grams for j in range(len(b) - n + 1))




def occurrence_id_de(doc_id: str, source_matrix_row: int) -> str:
    """Identificador imutável da ocorrência mestra (= linha da matriz)."""
    return f"{doc_id}::ROW_{int(source_matrix_row)}"


def hit_key_de(occurrence_id: str, canonical_term: str, matched_form: str,
               off_no: int, off_termo: int) -> str:
    """Chave de hit distinta dentro de uma ocorrência mestra."""
    return (
        f"{occurrence_id}|{canonical_term}|{matched_form}|"
        f"{int(off_no)}|{int(off_termo)}"
    )


def atribuir_match_ids(res: pd.DataFrame) -> pd.DataFrame:
    """Numera M001… por ocorrência e preenche hit_key."""
    if res.empty:
        res = res.copy()
        res["match_id"] = pd.Series(dtype=str)
        res["hit_key"] = pd.Series(dtype=str)
        return res
    out = res.copy()
    out["match_id"] = ""
    out["hit_key"] = ""
    cols_ord = [c for c in ("off_termo", "canonical_term", "matched_form",
                            "idx_termo") if c in out.columns]
    for _occ, grp in out.groupby("texture_occurrence_id", sort=False):
        ordem = grp.sort_values(cols_ord, kind="mergesort").index
        for n, i in enumerate(ordem, 1):
            mid = f"M{n:03d}"
            out.at[i, "match_id"] = mid
            out.at[i, "hit_key"] = hit_key_de(
                str(out.at[i, "texture_occurrence_id"]),
                str(out.at[i, "canonical_term"]),
                str(out.at[i, "matched_form"]),
                int(out.at[i, "off_no"]),
                int(out.at[i, "off_termo"]),
            )
    return out


def deduplicar_hits_exactos(res: pd.DataFrame) -> pd.DataFrame:
    """Remove cópias exactas do mesmo hit_key (mantém a primeira)."""
    if res.empty or "hit_key" not in res.columns:
        return res
    return res.drop_duplicates(subset=["hit_key"], keep="first").copy()


def _score_sobrevivente_janela(row) -> tuple:
    """Maior = melhor candidato a representar um grupo de janelas."""
    nuc = 1 if bool(row.get("nuclear")) else 0
    ctx = len(str(row.get("contexto") or ""))
    # preferir a linha de matriz mais antiga (menor número)
    smr = -int(row.get("source_matrix_row") or 10**12)
    return (nuc, ctx, smr)


def _grupos_citacao_entre_docs(out: pd.DataFrame, ngrama: int = 8,
                               max_posting: int = 30) -> list[list]:
    """Grupos de janelas quase iguais abrangendo ``doc_id`` distintos.

    Critério: partilha de pelo menos um n-grama contíguo de ``ngrama``
    tokens (coerente com ``_partilha_ngrama``), com bloqueio por índice
    invertido para evitar comparação O(n²).
    """
    grams_idx: dict[str, list] = defaultdict(list)
    presentes: set = set()
    for i, row in out.iterrows():
        toks = [w for w, _ in tokeniza(str(row["contexto"]))]
        if len(toks) < ngrama:
            continue
        presentes.add(i)
        for k in range(len(toks) - ngrama + 1):
            grams_idx[" ".join(toks[k:k + ngrama])].append(i)

    parent = {i: i for i in presentes}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for ids in grams_idx.values():
        ids = sorted(set(ids))
        if len(ids) < 2 or len(ids) > max_posting:
            continue
        base = ids[0]
        for other in ids[1:]:
            ra, rb = find(base), find(other)
            if ra != rb:
                parent[ra] = rb

    grupos: dict = defaultdict(list)
    for i in presentes:
        grupos[find(i)].append(i)
    return [sorted(v) for v in grupos.values()
            if len(v) > 1 and out.loc[v, "doc_id"].nunique() > 1]


def _score_citacao(row) -> tuple:
    """Sobrevivente preferido num grupo de citação entre documentos."""
    caminho = str(row.get("caminho_ficheiro") or "")
    nome = caminho.replace("\\", "/").rsplit("/", 1)[-1]
    pref = 0
    if "todos os textos" in caminho:
        pref += 4                       # pasta catalogada do corpus
    if re.match(r"^\(\d{4}\)", nome):
        pref += 2                       # convenção '(AAAA)_Título'
    if "__" in nome:
        pref -= 1                       # cópias com sufixo de hash
    return (_score_sobrevivente_janela(row)[0], pref,
            *_score_sobrevivente_janela(row)[1:])


def fundir_janelas_e_marcar_duplicados(res: pd.DataFrame,
                                       ngrama: int = 8) -> pd.DataFrame:
    """Funde janelas KWIC deslocadas do *mesmo* hit; assinala passagens.

    Regras:
    - Mesmo ``doc_id`` + mesmo ``(canonical_term, matched_form)`` +
      partilha de n-grama → um sobrevivente; restantes
      ``nuclear=False``, ``motivo_exclusao=janela_sobreposta``.
    - Mesmo ``doc_id`` + n-grama partilhado com termos *diferentes* →
      apenas ``candidato_duplicado`` / ``grupo_passagem_id`` (não excluir:
      podem ser hits legítimos na mesma passagem / ocorrência mestra).
    - Citações quase iguais sob ``doc_id`` distintos → ``citacao_repetida``.
    """
    if res.empty or "doc_id" not in res.columns:
        return res
    out = res.copy()
    if "candidato_duplicado" not in out.columns:
        out["candidato_duplicado"] = ""
    if "grupo_passagem_id" not in out.columns:
        out["grupo_passagem_id"] = ""
    if "n_janelas_fundidas" not in out.columns:
        out["n_janelas_fundidas"] = 1
    out["n_janelas_fundidas"] = out["n_janelas_fundidas"].fillna(1).astype(int)

    # --- citações repetidas (texto quase igual, doc_ids distintos) -------
    # Antes: igualdade exacta da janela inteira, cega a variantes de
    # fronteira/OCR/edição. Agora: partilha de n-grama (o mesmo critério
    # já usado dentro do documento), com bloqueio por índice invertido.
    gid_cit = 0
    for membros in _grupos_citacao_entre_docs(out, ngrama=ngrama):
        gid_cit += 1
        gcit = f"C{gid_cit:04d}"
        tag = f"citacao_entre_doc_ids:{gcit}"
        for i in membros:
            prev = str(out.at[i, "candidato_duplicado"] or "")
            if "citacao_entre_doc_ids" not in prev:
                out.at[i, "candidato_duplicado"] = (
                    f"{prev}; {tag}" if prev else tag)
        # demover apenas hits lexicais repetidos entre documentos,
        # mantendo um sobrevivente por (canonical_term, matched_form)
        por_termo: dict = defaultdict(list)
        for i in membros:
            por_termo[(str(out.at[i, "canonical_term"]),
                       str(out.at[i, "matched_form"]))].append(i)
        for grupo in por_termo.values():
            if len(grupo) < 2:
                continue
            if len({out.at[i, "doc_id"] for i in grupo}) < 2:
                continue        # repetição interna: tratada pela fusão
            ranked = sorted(grupo, key=lambda i: _score_citacao(out.loc[i]),
                            reverse=True)
            for i in ranked[1:]:
                out.at[i, "nuclear"] = False
                if not out.at[i, "motivo_exclusao"]:
                    out.at[i, "motivo_exclusao"] = "citacao_repetida"

    # --- fusão / passagem no mesmo documento -----------------------------
    por_doc: dict = defaultdict(list)
    for i, row in out.iterrows():
        por_doc[row["doc_id"]].append(i)

    gid_pass = 0
    gid_jan = 0
    for idxs in por_doc.values():
        if len(idxs) < 2:
            continue
        assin = {
            i: [w for w, _ in tokeniza(str(out.at[i, "contexto"]))]
            for i in idxs
        }
        # Union-find leve por partilha de n-grama (passagem)
        parent = {i: i for i in idxs}

        def find(x):
            while parent[x] != x:
                parent[x] = parent[parent[x]]
                x = parent[x]
            return x

        def union(a, b):
            ra, rb = find(a), find(b)
            if ra != rb:
                parent[rb] = ra

        for a_i, a in enumerate(idxs):
            for b in idxs[a_i + 1:]:
                # Hits da mesma ocorrência mestra partilham o contexto por
                # definição — não são janelas KWIC deslocadas.
                if (str(out.at[a, "texture_occurrence_id"])
                        == str(out.at[b, "texture_occurrence_id"])):
                    continue
                if _partilha_ngrama(assin[a], assin[b], ngrama):
                    union(a, b)

        clusters: dict = defaultdict(list)
        for i in idxs:
            clusters[find(i)].append(i)

        for membros in clusters.values():
            if len(membros) < 2:
                continue
            gid_pass += 1
            gpass = f"P{gid_pass:04d}"
            for i in membros:
                out.at[i, "grupo_passagem_id"] = gpass
                prev = str(out.at[i, "candidato_duplicado"] or "")
                tag = f"passagem_sobreposta:{gpass}"
                if tag not in prev:
                    out.at[i, "candidato_duplicado"] = (
                        f"{prev}; {tag}" if prev else tag)

            # Dentro da passagem: fundir só hits lexicais idênticos
            por_termo: dict = defaultdict(list)
            for i in membros:
                chave = (
                    str(out.at[i, "canonical_term"]),
                    str(out.at[i, "matched_form"]),
                )
                por_termo[chave].append(i)
            for grupo in por_termo.values():
                if len(grupo) < 2:
                    continue
                gid_jan += 1
                gjan = f"J{gid_jan:04d}"
                ranked = sorted(
                    grupo,
                    key=lambda i: _score_sobrevivente_janela(out.loc[i]),
                    reverse=True,
                )
                keep = ranked[0]
                out.at[keep, "n_janelas_fundidas"] = len(ranked)
                tag_keep = f"janela_sobreposta:{gjan}"
                prev_k = str(out.at[keep, "candidato_duplicado"] or "")
                if tag_keep not in prev_k:
                    out.at[keep, "candidato_duplicado"] = (
                        f"{prev_k}; {tag_keep}" if prev_k else tag_keep)
                for i in ranked[1:]:
                    out.at[i, "nuclear"] = False
                    if not out.at[i, "motivo_exclusao"]:
                        out.at[i, "motivo_exclusao"] = "janela_sobreposta"
                    out.at[i, "n_janelas_fundidas"] = 1
                    prev = str(out.at[i, "candidato_duplicado"] or "")
                    tag = f"janela_sobreposta:{gjan}"
                    if tag not in prev:
                        out.at[i, "candidato_duplicado"] = (
                            f"{prev}; {tag}" if prev else tag)
    return out


def agregar_ocorrencias(res: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por ``texture_occurrence_id`` (nível 1)."""
    if res.empty or "texture_occurrence_id" not in res.columns:
        return pd.DataFrame()
    linhas = []
    for occ_id, grp in res.groupby("texture_occurrence_id", sort=False):
        # contexto canónico: o mais longo (melhor exibição)
        ctxs = grp["contexto"].astype(str)
        i_ctx = ctxs.str.len().idxmax()
        termos = sorted({str(t) for t in grp["canonical_term"].tolist() if t})
        formas = sorted({str(t) for t in grp["matched_form"].tolist() if t})
        nucs = grp["nuclear"].map(
            lambda v: v is True or str(v).lower() in {"true", "1", "sim"})
        motivos = [str(m) for m in grp["motivo_exclusao"].tolist()
                   if m and str(m).strip() and str(m).lower() != "nan"]
        linhas.append({
            "texture_occurrence_id": occ_id,
            "source_matrix_row": int(grp["source_matrix_row"].iloc[0]),
            "doc_id": grp["doc_id"].iloc[0],
            "caminho_ficheiro": grp["caminho_ficheiro"].iloc[0],
            "url": grp["url"].iloc[0] if "url" in grp.columns else "",
            "no": grp["no"].iloc[0] if "no" in grp.columns else "",
            "matched_terms": "; ".join(termos),
            "matched_forms": "; ".join(formas),
            "n_matches": int(len(grp)),
            "n_matches_nucleares": int(nucs.sum()),
            "nuclear": bool(nucs.any()),
            "canonical_context": ctxs.loc[i_ctx],
            "grupo_passagem_id": (
                str(grp["grupo_passagem_id"].iloc[0])
                if "grupo_passagem_id" in grp.columns else ""),
            "match_ids": "; ".join(
                str(x) for x in grp.sort_values("match_id")["match_id"]),
            "motivos_exclusao": "; ".join(sorted(set(motivos))),
            "dominio": (
                grp["dominio"].iloc[0] if "dominio" in grp.columns else ""),
        })
    cols = [
        "texture_occurrence_id", "source_matrix_row", "doc_id",
        "caminho_ficheiro", "url", "no", "matched_terms", "matched_forms",
        "n_matches", "n_matches_nucleares", "nuclear", "canonical_context",
        "grupo_passagem_id", "match_ids", "motivos_exclusao", "dominio",
    ]
    return pd.DataFrame(linhas)[cols]
