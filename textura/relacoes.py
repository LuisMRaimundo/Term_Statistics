#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Classificação sintáctica (spaCy / heurística) e avisos de revisão."""

from __future__ import annotations

import re
import sys
from typing import Any

import numpy as np
import pandas as pd

from textura.config import MODALIDADE, RELACOES_NUCLEARES
from textura.revisao import etiqueta_exacta, etiqueta_prefixada, juntar_etiquetas
from textura.tokenizacao import anota_sintaxe, tokeniza


def _token_em(doc, offset: int):
    """Token cujo intervalo de caracteres contém o offset dado."""
    for t in doc:
        if t.idx <= offset < t.idx + len(t.text):
            return t
    return None


def _resultado_rel(rel: str, governante: str, percurso: str,
                   orientacao: str = "", *,
                   nucleo_prop: str = "", revisao: str = "",
                   matched: str = "") -> dict:
    nuclear = rel in RELACOES_NUCLEARES
    # R5: governante nunca é o próprio termo
    gov = governante
    if matched and gov and gov.lower() == matched.lower():
        gov = nucleo_prop or ""
    if not orientacao:
        if nuclear and rel in ("nominal_composto", "nominal_genitiva",
                               "adverbial"):
            orientacao = "no_sobre_termo"
        elif nuclear:
            orientacao = "termo_sobre_no"
        else:
            orientacao = "termo_sobre_outro"
    return {
        "relacao_sintactica": rel,
        "orientacao": orientacao,
        "governante": gov,
        "percurso_dep": percurso,
        "nuclear": nuclear,
        "motivo_exclusao": "" if nuclear else (rel or "indeterminada"),
        "nucleo_da_propriedade": nucleo_prop or ("" if not nuclear else ""),
        "revisao_sugerida": revisao,
    }


def _mesmo_token(a, b) -> bool:
    """Comparação estável entre tokens spaCy (evitar `is`, frágil)."""
    return a is not None and b is not None and a.i == b.i and a.doc is b.doc


def _no_no_subtree(tok, t_no) -> bool:
    return _mesmo_token(tok, t_no) or any(_mesmo_token(x, t_no)
                                          for x in tok.subtree)


def _gov_efectivo(tok) -> str:
    """Núcleo sintáctico do termo (head após conj), nunca o próprio texto."""
    base = tok
    while base.dep_ == "conj" and base.head.i != base.i:
        base = base.head
    h = base.head
    if h.i == base.i:
        return ""
    return h.text


_PREPS_GENITIVO_OMISSAO = frozenset({"of", "de", "in", "on"})
_PREPS_ASSOC_OMISSAO = frozenset({"with"})


def _tem_complemento_genitivo(
    termo, t_no, preps_genitivo: frozenset[str] | None = None,
) -> bool:
    """True se T rege prep cujo objecto é N (qualquer papel de T)."""
    preps = preps_genitivo if preps_genitivo is not None else _PREPS_GENITIVO_OMISSAO
    base = termo
    while base.dep_ == "conj" and base.head.i != base.i:
        base = base.head
    for node in (termo, base):
        for child in node.children:
            if child.dep_ == "prep" and child.text.lower() in preps:
                for gc in child.children:
                    if gc.dep_ in ("pobj", "nmod") and _no_no_subtree(
                            gc, t_no):
                        return True
    # N sob prep regida por T
    if t_no.dep_ in ("pobj", "nmod"):
        prep = t_no.head
        regente = prep.head if prep.dep_ == "prep" else prep
        if _mesmo_token(regente, termo) or _mesmo_token(regente, base):
            return True
    return False


def _amod_coordenado_do_no(t_te, t_no) -> bool:
    """T é adjectivo coordenado com outro amod cujo núcleo é N."""
    if t_no.pos_ not in ("NOUN", "PROPN"):
        return False
    # irmãos amod de N
    for am in t_no.children:
        if am.dep_ != "amod":
            continue
        if _mesmo_token(am, t_te):
            return True
        # cadeia de conjunção
        conj = list(am.conjuncts) if hasattr(am, "conjuncts") else []
        if any(_mesmo_token(c, t_te) for c in conj):
            return True
        # T.conj aponta para am, ou vice-versa
        b = t_te
        while b.dep_ == "conj" and b.head.i != b.i:
            if _mesmo_token(b.head, am):
                return True
            b = b.head
    # heurística linear: ADJ ... and/or ADJ NOUN=N
    if t_te.pos_ == "ADJ" and t_te.i < t_no.i:
        entre = list(t_te.doc[t_te.i + 1:t_no.i])
        if entre and all(
                t.text.lower() in {"and", "or", "but", ",", "the", "a", "an"}
                or t.dep_ in ("cc", "punct", "det", "amod", "conj")
                for t in entre):
            return True
    return False


_RX_TEXTURAL = re.compile(r"^textur", re.IGNORECASE)


def _e_token_textural(tok) -> bool:
    """Token pertencente ao paradigma lexical de 'textura'."""
    return bool(_RX_TEXTURAL.match(tok.text))


def _coordenacao_heterogenea(t_no) -> str:
    """Nome não textural coordenado com N ('of texture and dynamics',
    'of colors and textures'). Devolve o texto do conjunto ou ''."""
    cands = list(getattr(t_no, "conjuncts", []) or [])
    b = t_no
    while b.dep_ == "conj" and b.head.i != b.i:
        cands.append(b.head)
        b = b.head
    cands.extend(c for c in t_no.children if c.dep_ == "conj")
    for c in cands:
        if c.pos_ in ("NOUN", "PROPN") and not _e_token_textural(c):
            return c.text
    return ""


def _associativa_heterogenea(
    t_te, t_no, preps_associativa: frozenset[str] | None = None,
) -> str:
    """T rege prep associativa cujo objecto é nome não textural, estando N
    fora desse complemento ('textures combined with lyrics')."""
    preps = (
        preps_associativa if preps_associativa is not None
        else _PREPS_ASSOC_OMISSAO
    )
    base = t_te
    while base.dep_ == "conj" and base.head.i != base.i:
        base = base.head
    for node in {t_te, base}:
        for child in node.children:
            if child.dep_ == "prep" and child.text.lower() in preps:
                for gc in child.children:
                    if gc.dep_ in ("pobj", "nmod") \
                            and gc.pos_ in ("NOUN", "PROPN") \
                            and not _e_token_textural(gc) \
                            and not _no_no_subtree(gc, t_no):
                        return gc.text
    return ""


def relacao_dependencia(
    doc, off_no: int, off_termo: int, *,
    preps_genitivo: frozenset[str] | None = None,
    preps_associativa: frozenset[str] | None = None,
) -> dict:
    """Classificação sintáctica + sinalização de coordenação heterogénea.

    A taxonomia nuclear é a de ``_relacao_dependencia_base``; este
    invólucro acrescenta, sem alterar ``nuclear``, avisos de revisão
    quando o vínculo termo–textura passa por coordenação com elementos
    não texturais ou por associação — os padrões que a revisão
    humana demonstrou serem os falsos positivos mais frequentes.
    """
    res = _relacao_dependencia_base(
        doc, off_no, off_termo, preps_genitivo=preps_genitivo)
    if not res.get("nuclear"):
        return res
    t_no, t_te = _token_em(doc, off_no), _token_em(doc, off_termo)
    if t_no is None or t_te is None:
        return res
    avisos = []
    coord = _coordenacao_heterogenea(t_no)
    if coord:
        avisos.append(etiqueta_prefixada("coordenacao_heterogenea", coord))
    assoc = _associativa_heterogenea(
        t_te, t_no, preps_associativa=preps_associativa)
    if assoc:
        avisos.append(etiqueta_prefixada("associativa_com_nao_textural", assoc))
    if avisos:
        prev = res.get("revisao_sugerida") or ""
        res["revisao_sugerida"] = juntar_etiquetas(prev, *avisos)
    return res


def _relacao_dependencia_base(
    doc, off_no: int, off_termo: int, *,
    preps_genitivo: frozenset[str] | None = None,
) -> dict:
    """Classifica a relação sintáctica (taxonomia nuclear / não nuclear)."""
    t_no, t_te = _token_em(doc, off_no), _token_em(doc, off_termo)
    if t_no is None or t_te is None:
        return _resultado_rel("indeterminada", "", "", "")

    matched = t_te.text
    percurso = f"{t_te.text}/{t_te.dep_}->{t_te.head.text}"
    era_conj = t_te.dep_ == "conj"
    base = t_te
    while base.dep_ == "conj" and base.head.i != base.i:
        base = base.head
    gov0 = _gov_efectivo(t_te)

    # R4.2 — genitiva ANTES de qualquer teste de governante directo
    if _tem_complemento_genitivo(t_te, t_no, preps_genitivo=preps_genitivo):
        rev = (
            etiqueta_exacta("genitiva_por_complemento")
            if t_te.dep_ not in ("pobj", "nmod", "") else ""
        )
        return _resultado_rel(
            "nominal_genitiva", gov0, percurso, "no_sobre_termo",
            nucleo_prop=t_no.text, revisao=rev, matched=matched)

    # --- adverbial: «texturally uniform» ----------------------------------
    if t_no.dep_ == "advmod" and _mesmo_token(t_no.head, t_te):
        return _resultado_rel(
            "adverbial", gov0 or t_te.head.text, percurso, "no_sobre_termo",
            nucleo_prop=t_no.text, matched=matched)

    # --- adverbial de grau / verbal ---------------------------------------
    if base.dep_ == "advmod" and base.head.pos_ in ("ADJ", "ADV"):
        if not _mesmo_token(base.head, t_no):
            return _resultado_rel(
                "adverbial_de_grau", base.head.text, percurso,
                "termo_sobre_outro", nucleo_prop=base.head.text,
                matched=matched)
    if base.dep_ == "advmod" and base.head.pos_ == "VERB":
        return _resultado_rel(
            "adverbial_verbal", base.head.text, percurso,
            "termo_sobre_outro", nucleo_prop=base.head.text, matched=matched)

    # --- predicação secundária --------------------------------------------
    if base.dep_ in ("oprd", "xcomp"):
        pred = base.head
        objs = [c for c in pred.children
                if c.dep_ in ("dobj", "obj", "oprd", "attr")]
        if any(_no_no_subtree(o, t_no) for o in objs) or any(
                _no_no_subtree(s, t_no) for s in pred.children
                if s.dep_ in ("nsubj", "nsubjpass")):
            return _resultado_rel(
                "predicativa_secundaria", gov0 or pred.text, percurso,
                "termo_sobre_no", nucleo_prop=t_no.text, matched=matched)

    # --- predicação: acomp/attr — reclassificar após conj -----------------
    if base.dep_ in ("acomp", "attr"):
        pred = base.head
        sujeitos = [c for c in pred.children
                    if c.dep_ in ("nsubj", "nsubjpass")]
        if sujeitos:
            s = sujeitos[0]
            if _no_no_subtree(s, t_no):
                return _resultado_rel(
                    "predicativa", gov0 or pred.text, percurso,
                    "termo_sobre_no", nucleo_prop=t_no.text, matched=matched)
            # attr de factor ... is uniformity of texture — já coberto por genitiva
            return _resultado_rel(
                "incidental", s.text, percurso, "termo_sobre_outro",
                nucleo_prop=s.text, matched=matched)

    # --- atributiva directa (também após normalização de conj) ------------
    if base.dep_ in ("amod", "appos") and _mesmo_token(base.head, t_no):
        rev = etiqueta_exacta("atributiva_via_conj") if era_conj else ""
        return _resultado_rel(
            "atributiva", gov0 or t_no.text, percurso, "termo_sobre_no",
            nucleo_prop=t_no.text, revisao=rev, matched=matched)

    # R4.3 — modificação partilhada / coordenação adjectival
    if _amod_coordenado_do_no(t_te, t_no):
        return _resultado_rel(
            "atributiva", gov0 or t_no.text, percurso, "termo_sobre_no",
            nucleo_prop=t_no.text,
            revisao=etiqueta_exacta("atributiva_coordenada"),
            matched=matched)

    # --- nominal composto: «textural diversity» ---------------------------
    if t_no.dep_ in ("amod", "compound") and _mesmo_token(t_no.head, t_te):
        return _resultado_rel(
            "nominal_composto", gov0 or t_te.head.text, percurso,
            "no_sobre_termo", nucleo_prop=t_no.text, matched=matched)

    if base.dep_ == "compound" and _mesmo_token(base.head, t_no):
        return _resultado_rel(
            "atributiva", gov0 or t_no.text, percurso, "termo_sobre_no",
            nucleo_prop=t_no.text, matched=matched)

    # --- coordenação entre constituintes distintos ------------------------
    if era_conj:
        gov = base.head if base.head.i != base.i else base
        if not _mesmo_token(gov, t_no) and not _no_no_subtree(gov, t_no):
            return _resultado_rel(
                "coordenada", gov.text, percurso, "termo_sobre_outro",
                nucleo_prop=gov.text, matched=matched)

    gov = base.head
    return _resultado_rel(
        "incidental", gov.text if gov is not None else "",
        percurso, "termo_sobre_outro",
        nucleo_prop=gov.text if gov is not None else "", matched=matched)


def _escopo_negacao(doc, off_termo: int, off_no: int) -> str:
    """Negação: 'directo' | 'indirecto' | 'nao'."""
    t_te = _token_em(doc, off_termo)
    if t_te is None:
        return "nao"
    base = t_te
    while base.dep_ == "conj" and base.head.i != base.i:
        base = base.head
    # predicado em que o termo participa
    pred = base.head if base.dep_ in (
        "acomp", "attr", "oprd", "xcomp", "amod", "advmod") else t_te.head

    negadores = []
    for tok in doc:
        w = tok.text.lower().replace("'", "'")
        if tok.dep_ == "neg" or w in {"not", "never", "n't"} or \
                tok.lemma_.lower() in {"not", "never"}:
            negadores.append(tok)

    if not negadores:
        return "nao"

    for neg in negadores:
        # neg domina o predicado do termo?
        cabeças = {neg.head}
        if pred in list(neg.head.subtree) or _mesmo_token(neg.head, pred) \
                or _mesmo_token(neg.head, t_te) or t_te in list(neg.head.subtree):
            # directo: o head do neg é o mesmo predicado do termo
            if _mesmo_token(neg.head, pred) or _mesmo_token(neg.head, base.head):
                # e o termo é complemento/modificador desse predicado
                if base.dep_ in ("acomp", "attr", "oprd", "xcomp") or \
                        (base.dep_ == "amod" and _mesmo_token(base.head,
                                                              _token_em(doc, off_no) or base.head)):
                    # amod sob N: neg no verbo superior = nao/indirecto
                    if base.dep_ == "amod":
                        # «is not uniform» (acomp) vs «does not create uniform texture»
                        if neg.head.pos_ == "AUX" or neg.head.lemma_ in {
                                "be", "remain", "become", "seem", "appear"}:
                            if _mesmo_token(base.head, _token_em(doc, off_no)):
                                # texture is not uniform — termo é acomp, não amod
                                pass
                        if not _mesmo_token(neg.head, base.head):
                            # neg noutro verbo: «does not ... uniform texture»
                            return "indirecto"
                    else:
                        return "directo"
            # neg em cláusula superior que contém o NP
            if t_te in list(neg.head.subtree):
                # se há verbo interveniente entre neg.head e o termo
                if neg.head.pos_ == "VERB" and not _mesmo_token(neg.head, pred):
                    return "indirecto"
                if _mesmo_token(neg.head, pred):
                    return "directo"
    # negação na frase mas fora do predicado do termo
    for neg in negadores:
        if t_te in list(neg.head.subtree):
            return "nao"
    return "nao"


def anota_com_spacy(
    res: pd.DataFrame, modelo: str, *,
    obrigatorio: bool = True,
    preps_genitivo: frozenset[str] | None = None,
    preps_associativa: frozenset[str] | None = None,
) -> pd.DataFrame:
    """Classifica cada linha pela árvore de dependências (spaCy).

    Língua/modelo são da execução (``--lingua`` / registo); não há selecção
    por linha. Se o modelo faltar e ``obrigatorio`` for False, degrada para
    a classificação já presente / heurística do chamador.
    """
    try:
        import spacy
    except ImportError:
        msg = ("spaCy nao instalado. Execute: pip install 'spacy>=3.7' "
               f"&& python -m spacy download {modelo}")
        if obrigatorio:
            raise SystemExit(msg) from None
        print("AVISO: " + msg + " - a usar heuristica.", file=sys.stderr)
        return anota_com_heuristica(res)
    try:
        nlp = spacy.load(modelo, disable=["ner", "lemmatizer", "textcat"])
    except OSError:
        msg = (f"modelo spaCy '{modelo}' indisponivel. Execute:\n"
               f"  python -m spacy download {modelo}")
        if obrigatorio:
            raise SystemExit(msg) from None
        print("AVISO: " + msg + " - a usar heuristica.", file=sys.stderr)
        return anota_com_heuristica(res)

    contextos = res["contexto"].astype(str).unique().tolist()
    print(f"      a analisar sintaxe de {len(contextos)} contextos unicos "
          f"(modelo={modelo}) ...", flush=True)
    docs = {c: d for c, d in zip(contextos, nlp.pipe(contextos, batch_size=64))}

    keys = ("relacao_sintactica", "orientacao", "governante", "percurso_dep",
            "nuclear", "motivo_exclusao", "nucleo_da_propriedade",
            "revisao_sugerida", "negado", "modalizado")
    cols = {k: [] for k in keys}
    for t in res.itertuples(index=False):
        doc = docs[str(t.contexto)]
        r = relacao_dependencia(
            doc, int(t.off_no), int(t.off_termo),
            preps_genitivo=preps_genitivo,
            preps_associativa=preps_associativa,
        )
        for k in ("relacao_sintactica", "orientacao", "governante",
                  "percurso_dep", "nuclear", "motivo_exclusao",
                  "nucleo_da_propriedade", "revisao_sugerida"):
            cols[k].append(r[k])
        cols["negado"].append(_escopo_negacao(
            doc, int(t.off_termo), int(t.off_no)))
        # R8: modal no escopo local do termo (±5 tokens ou ancestral)
        t_te = _token_em(doc, int(t.off_termo))
        mod = False
        if t_te is not None:
            for tok in doc:
                w = tok.text.lower()
                if w in MODALIDADE or tok.lemma_.lower() in MODALIDADE:
                    if abs(tok.i - t_te.i) <= 5:
                        mod = True
                        break
                    if t_te in list(tok.subtree) or tok in list(t_te.ancestors):
                        mod = True
                        break
        cols["modalizado"].append(mod)

    res = res.copy()
    for k, vals in cols.items():
        res[k] = vals
    res["fonte_classificacao"] = "dependencias"
    # R10: nao duplicar relacao / caminho_dep / atribuicao
    for drop in ("relacao", "caminho_dep", "atribuicao", "caminho"):
        if drop in res.columns:
            res = res.drop(columns=[drop])
    return res


def anota_com_heuristica(res: pd.DataFrame) -> pd.DataFrame:
    """Preenche taxonomia reduzida sem spaCy; fonte_classificacao=heuristica."""
    res = res.copy()
    rels, nucs, mots, oris, govs = [], [], [], [], []
    grads, mods, negs = [], [], []
    for t in res.itertuples(index=False):
        toks = tokeniza(str(t.contexto))
        # localizar índices por offset
        i_no = next((i for i, (_, o) in enumerate(toks)
                     if o == int(t.off_no)), None)
        i_te = next((i for i, (_, o) in enumerate(toks)
                     if o == int(t.off_termo)), None)
        if i_no is None or i_te is None:
            rels.append("indeterminada"); nucs.append(False)
            mots.append("indeterminada"); oris.append(""); govs.append("")
            grads.append(False); mods.append(False); negs.append(None)
            continue
        neg, grad, mod, rel = anota_sintaxe(
            toks, i_no, i_te, str(t.contexto))
        nuclear = rel in RELACOES_NUCLEARES
        rels.append(rel); nucs.append(nuclear)
        mots.append("" if nuclear else rel)
        oris.append("termo_sobre_no" if nuclear else "")
        govs.append(toks[i_no][0] if nuclear else "")
        grads.append(grad); mods.append(mod); negs.append(neg)
    res["relacao_sintactica"] = rels
    res["relacao"] = rels
    res["nuclear"] = nucs
    res["motivo_exclusao"] = mots
    res["orientacao"] = oris
    res["governante"] = govs
    res["percurso_dep"] = ""
    res["caminho_dep"] = ""
    res["fonte_classificacao"] = "heuristica"
    res["graduado"] = grads
    res["modalizado"] = mods
    res["negado"] = negs
    res["atribuicao"] = np.where(res["nuclear"], "genuína", "incidental")
    return res
