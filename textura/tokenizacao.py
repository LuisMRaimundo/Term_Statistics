#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Normalização, tokenização, NEAR pairing e consulta booleana."""

from __future__ import annotations

import re
from collections import Counter
from typing import Iterable

from textura.config import (
    ABREVIATURAS, CAMPO, COPULAS, GRADUACAO, MODALIDADE, NEGACAO, NOS,
    RE_FIM_FRASE, RE_HIFEN_QUEBRA, RE_TOKEN, RELACOES_NUCLEARES,
)


def normaliza(texto: str) -> str:
    """Repara quebras de linha hifenizadas ('propor- tions' -> 'proportions')."""
    texto = str(texto).replace("\n", " ").replace("\r", " ")
    texto = RE_HIFEN_QUEBRA.sub(r"\1\2", texto)
    return re.sub(r"\s+", " ", texto)


def fronteiras_frase(texto: str) -> list[int]:
    """Devolve os offsets de carácter onde termina uma frase.

    Um ponto não fecha frase quando: (i) a palavra que o precede é uma
    abreviatura conhecida; (ii) está entre dígitos (decimal); (iii) é
    parte de reticências. O ponto-e-vírgula é tratado como fronteira,
    por separar predicações autónomas — decisão a declarar no protocolo.
    """
    limites = []
    for m in RE_FIM_FRASE.finditer(texto + " "):
        i = m.start()
        anterior = texto[max(0, i - 30):i]
        palavra = re.search(r"([A-Za-zÀ-ÿ.]+)$", anterior)
        if palavra:
            cand = palavra.group(1).lower().rstrip(".")
            if cand in ABREVIATURAS:
                continue
        if i > 0 and texto[i - 1].isdigit() and i + 1 < len(texto) and texto[i + 1:i + 2].isdigit():
            continue
        if texto[i:i + 3] == "...":
            continue
        limites.append(i)
    return limites


def indice_frase(pos: int, limites: list[int]) -> int:
    """Número da frase a que pertence um offset de carácter."""
    return sum(1 for L in limites if L < pos)


def tokeniza(texto: str) -> list[tuple[str, int]]:
    """Lista de (token minúsculo, offset de carácter)."""
    return [(m.group(0).lower(), m.start()) for m in RE_TOKEN.finditer(texto)]


# ---------------------------------------------------------------------------
# 3. OPERADOR NEAR/x
# ---------------------------------------------------------------------------

def _rx_palavra(p: str) -> re.Pattern:
    """Compila uma palavra, com truncatura à direita (*) e/ou à esquerda.

    Truncatura só à direita («uniform*», «static*»): ancora no início do
    token e NÃO atravessa hífen — assim «static*» não casa «static-chordal».
    Truncatura à esquerda («*uniform»): permite hífen e casa o composto
    pelo núcleo final («di-uniform»).
    """
    esq = p.startswith("*")
    dir_ = p.endswith("*")
    nucleo = re.escape(p.strip("*"))
    if esq and dir_:
        star = r"[\w\-]{0,20}"
        return re.compile("^" + star + nucleo + star + "$")
    if esq:
        return re.compile(r"^(?:[\w\-]*-)?" + nucleo + r"$")
    if dir_:
        # sem hífen no wildcard — A15 / T2
        return re.compile("^" + nucleo + r"\w{0,20}$")
    return re.compile("^" + nucleo + "$")


def compila_campo(campo: dict[str, list[str]]):
    """Converte truncaturas em sequências de padrões de token."""
    saida = []
    for etiqueta, padroes in campo.items():
        seqs = [[_rx_palavra(w) for w in p.split()] for p in padroes]
        seqs.sort(key=len, reverse=True)
        saida.append((etiqueta, seqs))
    return saida


def _casa_em(tokens, j, seq) -> int:
    """Devolve o comprimento da sequência se casar em j; senão 0."""
    if j + len(seq) > len(tokens):
        return 0
    for k, rx in enumerate(seq):
        if not rx.match(tokens[j + k][0]):
            return 0
    return len(seq)


def indices_no(tokens, nos_validos: set[str]) -> list[int]:
    """Índices de todas as ocorrências do nó no contexto tokenizado."""
    return [i for i, (w, _) in enumerate(tokens) if w in nos_validos]


def melhor_par_tokens(tokens, idxs_no: list[int], idx_termo: int,
                      n_palavras: int, limites: list[int],
                      mesma_frase: bool = True):
    """Como melhor_par, com verificação de frase sobre offsets reais."""
    span = list(range(idx_termo, idx_termo + n_palavras))
    melhor = None  # (dist, pos_esq, i_no)
    for i_no in idxs_no:
        if i_no in span:
            continue
        if mesma_frase and limites:
            f_no = indice_frase(tokens[i_no][1], limites)
            if any(indice_frase(tokens[k][1], limites) != f_no for k in span):
                continue
        dist = min(abs(k - i_no) for k in span)
        pos_esq = min(i_no, idx_termo)
        cand = (dist, pos_esq, i_no)
        if melhor is None or cand < melhor:
            melhor = cand
    if melhor is None:
        return None
    dist, _, i_no = melhor
    lado = "esq" if idx_termo < i_no else ("dir" if idx_termo > i_no else "—")
    return i_no, idx_termo, dist, lado


def encontra_termos(tokens, campo_compilado) -> list[dict]:
    """Todas as ocorrências do campo lexical no contexto (sem fixar nó)."""
    achados = []
    vistos = set()  # (etiqueta, idx_termo, n_palavras)
    for j in range(len(tokens)):
        for etiqueta, seqs in campo_compilado:
            casou = 0
            for seq in seqs:
                casou = _casa_em(tokens, j, seq)
                if casou:
                    break
            if not casou:
                continue
            chave = (etiqueta, j, casou)
            if chave in vistos:
                continue
            vistos.add(chave)
            forma = " ".join(tokens[k][0] for k in range(j, j + casou))
            achados.append({
                "termo_tipo": etiqueta,
                "termo_forma": forma,
                "idx_termo": j,
                "n_palavras": casou,
                "forma_em_composto": "-" in forma,
            })
            break  # uma etiqueta por posição (padrão mais longo já ordenado)
    return achados


def procura_near(tokens, no_idx, campo_compilado, n, limites, mesma_frase=True):
    """Co-ocorrências dentro de NEAR/n de um nó fixo (legado / banda)."""
    achados = []
    for t in encontra_termos(tokens, campo_compilado):
        par = melhor_par_tokens(
            tokens, [no_idx], t["idx_termo"], t["n_palavras"],
            limites, mesma_frase=mesma_frase)
        if par is None:
            continue
        i_no, idx_te, dist, lado = par
        if dist > n:
            continue
        achados.append({
            "termo_tipo": t["termo_tipo"],
            "termo_forma": t["termo_forma"],
            "distancia": dist,
            "lado": lado,
            "idx_termo": idx_te,
            "n_palavras": t["n_palavras"],
            "forma_em_composto": t["forma_em_composto"],
            "idx_no": i_no,
        })
    return achados


def emparelha_contexto(tokens, nos_validos, campo_compilado, near: int,
                       limites, mesma_frase: bool = True):
    """Emparelha cada termo com o nó mais próximo na mesma frase.

    Uma linha por (termo_tipo, idx_termo): o par de menor distância;
    empate → o mais à esquerda. Pares que atravessam fronteira de frase
    são descartados (contam em excluidos['fronteira_frase']).
    """
    idxs_no = indices_no(tokens, nos_validos)
    excluidos = {"fronteira_frase": 0}
    if not idxs_no:
        return [], excluidos, 0

    linhas = []
    # Uma ocorrência de termo = uma candidatura; retenção do melhor nó.
    por_chave = {}  # (termo_tipo, idx_termo) -> dict
    for t in encontra_termos(tokens, campo_compilado):
        par = melhor_par_tokens(
            tokens, idxs_no, t["idx_termo"], t["n_palavras"],
            limites, mesma_frase=mesma_frase)
        if par is None:
            excluidos["fronteira_frase"] += 1
            continue
        i_no, idx_te, dist, lado = par
        if dist > near:
            continue
        chave = (t["termo_tipo"], idx_te)
        cand = {
            "termo_tipo": t["termo_tipo"],
            "termo_forma": t["termo_forma"],
            "matched_form": t["termo_forma"],
            "canonical_term": t["termo_tipo"],
            "distancia": dist,
            "lado": lado,
            "idx_termo": idx_te,
            "idx_no": i_no,
            "n_palavras": t["n_palavras"],
            "forma_em_composto": t["forma_em_composto"],
            "n_nos_janela": len(idxs_no),
            "off_no": tokens[i_no][1],
            "off_termo": tokens[idx_te][1],
            "no": tokens[i_no][0],
        }
        ant = por_chave.get(chave)
        if ant is None or (dist, min(i_no, idx_te)) < (
                ant["distancia"], min(ant["idx_no"], ant["idx_termo"])):
            por_chave[chave] = cand

    # Colapsar várias ocorrências do mesmo tipo na janela: uma linha —
    # a do par globalmente mais próximo (empate → esquerda).
    por_tipo: dict[str, dict] = {}
    for cand in por_chave.values():
        etq = cand["termo_tipo"]
        ant = por_tipo.get(etq)
        chave_ord = (cand["distancia"], min(cand["idx_no"], cand["idx_termo"]))
        if ant is None or chave_ord < (
                ant["distancia"], min(ant["idx_no"], ant["idx_termo"])):
            por_tipo[etq] = cand

    return list(por_tipo.values()), excluidos, len(idxs_no)


def recalcular_distancia_lado(contexto: str, no_forma: str, matched_form: str,
                              *, mesma_frase: bool = True,
                              nos_validos: set[str] | None = None):
    """Recalcula distancia/lado a partir do contexto (teste de invariante)."""
    ctx = normaliza(contexto)
    toks = tokeniza(ctx)
    if nos_validos is None:
        no_l = no_forma.lower()
        nos_validos = {no_l}
        for conj in NOS.values():
            if no_l in conj:
                nos_validos = set(conj)
                break
    idxs_no = indices_no(toks, nos_validos)
    partes = matched_form.lower().split()
    L = len(partes)
    idxs_te = [j for j in range(len(toks) - L + 1)
               if [toks[k][0] for k in range(j, j + L)] == partes]
    if not idxs_no or not idxs_te:
        return None
    limites = fronteiras_frase(ctx) if mesma_frase else []
    melhor = None
    for j in idxs_te:
        par = melhor_par_tokens(toks, idxs_no, j, L, limites, mesma_frase)
        if par is None:
            continue
        i_no, idx_te, dist, lado = par
        chave = (dist, min(i_no, idx_te))
        if melhor is None or chave < melhor[0]:
            melhor = (chave, dist, lado)
    if melhor is None:
        return None
    return {"distancia": melhor[1], "lado": melhor[2]}
class _Consulta:
    def __init__(self, expr: str, etiquetas: set[str]):
        self.toks = re.findall(r"\(|\)|[^\s()]+", expr)
        self.i = 0
        self.etiquetas = etiquetas
        self.arvore = self._ou()
        if self.i < len(self.toks):
            raise ValueError(f"consulta mal formada junto de {self.toks[self.i]!r}")

    def _olha(self):
        return self.toks[self.i] if self.i < len(self.toks) else None

    def _ou(self):
        no = self._e()
        while (self._olha() or "").upper() == "OR":
            self.i += 1
            no = ("or", no, self._e())
        return no

    def _e(self):
        no = self._unario()
        while (self._olha() or "").upper() == "AND":
            self.i += 1
            no = ("and", no, self._unario())
        return no

    def _unario(self):
        t = self._olha()
        if t is None:
            raise ValueError("consulta incompleta")
        if t.upper() == "NOT":
            self.i += 1
            return ("not", self._unario())
        if t == "(":
            self.i += 1
            no = self._ou()
            if self._olha() != ")":
                raise ValueError("parêntese por fechar")
            self.i += 1
            return no
        self.i += 1
        if t not in self.etiquetas:
            raise ValueError(f"etiqueta desconhecida: {t!r}. "
                             f"Disponíveis: {sorted(self.etiquetas)}")
        return ("lit", t)

    def avalia(self, presentes: set[str]) -> bool:
        return self._av(self.arvore, presentes)

    def _av(self, no, p):
        op = no[0]
        if op == "lit":
            return no[1] in p
        if op == "not":
            return not self._av(no[1], p)
        if op == "and":
            return self._av(no[1], p) and self._av(no[2], p)
        return self._av(no[1], p) or self._av(no[2], p)


def anota_polaridade_linear(tokens, idx_termo, texto: str | None = None):
    """Graduação e modalidade por proximidade; negação sem «no.» de opus.

    Graduação: 3 tokens à esquerda. Modalidade: ±4 tokens (ex.: «texture can»).
    """
    jan_idx = range(max(0, idx_termo - 3), idx_termo)
    jan = [tokens[k][0] for k in jan_idx]
    graduado = any(w in GRADUACAO for w in jan)
    jan_mod = range(max(0, idx_termo - 4),
                    min(len(tokens), idx_termo + 5))
    modalizado = any(tokens[k][0] in MODALIDADE for k in jan_mod)

    negado = None
    for k in jan_idx:
        w = tokens[k][0]
        if w in {"no", "nos"}:
            # «no. 1» / «nos. 3-4»: não é negação
            fim = tokens[k][1] + len(w)
            if texto is not None:
                seq = texto[fim:fim + 4]
                if re.match(r"\.?\s*\d", seq):
                    continue
            continue  # sem texto, não afirmar negação por «no»
        if w in NEGACAO:
            negado = True
            break
    if negado is None:
        # só marcar False se virmos a janela e não houver negador claro
        negado = False if not any(
            tokens[k][0] in (NEGACAO - {"absence", "lack", "devoid", "lacking",
                                        "without"})
            for k in jan_idx) else True

    return negado, graduado, modalizado


def anota_sintaxe(tokens, no_idx, idx_termo, texto: str | None = None):
    """Heurística de função sintáctica (+ polaridade linear).

    Devolve (negado, graduado, modalizado, relacao).
    Mantém compatibilidade: os primeiros dois slots históricos eram
    (negado, modalizado≈graduação); o 3.º era a relação.
    """
    negado, graduado, modalizado = anota_polaridade_linear(
        tokens, idx_termo, texto)

    a, b = sorted((no_idx, idx_termo))
    entre = [tokens[k][0] for k in range(a + 1, b)]
    predicativa = any(w in COPULAS for w in entre)

    if predicativa:
        relacao = "predicativa"
    elif idx_termo == no_idx - 1 or idx_termo == no_idx + 1:
        relacao = "atributiva"
    else:
        relacao = "indeterminada"

    return negado, graduado, modalizado, relacao
