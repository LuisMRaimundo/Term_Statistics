#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
textura_near.py — mineração de co-ocorrências NEAR/x sobre a matriz KWIC
=======================================================================

Opera sobre TEXTURA_TUDO_MATRIZ_v2.xlsx (folha 'Neighbor Contexts', sem
linha de cabeçalho).

DECISÃO METODOLÓGICA CENTRAL
----------------------------
As colunas de vizinhança (1-5 e 7-11) NÃO são usadas para o cálculo da
distância. Razão: contêm palavras de conteúdo com remoção de palavras
funcionais e com janelas assimétricas — o que torna a contagem de
posições não comparável com uma janela NEAR/x em palavras. Além disso,
não conservam pontuação, pelo que não permitem detectar fronteira de
frase.

Toda a análise assenta na coluna 15 (contexto integral), que conserva
pontuação e ordem. As colunas de vizinhança são conservadas apenas como
metadados.

LIMITAÇÃO A DECLARAR
--------------------
O campo de contexto está truncado: mediana de 11 tokens à esquerda e 10
à direita. Em ~12,5% dos casos há menos de 5 tokens à direita. Essas
linhas são marcadas como censuradas (campo 'censurado_dir' /
'censurado_esq'): a ausência de co-ocorrência nelas não é evidência de
ausência, e deve ser tratada como dado em falta, não como zero.

Uso:
    python textura_near.py --xlsx CAMINHO.xlsx --near 4 --lingua en
    python textura_near.py --xlsx CAMINHO.xlsx --near 4 --limite 20000
"""

from __future__ import annotations

import argparse
import math
import re
import sys
import unicodedata
from collections import Counter
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font
from scipy import stats

try:
    import textura_stats as tst_avancada
except ImportError:
    tst_avancada = None

# ---------------------------------------------------------------------------
# 1. CONFIGURAÇÃO — editar aqui
# ---------------------------------------------------------------------------

COLUNAS = ["L5", "L4", "L3", "L2", "L1", "NODE",
           "R1", "R2", "R3", "R4", "R5",
           "caminho", "url", "raiz", "contexto", "n"]

# Formas do nó por língua (col. 6). Ajustar conforme necessário.
NOS = {
    "en": {"texture", "textures", "textural", "textured", "texturally", "texturing"},
    "pt": {"textura", "texturas", "textural", "texturais"},
    "de": {"textur", "texturen"},
}

# Campo lexical. Chave = etiqueta do tipo lexical; valor = truncaturas.
# O asterisco final funciona como truncatura à direita (prefixo).
CAMPO = {
    # --- núcleo da invariância -------------------------------------------
    "uniform":      ["uniform*"],
    "invariable":   ["invariab*", "invarian*"],
    "unvarying":    ["unvarying"],
    "immutable":    ["immutab*"],
    "unchanging":   ["unchanging", "unchanged"],
    # --- sinónimos e relacionados ----------------------------------------
    "constant":     ["constant*"],
    "consistent":   ["consisten*"],
    "regular":      ["regular*"],
    "stable":       ["stab*"],
    "steady":       ["stead*"],
    "sustained":    ["sustain*"],
    "static":       ["static", "stasis"],
    "monotonous":   ["monoton*"],
    # --- campo contrastante ----------------------------------------------
    "varied":       ["varied", "variety", "variet*"],
    "varying":      ["varying", "variab*"],
    "changing":     ["changing", "changeab*"],
    "irregular":    ["irregular*"],
    "unequal":      ["unequal", "uneven*"],
    "diverse":      ["divers*"],
    "mutable":      ["mutab*"],
    "multiform":    ["multiform*"],
    # --- eixo da homogeneidade (controlo) --------------------------------
    "homogeneous":  ["homogene*", "homogenous"],
    "heterogeneous": ["heterogene*"],
    # --- controlo negativo -----------------------------------------------
    "loud":         ["loud*"],
    "orchestral":   ["orchestral"],
}

# Marcadores de negação e de modalização (para a validação sintáctica).
NEGACAO = {"not", "no", "never", "nor", "neither", "without", "lacking",
           "hardly", "scarcely", "barely", "rarely", "seldom",
           "n't", "cannot", "isn", "aren", "wasn", "weren", "doesn", "don"}

MODALIZACAO = {"less", "more", "quite", "rather", "fairly", "somewhat",
               "relatively", "certain", "some", "largely", "broadly",
               "mostly", "generally", "apparently", "seemingly", "almost",
               "nearly", "increasingly", "essentially", "virtually",
               "comparatively", "slightly", "very", "highly", "fully"}

# Cópulas: presença entre nó e termo indicia construção predicativa.
COPULAS = {"is", "are", "was", "were", "be", "been", "being",
           "remains", "remain", "remained", "becomes", "become", "became",
           "seems", "seem", "seemed", "appears", "appear", "appeared",
           "stays", "stayed"}

# Abreviaturas que terminam em ponto sem fechar frase.
ABREVIATURAS = {
    "p", "pp", "vol", "vols", "no", "nos", "ed", "eds", "cf", "ibid", "op",
    "cit", "et", "al", "e.g", "i.e", "fig", "figs", "ex", "exx", "ms", "mss",
    "mr", "mrs", "ms", "dr", "prof", "st", "ca", "c", "n", "trans", "rev",
    "repr", "diss", "univ", "publ", "chap", "chaps", "sec", "secs", "bk",
    "jan", "feb", "mar", "apr", "jun", "jul", "aug", "sep", "sept", "oct",
    "nov", "dec", "esp", "viz", "etc", "bwv", "kv", "hob", "d", "k",
}

# ---------------------------------------------------------------------------
# 2. SEGMENTAÇÃO E TOKENIZAÇÃO
# ---------------------------------------------------------------------------

RE_TOKEN = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ'’]*")
RE_FIM_FRASE = re.compile(r"[.!?;]+[\"'”’\)\]]*\s")
RE_HIFEN_QUEBRA = re.compile(r"(\w)-\s+(\w)")


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
    """Compila uma palavra, com truncatura à direita (*) e/ou à esquerda."""
    esq = p.startswith("*")
    dir_ = p.endswith("*")
    nucleo = re.escape(p.strip("*"))
    return re.compile("^" + (r"\w*" if esq else "") + nucleo +
                      (r"\w*" if dir_ else "") + "$")


def compila_campo(campo: dict[str, list[str]]):
    """Converte truncaturas em sequências de padrões de token.

    Um padrão com espaços é uma expressão de várias palavras: cada
    elemento é comparado com um token consecutivo. Exemplos válidos:
        'uniform*'            uma palavra, truncada à direita
        '*varying'            truncada à esquerda
        'not uniform'         expressão de duas palavras
        'a certain uniform*'  expressão de três, a última truncada
    """
    saida = []
    for etiqueta, padroes in campo.items():
        seqs = [[_rx_palavra(w) for w in p.split()] for p in padroes]
        seqs.sort(key=len, reverse=True)   # a mais longa tem prioridade
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


def procura_near(tokens, no_idx, campo_compilado, n, limites, mesma_frase=True):
    """Devolve todas as co-ocorrências dentro de NEAR/n do nó.

    Uma co-ocorrência é retida quando: (i) a distância em tokens é <= n;
    (ii) — se mesma_frase — nó e termo pertencem à mesma frase. No caso
    de expressões de várias palavras, exige-se que TODOS os tokens da
    expressão satisfaçam ambas as condições.
    """
    achados = []
    frase_no = indice_frase(tokens[no_idx][1], limites)
    ini, fim = max(0, no_idx - n), min(len(tokens), no_idx + n + 1)
    for j in range(ini, fim):
        if j == no_idx:
            continue
        for etiqueta, seqs in campo_compilado:
            casou = 0
            for seq in seqs:
                casou = _casa_em(tokens, j, seq)
                if casou:
                    break
            if not casou:
                continue
            span = range(j, j + casou)
            if no_idx in span:
                continue
            if max(abs(k - no_idx) for k in span) > n:
                continue
            if mesma_frase and any(
                    indice_frase(tokens[k][1], limites) != frase_no for k in span):
                continue
            achados.append({
                "termo_tipo": etiqueta,
                "termo_forma": " ".join(tokens[k][0] for k in span),
                "distancia": min(abs(k - no_idx) for k in span),
                "lado": "esq" if j < no_idx else "dir",
                "idx_termo": j,
                "n_palavras": casou,
            })
            break
    return achados


# ---------------------------------------------------------------------------
# 3b. CONSULTA BOOLEANA SOBRE OS TERMOS ENCONTRADOS NA MESMA JANELA
# ---------------------------------------------------------------------------
#
# Gramática:
#     consulta := termo ( ('AND'|'OR') termo )*
#     termo    := 'NOT'? ( '(' consulta ')' | ETIQUETA )
#
# ETIQUETA é uma chave de CAMPO. AND tem precedência sobre OR.
# Exemplos:
#     "uniform AND NOT varied"
#     "(uniform OR constant OR consistent) AND NOT homogeneous"
#     "homogeneous AND uniform"

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


def anota_sintaxe(tokens, no_idx, idx_termo):
    """Heurística de polaridade e de função sintáctica.

    - negado / modalizado: marcador nas 3 posições anteriores ao termo.
    - predicativa: cópula entre nó e termo (em qualquer ordem).
    - atributiva: termo imediatamente antes do nó, sem cópula.
    """
    jan = [tokens[k][0] for k in range(max(0, idx_termo - 3), idx_termo)]
    negado = any(w in NEGACAO for w in jan)
    modalizado = any(w in MODALIZACAO for w in jan)

    a, b = sorted((no_idx, idx_termo))
    entre = [tokens[k][0] for k in range(a + 1, b)]
    predicativa = any(w in COPULAS for w in entre)

    if predicativa:
        relacao = "predicativa"
    elif idx_termo == no_idx - 1 or idx_termo == no_idx + 1:
        relacao = "atributiva"
    else:
        relacao = "indeterminada"

    return negado, modalizado, relacao


# ---------------------------------------------------------------------------
# 4. ESTATÍSTICA DE REFERÊNCIA
# ---------------------------------------------------------------------------

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


def _token_em(doc, offset: int):
    """Token cujo intervalo de caracteres contém o offset dado."""
    for t in doc:
        if t.idx <= offset < t.idx + len(t.text):
            return t
    return None


def relacao_dependencia(doc, off_no: int, off_termo: int):
    """Classifica a relação sintáctica entre o nó e o termo co-ocorrente.

    Devolve (relação, governante, caminho).

    As categorias distinguem atribuição genuína de atribuição incidental —
    a distinção que, no levantamento manual, absorveu perto de metade das
    ocorrências. Um termo é incidental quando o seu governante sintáctico
    não é o nó nem um predicado cujo sujeito seja o nó: em 'a constant
    variety of texture', 'constant' modifica 'variety', não 'texture'.
    """
    t_no, t_te = _token_em(doc, off_no), _token_em(doc, off_termo)
    if t_no is None or t_te is None:
        return "indeterminada", "", ""

    caminho = f"{t_te.text}/{t_te.dep_}→{t_te.head.text}"

    # coordenação: recuar até ao primeiro elemento da série
    base = t_te
    while base.dep_ == "conj" and base.head is not base:
        base = base.head
    coord = " (coordenada)" if base is not t_te else ""

    # --- predicação: 'the texture is/remains uniform' ----------------------
    if base.dep_ in ("acomp", "attr", "oprd", "xcomp"):
        pred = base.head
        sujeitos = [c for c in pred.children if c.dep_ in ("nsubj", "nsubjpass")]
        if sujeitos:
            s = sujeitos[0]
            if s is t_no or t_no in list(s.subtree):
                return "predicativa" + coord, pred.text, caminho
            return "incidental" + coord, s.text, caminho

    # --- modificação directa: 'uniform texture' ---------------------------
    if base.dep_ in ("amod", "compound", "nmod", "nummod", "advmod", "appos"):
        if base.head is t_no:
            return "atributiva" + coord, t_no.text, caminho
        return "incidental" + coord, base.head.text, caminho

    # --- relação oblíqua: o nó rege o termo por preposição -----------------
    if t_no in list(base.ancestors):
        return "dependente do nó" + coord, t_no.text, caminho

    # --- caso geral: identificar o nome que o termo efectivamente qualifica
    gov = base
    while gov.head is not gov and gov.pos_ not in ("NOUN", "PROPN"):
        gov = gov.head
    if gov is t_no:
        return "dependente do nó" + coord, t_no.text, caminho
    return "incidental" + coord, gov.text, caminho


def anota_com_spacy(res: pd.DataFrame, modelo: str) -> pd.DataFrame:
    """Acrescenta relação de dependência, governante e caminho a cada linha."""
    try:
        import spacy
    except ImportError:
        print("      spaCy não instalado — mantida a heurística. "
              "Instale com: pip install spacy", file=sys.stderr)
        return res
    try:
        nlp = spacy.load(modelo, disable=["ner", "lemmatizer", "textcat"])
    except OSError:
        print(f"      modelo '{modelo}' indisponível — mantida a heurística.\n"
              f"      Instale com: python -m spacy download {modelo}",
              file=sys.stderr)
        return res

    contextos = res["contexto"].unique().tolist()
    print(f"      a analisar sintaxe de {len(contextos)} contextos únicos ...",
          flush=True)
    docs = {c: d for c, d in zip(contextos, nlp.pipe(contextos, batch_size=64))}

    rel, gov, cam = [], [], []
    for t in res.itertuples(index=False):
        r, g, c = relacao_dependencia(docs[t.contexto], t.off_no, t.off_termo)
        rel.append(r); gov.append(g); cam.append(c)
    res = res.copy()
    res["relacao_dep"] = rel
    res["governante"] = gov
    res["caminho_dep"] = cam
    res["atribuicao"] = np.where(
        res["relacao_dep"].str.startswith("incidental"), "incidental", "genuína")
    return res


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


POLO_ESTABILIDADE = {
    "uniform", "invariable", "unvarying", "immutable", "unchanging",
    "constant", "consistent", "regular", "stable", "steady", "sustained",
    "static", "monotonous", "homogeneous",
}
POLO_VARIABILIDADE = {
    "varied", "varying", "changing", "irregular", "unequal", "diverse",
    "mutable", "multiform", "heterogeneous",
}


def polaridade(tipo: str, negado: bool) -> str | None:
    if tipo in POLO_ESTABILIDADE:
        base = "estabilidade"
    elif tipo in POLO_VARIABILIDADE:
        base = "variabilidade"
    else:
        return None
    if negado:
        base = "variabilidade" if base == "estabilidade" else "estabilidade"
    return base


# ---------------------------------------------------------------------------
# 5. GRÁFICOS
# ---------------------------------------------------------------------------

def grafico_frequencias(df, destino: Path):
    cont = df.groupby("termo_tipo")["caminho"].nunique().sort_values(ascending=False)
    fig, ax = plt.subplots(figsize=(9, max(4, 0.32 * len(cont))))
    ax.barh(cont.index[::-1], cont.values[::-1], color="#4a5a6a")
    ax.set_xlabel("Obras únicas com co-ocorrência")
    ax.set_ylabel("")
    ax.set_title("Dispersão do campo lexical (obras únicas)")
    ax.grid(axis="x", linewidth=0.4, alpha=0.5)
    fig.tight_layout()
    fig.savefig(destino, dpi=150)
    plt.close(fig)


def grafico_distancias(df, destino: Path):
    fig, ax = plt.subplots(figsize=(7, 4))
    for lado, cor in (("esq", "#4a5a6a"), ("dir", "#a8763e")):
        sub = df[df["lado"] == lado]["distancia"]
        ax.hist(sub, bins=np.arange(0.5, df["distancia"].max() + 1.5),
                alpha=0.65, label=f"{lado}erda" if lado == "esq" else "direita",
                color=cor)
    ax.set_xlabel("Distância em tokens ao nó")
    ax.set_ylabel("Co-ocorrências")
    ax.set_title("Distribuição da distância por lado")
    ax.legend()
    ax.grid(axis="y", linewidth=0.4, alpha=0.5)
    fig.tight_layout()
    fig.savefig(destino, dpi=150)
    plt.close(fig)


def grafico_polaridade(df, destino: Path):
    sub = df.dropna(subset=["polaridade"])
    tab = pd.crosstab(sub["relacao"], sub["polaridade"])
    fig, ax = plt.subplots(figsize=(7, 4))
    tab.plot(kind="bar", stacked=True, ax=ax,
             color=["#4a5a6a", "#a8763e"])
    ax.set_xlabel("Relação sintáctica")
    ax.set_ylabel("Co-ocorrências")
    ax.set_title("Polaridade por relação sintáctica")
    ax.grid(axis="y", linewidth=0.4, alpha=0.5)
    fig.tight_layout()
    fig.savefig(destino, dpi=150)
    plt.close(fig)


# ---------------------------------------------------------------------------
# 6. FLUXO PRINCIPAL
# ---------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", required=True, type=Path)
    ap.add_argument("--folha", default="Neighbor Contexts")
    ap.add_argument("--near", type=int, default=4)
    ap.add_argument("--lingua", default="en", choices=list(NOS) + ["todas"])
    ap.add_argument("--limite", type=int, default=None,
                    help="processar apenas as N primeiras linhas (teste)")
    ap.add_argument("--sem-fronteira", action="store_true",
                    help="não aplicar exclusão por fronteira de frase")
    ap.add_argument("--termos", type=Path, default=None,
                    help="ficheiro de texto com o campo lexical, uma linha por "
                         "tipo: 'etiqueta = padrao1, padrao2, ...'. Substitui CAMPO.")
    ap.add_argument("--consulta", default=None,
                    help="expressão booleana sobre etiquetas, avaliada por janela. "
                         "Ex.: \"(uniform OR constant) AND NOT varied\"")
    ap.add_argument("--banda", type=int, default=12,
                    help="limite da banda de referência para as medidas de "
                         "associação (tokens); deve exceder --near")
    ap.add_argument("--col-no", type=int, default=6,
                    help="nº da coluna do nó (1 = primeira)")
    ap.add_argument("--col-ctx", type=int, default=15,
                    help="nº da coluna do contexto integral")
    ap.add_argument("--col-src", type=int, default=12,
                    help="nº da coluna da fonte/ficheiro")
    ap.add_argument("--com-cabecalho", action="store_true",
                    help="a folha tem linha de cabeçalho")
    ap.add_argument("--col-url", type=int, default=13,
                    help="nº da coluna com a hiperligação (0 = nenhuma)")
    ap.add_argument("--sintaxe", default="heuristica",
                    choices=["heuristica", "spacy"],
                    help="método de identificação da relação sintáctica")
    ap.add_argument("--modelo", default="en_core_web_sm",
                    help="modelo spaCy (en_core_web_sm, pt_core_news_sm, ...)")
    ap.add_argument("--saida", type=Path, default=Path("resultado_near.xlsx"))
    args = ap.parse_args()

    campo = dict(CAMPO)
    if args.termos:
        campo = {}
        for linha in args.termos.read_text(encoding="utf-8").splitlines():
            linha = linha.split("#", 1)[0].strip()
            if not linha or "=" not in linha:
                continue
            etq, pads = linha.split("=", 1)
            etq = etq.strip()
            # sintaxe opcional 'etiqueta : polo', com polo em {E, V, -}
            if ":" in etq:
                etq, polo = (x.strip() for x in etq.split(":", 1))
                if polo.upper().startswith("E"):
                    POLO_ESTABILIDADE.add(etq)
                elif polo.upper().startswith("V"):
                    POLO_VARIABILIDADE.add(etq)
            campo[etq] = [p.strip() for p in pads.split(",") if p.strip()]
        print(f"      campo lexical externo: {len(campo)} tipos", flush=True)

    consulta = _Consulta(args.consulta, set(campo)) if args.consulta else None

    print(f"[1/5] A ler {args.xlsx.name} ...", flush=True)
    bruto = pd.read_excel(args.xlsx, sheet_name=args.folha,
                          header=0 if args.com_cabecalho else None,
                          nrows=args.limite)
    total_bruto = len(bruto)
    ncols = bruto.shape[1]
    for rot, k in (("nó", args.col_no), ("contexto", args.col_ctx),
                   ("fonte", args.col_src)):
        if not 1 <= k <= ncols:
            print(f"Coluna de {rot} ({k}) fora do intervalo 1–{ncols}.",
                  file=sys.stderr)
            return 2
    df = pd.DataFrame({
        "NODE": bruto.iloc[:, args.col_no - 1],
        "contexto": bruto.iloc[:, args.col_ctx - 1],
        "caminho": bruto.iloc[:, args.col_src - 1],
        "url": (bruto.iloc[:, args.col_url - 1] if 1 <= args.col_url <= ncols
                else ""),
    })
    print(f"      {total_bruto} linhas, {ncols} colunas | nó=col{args.col_no} "
          f"contexto=col{args.col_ctx} fonte=col{args.col_src}", flush=True)

    # --- limpeza declarada -------------------------------------------------
    df = df.dropna(subset=["contexto"])
    nos_validos = (set().union(*NOS.values()) if args.lingua == "todas"
                   else NOS[args.lingua])
    df["NODE"] = df["NODE"].astype(str).str.lower()
    df = df[df["NODE"].isin(nos_validos)]
    print(f"      {total_bruto} linhas -> {len(df)} após filtro de língua "
          f"'{args.lingua}' e contexto não vazio", flush=True)

    campo_c = compila_campo(campo)
    registos, censura_esq, censura_dir, sem_no = 0, 0, 0, 0
    linhas, janelas = [], []
    tot_near = tot_banda = 0
    hits_near, hits_banda = Counter(), Counter()
    partes_termo: dict[str, Counter] = {e: Counter() for e in campo}
    tam_parte: Counter = Counter()

    print(f"[2/5] A extrair co-ocorrências NEAR/{args.near} ...", flush=True)
    for t in df.itertuples(index=False):
        ctx = normaliza(t.contexto)
        toks = tokeniza(ctx)
        if not toks:
            continue
        idxs = [i for i, (w, _) in enumerate(toks) if w == t.NODE]
        if not idxs:
            sem_no += 1
            continue
        limites = [] if args.sem_fronteira else fronteiras_frase(ctx)
        i = idxs[len(idxs) // 2]          # nó central em caso de repetição

        c_esq = i < args.near
        c_dir = (len(toks) - i - 1) < args.near
        censura_esq += c_esq
        censura_dir += c_dir

        achados_tudo = procura_near(toks, i, campo_c, args.banda, limites,
                                    mesma_frase=not args.sem_fronteira)
        achados = [a for a in achados_tudo if a["distancia"] <= args.near]
        distantes = [a for a in achados_tudo if a["distancia"] > args.near]

        tot_near += conta_tokens(toks, i, limites, 0, args.near,
                                 not args.sem_fronteira)
        tot_banda += conta_tokens(toks, i, limites, args.near, args.banda,
                                  not args.sem_fronteira)
        for a in distantes:
            hits_banda[a["termo_tipo"]] += 1
        for a in achados:
            hits_near[a["termo_tipo"]] += 1
            partes_termo[a["termo_tipo"]][t.caminho] += 1
        tam_parte[t.caminho] += 1

        for a in achados:
            neg, mod, rel = anota_sintaxe(toks, i, a["idx_termo"])
            linhas.append({
                "no": t.NODE,
                "termo_tipo": a["termo_tipo"],
                "termo_forma": a["termo_forma"],
                "n_palavras": a["n_palavras"],
                "distancia": a["distancia"],
                "lado": a["lado"],
                "negado": neg,
                "modalizado": mod,
                "relacao": rel,
                "polaridade": polaridade(a["termo_tipo"], neg),
                "censurado_esq": c_esq,
                "censurado_dir": c_dir,
                "off_no": toks[i][1],
                "off_termo": toks[a["idx_termo"]][1],
                "caminho": t.caminho,
                "url": t.url,
                "contexto": ctx,
            })
        if achados:
            presentes = {a["termo_tipo"] for a in achados}
            janelas.append({
                "no": t.NODE,
                "n_termos": len(presentes),
                "termos": " + ".join(sorted(presentes)),
                "conjunto": presentes,
                "consulta_satisfeita": (consulta.avalia(presentes)
                                        if consulta else None),
                "caminho": t.caminho,
                "contexto": ctx,
            })
        registos += 1

    res = pd.DataFrame(linhas)
    print(f"      {registos} linhas analisadas | {len(res)} co-ocorrências | "
          f"{sem_no} sem nó localizável", flush=True)
    if res.empty:
        print("Nenhuma co-ocorrência. Verifique o campo lexical.", file=sys.stderr)
        return 1

    if args.sintaxe == "spacy":
        print("[2b/5] Análise de dependências ...", flush=True)
        res = anota_com_spacy(res, args.modelo)
        if "atribuicao" in res.columns:
            n_inc = int((res["atribuicao"] == "incidental").sum())
            print(f"      {n_inc} de {len(res)} classificadas como incidentais "
                  f"({100*n_inc/len(res):.1f}%)", flush=True)

    # --- desduplicação por obra -------------------------------------------
    res_obra = res.drop_duplicates(subset=["caminho", "termo_tipo"])

    # --- associação, dispersão e reamostragem ------------------------------
    print("[3/5] A calcular estatística ...", flush=True)
    n_jan = len(janelas)
    filas_assoc = []
    for etq in campo:
        m = medidas_associacao(hits_near[etq], hits_banda[etq],
                               tot_near, tot_banda, n_jan)
        if not m:
            continue
        m = {"termo_tipo": etq, **m}
        m["obras"] = int(res_obra[res_obra["termo_tipo"] == etq]["caminho"].nunique())
        m["DP_Gries"] = dispersao_gries_dp(partes_termo[etq], tam_parte)
        m["D_Juilland"] = juilland_d(partes_termo[etq], len(tam_parte))
        filas_assoc.append(m)

    assoc = pd.DataFrame(filas_assoc)
    if not assoc.empty:
        assoc["p_fisher_BH"] = benjamini_hochberg(assoc["p_fisher"].values)
        assoc = assoc.sort_values("log_likelihood_G2", ascending=False)
        for c in ("p_fisher", "p_fisher_BH"):
            assoc[c] = assoc[c].map(lambda v: f"{v:.4g}")
        assoc = assoc[["termo_tipo", "O11_janela", "E11_esperado", "O12_banda_ref",
                       "obras", "log_likelihood_G2", "MI", "MI3", "t_score",
                       "z_score", "logDice", "DeltaP", "razao_possib",
                       "IC95_inf", "IC95_sup", "p_fisher", "p_fisher_BH",
                       "DP_Gries", "D_Juilland"]]


    cont = res_obra["termo_tipo"].value_counts()
    S = int((cont > 0).sum())
    resumo = pd.DataFrame({
        "indicador": [
            "Linhas KWIC analisadas",
            "Co-ocorrências brutas",
            "Co-ocorrências por obra (desduplicadas)",
            "Obras únicas com pelo menos uma co-ocorrência",
            "Obras únicas no subcorpus analisado",
            "Tipos lexicais atestados (S)",
            "Entropia de Shannon (H)",
            "Equitabilidade de Pielou (J)",
            "Inverso de Simpson (1/D)",
            "Linhas censuradas à esquerda",
            "Linhas censuradas à direita",
            f"Janela aplicada",
            "Exclusão por fronteira de frase",
            "Janelas com dois ou mais tipos lexicais",
            "Consulta booleana aplicada",
            "Janelas que satisfazem a consulta",
        ],
        "valor": [
            registos, len(res), len(res_obra),
            res_obra["caminho"].nunique(), df["caminho"].nunique(), S,
            round(shannon(cont.values), 4),
            round(pielou(cont.values), 4),
            round(simpson_inverso(cont.values), 4),
            censura_esq, censura_dir,
            f"NEAR/{args.near}",
            "não" if args.sem_fronteira else "sim",
            sum(1 for j in janelas if j["n_termos"] > 1),
            args.consulta or "—",
            (sum(1 for j in janelas if j["consulta_satisfeita"])
             if consulta else "—"),
        ],
    })

    # frequências por tipo
    freq = (res_obra.groupby("termo_tipo")
            .agg(obras=("caminho", "nunique"),
                 ocorrencias=("termo_tipo", "size"))
            .sort_values("obras", ascending=False)
            .reset_index())
    freq["proporcao"] = (freq["obras"] / freq["obras"].sum()).round(4)

    # teste binomial de polaridade + BH
    pol = res_obra.dropna(subset=["polaridade"])
    testes = []
    if len(pol):
        n_est = int((pol["polaridade"] == "estabilidade").sum())
        n_tot = len(pol)
        bt = stats.binomtest(n_est, n_tot, 0.5, alternative="two-sided")
        ic = bootstrap_proporcao((pol["polaridade"] == "estabilidade").values)
        testes.append({"teste": "Binomial — pólo de estabilidade vs 0,5",
                       "estatistica": f"{n_est}/{n_tot} = {n_est/n_tot:.3f}  "
                                      f"[IC95% bootstrap {ic[0]:.3f}–{ic[1]:.3f}]",
                       "p": bt.pvalue})
        # χ²: relação sintáctica × polaridade
        tab = pd.crosstab(pol["relacao"], pol["polaridade"])
        if tab.shape[0] > 1 and tab.shape[1] > 1:
            chi2, p, gl, _ = stats.chi2_contingency(tab)
            testes.append({"teste": "χ² — relação sintáctica × polaridade",
                           "estatistica": f"χ²({gl}) = {chi2:.3f}", "p": p})
        # χ²: lado × polaridade
        tab2 = pd.crosstab(pol["lado"], pol["polaridade"])
        if tab2.shape[0] > 1 and tab2.shape[1] > 1:
            chi2, p, gl, _ = stats.chi2_contingency(tab2)
            testes.append({"teste": "χ² — lado × polaridade",
                           "estatistica": f"χ²({gl}) = {chi2:.3f}", "p": p})
    tst = pd.DataFrame(testes)
    if not tst.empty:
        tst["p_ajustado_BH"] = benjamini_hochberg(tst["p"].values)
        tst["p"] = tst["p"].map(lambda v: f"{v:.5g}")
        tst["p_ajustado_BH"] = tst["p_ajustado_BH"].map(lambda v: f"{v:.5g}")

    # --- gráficos ----------------------------------------------------------
    print("[4/5] A gerar gráficos ...", flush=True)
    base = args.saida.parent
    g1, g2, g3 = base / "_g_freq.png", base / "_g_dist.png", base / "_g_pol.png"
    grafico_frequencias(res_obra, g1)
    grafico_distancias(res, g2)
    if len(pol):
        grafico_polaridade(res_obra, g3)

    # --- exportação --------------------------------------------------------
    print(f"[5/5] A escrever {args.saida.name} ...", flush=True)
    jan = pd.DataFrame(janelas)
    etiquetas = sorted(campo)
    mat = pd.DataFrame(0, index=etiquetas, columns=etiquetas, dtype=int)
    for conj in jan["conjunto"]:
        for a in conj:
            for b in conj:
                mat.at[a, b] += 1
    mat = mat.loc[mat.sum(axis=1) > 0, mat.sum(axis=0) > 0]

    with pd.ExcelWriter(args.saida, engine="openpyxl") as xw:
        resumo.to_excel(xw, sheet_name="1_Resumo", index=False)
        freq.to_excel(xw, sheet_name="2_Frequencias", index=False)
        if not tst.empty:
            tst.to_excel(xw, sheet_name="3_Testes", index=False)
        (pd.crosstab(res_obra["termo_tipo"], res_obra["relacao"])
         .to_excel(xw, sheet_name="4_Sintaxe"))
        mat.to_excel(xw, sheet_name="5_Coocorrencia")
        if not assoc.empty:
            assoc.to_excel(xw, sheet_name="9_Associacao", index=False)
        if "atribuicao" in res.columns:
            (pd.crosstab(res["termo_tipo"], res["relacao_dep"])
             .to_excel(xw, sheet_name="10_Dependencias"))
            (res[res["atribuicao"] == "incidental"]
             .groupby("governante").size().sort_values(ascending=False)
             .rename("ocorrencias").reset_index()
             .to_excel(xw, sheet_name="11_Governantes", index=False))
        (jan.drop(columns=["conjunto"])
            .to_excel(xw, sheet_name="6_Janelas", index=False))
        if consulta is not None:
            (jan[jan["consulta_satisfeita"]].drop(columns=["conjunto"])
                .to_excel(xw, sheet_name="7_Consulta", index=False))
        res.to_excel(xw, sheet_name="8_Concordancia", index=False)

    # --- estatística avançada ---------------------------------------------
    if tst_avancada is not None:
        print("      camada avançada: efeito, modelação, correspondências ...",
              flush=True)
        if not assoc.empty:
            assoc["log_ratio"] = [
                tst_avancada.log_ratio(int(r.O11_janela), int(r.O12_banda_ref),
                                       tot_near, tot_banda)
                for r in assoc.itertuples(index=False)]
        col_rel = "relacao_dep" if "relacao_dep" in res.columns else "relacao"
        modelo_tab, modelo_ajuste = tst_avancada.regressao_logistica(
            res, preditores=(col_rel, "lado", "distancia"))
        tab_ac = pd.crosstab(res["termo_tipo"], res[col_rel])
        g_ac, g_dend = base / "_g_ac.png", base / "_g_dend.png"
        ac_lin, ac_col, ac_prop = tst_avancada.analise_correspondencias(tab_ac, g_ac)
        perfil, _ = tst_avancada.perfis_e_dendrograma(res, g_dend)
        riqueza = tst_avancada.riqueza_lexical(res)
        if len(pol):
            n_e = int((pol["polaridade"] == "estabilidade").sum())
            riqueza["factor de Bayes (polaridade vs 0,5)"] = \
                tst_avancada.bayes_factor_proporcao(n_e, len(pol))
        riqueza_df = pd.DataFrame({"indicador": list(riqueza),
                                   "valor": list(riqueza.values())})
        with pd.ExcelWriter(args.saida, engine="openpyxl", mode="a",
                            if_sheet_exists="replace") as xw:
            riqueza_df.to_excel(xw, sheet_name="12_Riqueza", index=False)
            if not modelo_tab.empty:
                modelo_tab.to_excel(xw, sheet_name="13_Regressao", index=False)
                pd.DataFrame({"indicador": list(modelo_ajuste),
                              "valor": list(modelo_ajuste.values())}).to_excel(
                    xw, sheet_name="13_Ajuste_modelo", index=False)
            if not ac_lin.empty:
                ac_lin.to_excel(xw, sheet_name="14_AC_tipos")
                ac_col.to_excel(xw, sheet_name="14_AC_relacoes")
                pd.DataFrame({"dimensao": [f"dim{i+1}" for i in range(len(ac_prop))],
                              "inercia_pct": ac_prop}).to_excel(
                    xw, sheet_name="14_AC_inercia", index=False)
            if not perfil.empty:
                perfil.to_excel(xw, sheet_name="15_Perfis", index=False)
            if not assoc.empty:
                assoc.to_excel(xw, sheet_name="9_Associacao", index=False)

    wb = load_workbook(args.saida)
    ws = wb.create_sheet("6_Graficos")
    extras = [(base / "_g_ac.png", "A100"), (base / "_g_dend.png", "A140")]
    for k, (p, cel) in enumerate([(g1, "A1"), (g2, "A40"), (g3, "A70")] + extras):
        if p.exists():
            ws.add_image(XLImage(str(p)), cel)
    if "url" in res.columns and args.col_url:
        wsc = wb["8_Concordancia"]
        cabec = [c.value for c in wsc[1]]
        if "url" in cabec:
            i_url = cabec.index("url") + 1
            i_src = cabec.index("caminho") + 1 if "caminho" in cabec else i_url
            for lin in range(2, wsc.max_row + 1):
                alvo = wsc.cell(row=lin, column=i_url).value
                if isinstance(alvo, str) and alvo.strip():
                    cel = wsc.cell(row=lin, column=i_src)
                    cel.hyperlink = alvo.strip()
                    cel.style = "Hyperlink"

    for nome in wb.sheetnames:
        s = wb[nome]
        for c in s[1]:
            c.font = Font(name="Arial", bold=True)
            c.alignment = Alignment(vertical="center", wrap_text=True)
        s.freeze_panes = "A2"
    wb.save(args.saida)

    print("\n" + resumo.to_string(index=False))
    if not tst.empty:
        print("\n" + tst.to_string(index=False))
    print(f"\nConcluído: {args.saida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
