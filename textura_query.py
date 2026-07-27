#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
textura_query.py — motor de consulta booleana sobre contexto KWIC
================================================================

Gramática (precedência crescente):
    expr     := or_expr
    or_expr  := and_expr ( (OR | NOR) and_expr )*
    and_expr := near_expr ( AND near_expr )*
    near_expr:= not_expr ( NEAR/n not_expr )*
    not_expr := NOT not_expr | primary
    primary  := '(' expr ')' | termo | "frase"

Operadores: AND, OR, NOR, NOT, NEAR/n
Wildcards:  * (qualquer sequência)   ? (um carácter)
Frases:     "not uniform"  ou  not uniform  (tokens consecutivos)

NOR é o NOR booleano: A NOR B ≡ ¬(A ∨ B).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


RE_TOKEN = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ'’]*")
RE_HIFEN = re.compile(r"(\w)-\s+(\w)")
RE_FIM_FRASE = re.compile(r"[.!?;]+[\"'”’\)\]]*\s")
RE_NEAR = re.compile(r"^NEAR/(\d+)$", re.I)
OPS = {"AND", "OR", "NOR", "NOT"}
# Operadores colados a wildcards: unchangin*OR → unchangin* + OR
# (só após '*', para não partir palavras como COLOR / STANDARD)
_RE_OP_COLADO_FIM = re.compile(
    r"^(?P<pre>.+\*)(?P<op>NOR|AND|NOT|OR)$", re.I)
_RE_OP_COLADO_INI = re.compile(
    r"^(?P<op>NOR|AND|NOT|OR)(?P<pos>.+\*?)$", re.I)

ABREVIATURAS = {
    "p", "pp", "vol", "vols", "no", "nos", "ed", "eds", "cf", "ibid", "op",
    "cit", "et", "al", "e.g", "i.e", "fig", "figs", "ex", "exx", "ms", "mss",
    "mr", "mrs", "ms", "dr", "prof", "st", "ca", "c", "n", "trans", "rev",
    "repr", "diss", "univ", "publ", "chap", "chaps", "sec", "secs", "bk",
    "jan", "feb", "mar", "apr", "jun", "jul", "aug", "sep", "sept", "oct",
    "nov", "dec", "esp", "viz", "etc", "bwv", "kv", "hob", "d", "k",
}

COPULAS = {
    "is", "are", "was", "were", "be", "been", "being",
    "remains", "remain", "remained", "becomes", "become", "became",
    "seems", "seem", "seemed", "appears", "appear", "appeared",
    "stays", "stayed",
}
# Só preposições que tipicamente especificam / caracterizam (não locativas vagas).
PREPS_ESP = {"of", "with", "without", "as"}
DETS = {
    "the", "a", "an", "this", "that", "these", "those", "its", "his", "her",
    "their", "our", "any", "some", "no", "each", "every", "both", "such",
}
ADVS_GRAU = {
    "very", "quite", "rather", "more", "most", "less", "least", "so", "too",
    "highly", "fairly", "somewhat", "increasingly", "relatively", "extremely",
    "slightly", "almost", "nearly", "entirely", "wholly", "purely", "truly",
    "especially", "particularly", "remarkably",
}
NEG = {"not", "never", "n't"}
VERBOS_CARACT = {
    "called", "termed", "named", "described", "defined", "characterized",
    "characterised", "regarded", "considered", "known", "labelled", "labeled",
}
FUNCIONAIS_LEVES = DETS | ADVS_GRAU | NEG | {"as", "by", "to"}


def normaliza(texto: str) -> str:
    texto = str(texto).replace("\n", " ").replace("\r", " ")
    texto = RE_HIFEN.sub(r"\1\2", texto)
    return re.sub(r"\s+", " ", texto).strip()


def tokeniza(texto: str) -> list[tuple[str, int, int]]:
    """Lista de (token minúsculo, início, fim) em caracteres."""
    return [(m.group(0).lower(), m.start(), m.end())
            for m in RE_TOKEN.finditer(texto)]


def fronteiras_frase(texto: str) -> list[int]:
    """Offsets de carácter onde termina uma frase (. ! ? ;)."""
    limites = []
    for m in RE_FIM_FRASE.finditer(texto + " "):
        i = m.start()
        anterior = texto[max(0, i - 30):i]
        palavra = re.search(r"([A-Za-zÀ-ÿ.]+)$", anterior)
        if palavra:
            cand = palavra.group(1).lower().rstrip(".")
            if cand in ABREVIATURAS:
                continue
        if (i > 0 and texto[i - 1].isdigit()
                and i + 1 < len(texto) and texto[i + 1:i + 2].isdigit()):
            continue
        if texto[i:i + 3] == "...":
            continue
        limites.append(i)
    return limites


def indice_frase(pos: int, limites: list[int]) -> int:
    return sum(1 for L in limites if L < pos)


def mesma_frase_spans(tokens, span_a, span_b, limites) -> bool:
    """True se todos os tokens dos dois spans estão na mesma frase."""
    if not limites:
        return True
    ids = set()
    for a, b in (span_a, span_b):
        for k in range(a, b):
            if 0 <= k < len(tokens):
                ids.add(indice_frase(tokens[k][1], limites))
    return len(ids) == 1


def relacao_sintactica(tokens, i: int, j: int) -> str | None:
    """Relação em que um termo define / caracteriza o outro (ou se ligam
    por predicação ou especificação).

    Aceita apenas:
      - caracterização atributiva: «uniform texture», «very dense texture»
      - caracterização predicativa: «the texture is uniform»
      - especificação preposicional: «texture of polyphony», «melody with texture»
      - definição metalinguística: «texture described as uniform»

    Rejeita co-ocorrência incidental na mesma janela sem vínculo gramatical
    de definição/caracterização.
    """
    if i == j:
        return "identico"
    a, b = sorted((i, j))
    dist = b - a
    entre = [tokens[k][0] for k in range(a + 1, b)]

    # 1. Modificação directa (adjacentes)
    if dist == 1:
        return "caracterizacao_atributiva"

    # 2. Cadeia curta só com determinantes / grau / negação
    #    ex.: «a uniform texture», «very dense textures»
    if dist <= 3 and entre and all(w in (DETS | ADVS_GRAU | NEG) for w in entre):
        return "caracterizacao_atributiva"

    # 3. Predicação: «texture is / remains / becomes uniform»
    if any(w in COPULAS for w in entre):
        extras = [w for w in entre
                  if w not in COPULAS and w not in FUNCIONAIS_LEVES]
        if len(extras) <= 1:
            return "caracterizacao_predicativa"

    # 4. Especificação: «texture of X», «X with texture», «texture as Y»
    if dist <= 3:
        preps = [w for w in entre if w in PREPS_ESP]
        if len(preps) == 1:
            resto = [w for w in entre
                     if w not in PREPS_ESP and w not in DETS and w not in ADVS_GRAU]
            if not resto:
                return "especificacao_preposicional"

    # 5. Definição / caracterização explícita
    if any(w in VERBOS_CARACT for w in entre):
        resto = [w for w in entre
                 if w not in VERBOS_CARACT and w not in FUNCIONAIS_LEVES
                 and w not in COPULAS]
        if len(resto) <= 1:
            return "definicao_metalinguistica"

    return None


def _relacao_spans(tokens, span_a, span_b) -> str | None:
    """Tenta relação entre extremos dos spans (primeiro/último token)."""
    a0, a1 = span_a
    b0, b1 = span_b
    for i in (a0, a1 - 1):
        for j in (b0, b1 - 1):
            if i < 0 or j < 0:
                continue
            rel = relacao_sintactica(tokens, i, j)
            if rel:
                return rel
    return None


def pares_validos(tokens, spans_a, spans_b, limites, mesma_frase: bool,
                  exigir_sintaxe: bool):
    """Gera pares (span_a, span_b, dist, rel) que passam os filtros."""
    for a0, a1 in spans_a:
        for b0, b1 in spans_b:
            if mesma_frase and not mesma_frase_spans(
                    tokens, (a0, a1), (b0, b1), limites):
                continue
            rel = _relacao_spans(tokens, (a0, a1), (b0, b1))
            if exigir_sintaxe and rel is None:
                continue
            if a1 - 1 < b0:
                d = b0 - (a1 - 1)
            elif b1 - 1 < a0:
                d = a0 - (b1 - 1)
            else:
                d = 0
            yield (a0, a1), (b0, b1), d, rel or ""


# Limite da expansão de '*' — evita casar tokens PDF colados (textureand, …)
_MAX_STAR = 14

# Sufixos típicos de colagem OCR/PDF (só rejeitados se o token for longo)
_RE_COLADO = re.compile(
    r"(and|the|of|for|with|from|into|onto|kinds|spaces|introduced|"
    r"determining|morphisms|bicontinuous|discrete)$",
    re.I,
)


def _wildcard_para_regex(padrao: str) -> re.Pattern:
    """Converte um padrão com * e ? num regex ancorado à palavra."""
    partes = []
    for ch in padrao:
        if ch == "*":
            partes.append(rf"\w{{0,{_MAX_STAR}}}")
        elif ch == "?":
            partes.append(r"\w")
        else:
            partes.append(re.escape(ch))
    return re.compile("^" + "".join(partes) + "$", re.I)


def token_parece_colado(forma: str) -> bool:
    """Heurística: token longo que termina em cola tipográfica PDF."""
    f = (forma or "").strip()
    if len(f) > 22:
        return True
    # textureand (10), variouskinds (12), …
    if len(f) >= 10 and _RE_COLADO.search(f):
        return True
    return False


def forma_casa_padrao(forma: str, padrao: str) -> bool:
    """True se a forma (1+ tokens) casa o padrão adjudicado."""
    if token_parece_colado(forma.replace(" ", "")):
        return False
    palavras_p = padrao.split()
    palavras_f = forma.split()
    if len(palavras_p) != len(palavras_f):
        return False
    return all(
        _wildcard_para_regex(p).match(f)
        for p, f in zip(palavras_p, palavras_f)
    )


def forma_no_lexico(forma: str, padroes: list[str]) -> bool:
    """A forma pertence ao léxico adjudicado (casa algum padrão)?"""
    return any(forma_casa_padrao(forma, p) for p in padroes)


def padroes_da_arvore(no) -> list[str]:
    """Folhas (padrões) da árvore de consulta — léxico adjudicado da pesquisa."""
    op = no[0]
    if op == "term":
        return [no[1]]
    if op == "not":
        return padroes_da_arvore(no[1])
    if op in ("and", "or", "nor"):
        return padroes_da_arvore(no[1]) + padroes_da_arvore(no[2])
    if op == "near":
        return padroes_da_arvore(no[2]) + padroes_da_arvore(no[3])
    return []


def near_da_arvore(no) -> int | None:
    """Primeiro NEAR/n encontrado na consulta, se existir."""
    op = no[0]
    if op == "near":
        return int(no[1])
    if op == "not":
        return near_da_arvore(no[1])
    if op in ("and", "or", "nor"):
        return near_da_arvore(no[1]) or near_da_arvore(no[2])
    return None


def campo_desde_padroes(padroes: list[str]) -> dict[str, list[str]]:
    """Campo lexical estilo textura_near: etiqueta → [padrão]."""
    campo: dict[str, list[str]] = {}
    for p in padroes:
        etq = re.sub(r"[^\w]+", "_", p.strip("*?")).strip("_").lower() or "termo"
        base, n = etq, 2
        while etq in campo:
            etq = f"{base}_{n}"
            n += 1
        campo[etq] = [p]
    return campo


def e_padrao_no(padrao: str, node_patterns: list[str] | None = None) -> bool:
    """True se o padrão da consulta é (ou coincide com) um node_pattern."""
    try:
        import textura_lexico as tlex
        nodes = node_patterns or list(tlex.NODE_PATTERNS)
    except Exception:  # noqa: BLE001
        nodes = node_patterns or ["textur*"]
    p = (padrao or "").strip().lower()
    for n in nodes:
        nl = n.lower()
        if p == nl:
            return True
        # textur* ≡ textur / texture* curto
        if nl.endswith("*") and p.rstrip("*") == nl.rstrip("*"):
            return True
        if p.endswith("*") and p.rstrip("*") == nl.rstrip("*"):
            return True
    return False


def separar_node_e_campo(padroes: list[str],
                         node_patterns: list[str] | None = None
                         ) -> tuple[list[str], list[str]]:
    """Separa padrões do nó vs padrões do campo lexical (req. 1–2)."""
    nos, campo = [], []
    for p in padroes:
        (nos if e_padrao_no(p, node_patterns) else campo).append(p)
    return list(dict.fromkeys(nos)), list(dict.fromkeys(campo))


def forma_e_no(forma: str, node_patterns: list[str] | None = None) -> bool:
    """A matched_form pertence ao nó (não deve entrar no campo)?"""
    try:
        import textura_lexico as tlex
        nodes = node_patterns or list(tlex.NODE_PATTERNS)
    except Exception:  # noqa: BLE001
        nodes = node_patterns or ["textur*"]
    return forma_no_lexico(forma, nodes)


@dataclass
class Hit:
    ok: bool
    spans: list[tuple[int, int]] = field(default_factory=list)   # índices de token
    formas: list[str] = field(default_factory=list)
    distancias_near: list[int] = field(default_factory=list)
    relacoes: list[str] = field(default_factory=list)

    @staticmethod
    def falso() -> "Hit":
        return Hit(False)

    @staticmethod
    def vazio_ok() -> "Hit":
        """Verdadeiro sem spans (ex.: NOT sem âncora positiva)."""
        return Hit(True)


def _junta(a: Hit, b: Hit, ok: bool) -> Hit:
    if not ok:
        return Hit.falso()
    return Hit(True, a.spans + b.spans, a.formas + b.formas,
               a.distancias_near + b.distancias_near,
               a.relacoes + b.relacoes)


class ConsultaBooleana:
    """Parser + avaliador de consultas booleanas sobre tokens."""

    def __init__(self, expr: str, mesma_frase: bool = True,
                 exigir_sintaxe: bool = True):
        expr = (expr or "").strip()
        if not expr:
            raise ValueError("consulta vazia")
        self.fonte = expr
        self.mesma_frase = mesma_frase
        self.exigir_sintaxe = exigir_sintaxe
        self.toks = self._lex(expr)
        self.i = 0
        self.arvore = self._or()
        if self.i < len(self.toks):
            raise ValueError(f"consulta mal formada junto de {self.toks[self.i]!r}")
        # léxico adjudicado = padrões efectivamente escritos na consulta
        self.padroes = list(dict.fromkeys(padroes_da_arvore(self.arvore)))
        self.near_n = near_da_arvore(self.arvore)

    @staticmethod
    def _partir_ops_colados(tok: str) -> list[str]:
        """Separa OR/AND/… colados a um termo com '*' (ex.: unchangin*OR)."""
        if not tok or tok.upper() in OPS or RE_NEAR.match(tok):
            return [tok]
        m = _RE_OP_COLADO_FIM.match(tok)
        if m:
            return [m.group("pre"), m.group("op").upper()]
        m = _RE_OP_COLADO_INI.match(tok)
        if m and m.group("pos").upper() not in OPS:
            return [m.group("op").upper(), m.group("pos")]
        return [tok]

    @staticmethod
    def _lex(expr: str) -> list[str]:
        out, i, n = [], 0, len(expr)
        while i < n:
            c = expr[i]
            if c.isspace():
                i += 1
                continue
            if c in "()":
                out.append(c); i += 1; continue
            if c in "\"'":
                q, j = c, i + 1
                while j < n and expr[j] != q:
                    j += 1
                if j >= n:
                    raise ValueError("aspas por fechar")
                out.append(expr[i + 1:j]); i = j + 1; continue
            j = i
            while j < n and (not expr[j].isspace()) and expr[j] not in "()\"'":
                j += 1
            out.extend(ConsultaBooleana._partir_ops_colados(expr[i:j]))
            i = j
        return out

    def _olha(self):
        return self.toks[self.i] if self.i < len(self.toks) else None

    def _or(self):
        no = self._and()
        while True:
            t = (self._olha() or "")
            up = t.upper()
            if up == "OR":
                self.i += 1
                no = ("or", no, self._and())
            elif up == "NOR":
                self.i += 1
                no = ("nor", no, self._and())
            else:
                break
        return no

    def _and(self):
        no = self._near()
        while (self._olha() or "").upper() == "AND":
            self.i += 1
            no = ("and", no, self._near())
        return no

    def _near(self):
        no = self._not()
        while True:
            t = self._olha()
            if t is None:
                break
            m = RE_NEAR.match(t)
            if not m:
                break
            self.i += 1
            no = ("near", int(m.group(1)), no, self._not())
        return no

    def _not(self):
        t = self._olha()
        if t is None:
            raise ValueError("consulta incompleta")
        if t.upper() == "NOT":
            self.i += 1
            return ("not", self._not())
        return self._primary()

    def _primary(self):
        t = self._olha()
        if t is None:
            raise ValueError("consulta incompleta")
        if t == "(":
            self.i += 1
            no = self._or()
            if self._olha() != ")":
                t2 = self._olha()
                dica = ""
                if t2 and "*" in (t2 or "") and any(
                        t2.upper().endswith(op) for op in OPS):
                    dica = " (verifique espaços antes de OR/AND)"
                raise ValueError(
                    f"parêntese por fechar; encontrado {t2!r}{dica}")
            self.i += 1
            return no
        if t.upper() in OPS or RE_NEAR.match(t):
            raise ValueError(f"operador inesperado: {t!r}")
        self.i += 1
        return ("term", t)

    # ------------------------------------------------------------------ eval
    def avalia(self, tokens: list[tuple], texto: str = "") -> Hit:
        """tokens: (forma, início, fim). texto usado para fronteiras de frase."""
        formas = [t[0] for t in tokens]
        limites = fronteiras_frase(texto) if self.mesma_frase else []
        return self._av(self.arvore, formas, tokens, limites)

    def _casamentos(self, formas: list[str], padrao: str) -> list[tuple[int, int, str]]:
        """Devolve (início, fim_excl, forma_juntada) para cada casamento."""
        palavras = padrao.split()
        if not palavras:
            return []
        rxs = [_wildcard_para_regex(p) for p in palavras]
        L, n = len(rxs), len(formas)
        hits = []
        for j in range(n - L + 1):
            if all(rxs[k].match(formas[j + k]) for k in range(L)):
                forma = " ".join(formas[j:j + L])
                if token_parece_colado(forma.replace(" ", "")):
                    continue
                hits.append((j, j + L, forma))
        return hits

    def filtra_formas(self, formas: list[str]) -> list[str]:
        """Mantém só formas do léxico adjudicado da consulta."""
        out = []
        for f in formas:
            if f and forma_no_lexico(f, self.padroes) and not token_parece_colado(
                    f.replace(" ", "")):
                out.append(f)
        return list(dict.fromkeys(out))

    def _filtra_pares(self, a: Hit, b: Hit, tokens, limites, max_dist=None):
        """Aplica filtros de frase/sintaxe (e NEAR, se max_dist)."""
        melhores, dist_ok, formas, rels = [], [], [], []
        for sa, sb, d, rel in pares_validos(
                tokens, a.spans, b.spans, limites,
                self.mesma_frase, self.exigir_sintaxe):
            if max_dist is not None and d > max_dist:
                continue
            melhores.extend([sa, sb])
            dist_ok.append(d)
            formas.append(" ".join(tokens[k][0] for k in range(*sa)))
            formas.append(" ".join(tokens[k][0] for k in range(*sb)))
            if rel:
                rels.append(rel)
        if not melhores:
            return Hit.falso()
        visto, spans = set(), []
        for s in melhores:
            if s not in visto:
                visto.add(s)
                spans.append(s)
        return Hit(True, spans=spans,
                   formas=list(dict.fromkeys(formas)),
                   distancias_near=dist_ok if max_dist is not None else [],
                   relacoes=list(dict.fromkeys(rels)))

    def _av(self, no, formas: list[str], tokens, limites) -> Hit:
        op = no[0]
        if op == "term":
            cas = self._casamentos(formas, no[1])
            if not cas:
                return Hit.falso()
            return Hit(True,
                       spans=[(a, b) for a, b, _ in cas],
                       formas=[f for _, _, f in cas])

        if op == "not":
            h = self._av(no[1], formas, tokens, limites)
            return Hit.falso() if h.ok else Hit.vazio_ok()

        if op == "and":
            a = self._av(no[1], formas, tokens, limites)
            b = self._av(no[2], formas, tokens, limites)
            if not (a.ok and b.ok):
                return Hit.falso()
            # um lado sem spans (ex. NOT): basta a conjunção lógica
            if not a.spans or not b.spans:
                return _junta(a, b, True)
            # ambos com âncoras: exigir mesma frase + relação sintáctica
            if self.mesma_frase or self.exigir_sintaxe:
                return self._filtra_pares(a, b, tokens, limites)
            return _junta(a, b, True)

        if op == "or":
            a = self._av(no[1], formas, tokens, limites)
            b = self._av(no[2], formas, tokens, limites)
            if a.ok and b.ok:
                return _junta(a, b, True)
            if a.ok:
                return a
            if b.ok:
                return b
            return Hit.falso()

        if op == "nor":
            a = self._av(no[1], formas, tokens, limites)
            b = self._av(no[2], formas, tokens, limites)
            return Hit.vazio_ok() if not (a.ok or b.ok) else Hit.falso()

        if op == "near":
            n, esq, dir_ = no[1], no[2], no[3]
            a = self._av(esq, formas, tokens, limites)
            b = self._av(dir_, formas, tokens, limites)
            if not (a.ok and b.ok) or not a.spans or not b.spans:
                return Hit.falso()
            return self._filtra_pares(a, b, tokens, limites, max_dist=n)

        raise ValueError(f"nó desconhecido: {op}")


def snippet_destacado(texto: str, tokens: list[tuple], spans: list[tuple[int, int]],
                      raio: int = 12) -> str:
    """Recorta o contexto em torno dos spans e marca os termos com «...»."""
    if not tokens:
        return texto[:240]
    if not spans:
        return texto[:240]

    ini_tok = max(0, min(s[0] for s in spans) - raio)
    fim_tok = min(len(tokens), max(s[1] for s in spans) + raio)
    marcar = set()
    for a, b in spans:
        marcar.update(range(a, b))

    pecas = []
    k = ini_tok
    while k < fim_tok:
        if k in marcar:
            j = k
            while j < fim_tok and j in marcar:
                j += 1
            pecas.append("«" + " ".join(tokens[t][0] for t in range(k, j)) + "»")
            k = j
        else:
            pecas.append(tokens[k][0])
            k += 1
    prefixo = "…" if ini_tok > 0 else ""
    sufixo = "…" if fim_tok < len(tokens) else ""
    return prefixo + " ".join(pecas) + sufixo


def nome_documento(caminho: str) -> str:
    if not caminho:
        return ""
    s = str(caminho).replace("\\", "/").rstrip("/")
    return s.rsplit("/", 1)[-1]


def hiperligacao(url, caminho) -> str:
    """Prefere a URL da matriz; senão constrói file:/// a partir do caminho."""
    if isinstance(url, str) and url.strip():
        return url.strip()
    if isinstance(caminho, str) and caminho.strip():
        p = caminho.strip().replace("\\", "/")
        if p.startswith("file:"):
            return p
        return "file:///" + p.lstrip("/")
    return ""
