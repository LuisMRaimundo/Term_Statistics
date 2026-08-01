#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Léxico e separação nó / campo lexical."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from textura.lexico import (
    NOS,
    POLO_ESTABILIDADE as _POLO_E,
    POLO_VARIABILIDADE as _POLO_V,
)

NODE_PATTERNS = {
    "textur*", "texture*", "textures*", "textural*", "texturally*",
    "textura*", "texturas*", "texturais*", "textur",
}

# Cópias mutáveis: carregar_campo_termos pode acrescentar etiquetas.
POLO_ESTABILIDADE = set(_POLO_E)
POLO_VARIABILIDADE = set(_POLO_V)


@dataclass(frozen=True)
class TermoConfig:
    canonical_term: str
    patterns: tuple[str, ...]
    polaridade: str | None  # estabilidade | variabilidade | None
    dominio: str = "campo_lexical"


# Léxico adjudicado dos dois pólos (sem controlos loud/orchestral).
SEARCH_TERMS: dict[str, TermoConfig] = {
    "uniform": TermoConfig("uniform", ("uniform*",), "estabilidade"),
    "invariable": TermoConfig("invariable", ("invariab*", "invarian*"), "estabilidade"),
    "unvarying": TermoConfig("unvarying", ("unvarying",), "estabilidade"),
    "immutable": TermoConfig("immutable", ("immutab*",), "estabilidade"),
    "unchanging": TermoConfig("unchanging", ("unchanging", "unchanged"), "estabilidade"),
    "constant": TermoConfig("constant", ("constant*",), "estabilidade"),
    "consistent": TermoConfig("consistent", ("consisten*",), "estabilidade"),
    "regular": TermoConfig("regular", ("regular*",), "estabilidade"),
    "stable": TermoConfig("stable", ("stab*",), "estabilidade"),
    "steady": TermoConfig("steady", ("stead*",), "estabilidade"),
    "sustained": TermoConfig("sustained", ("sustain*",), "estabilidade"),
    "static": TermoConfig("static", ("static", "stasis"), "estabilidade"),
    "monotonous": TermoConfig("monotonous", ("monoton*",), "estabilidade"),
    "homogeneous": TermoConfig("homogeneous", ("homogene*", "homogenous"), "estabilidade"),
    "varied": TermoConfig("varied", ("varied", "variety", "variet*"), "variabilidade"),
    "varying": TermoConfig("varying", ("varying", "variab*"), "variabilidade"),
    "changing": TermoConfig("changing", ("changing", "changeab*"), "variabilidade"),
    "irregular": TermoConfig("irregular", ("irregular*",), "variabilidade"),
    "unequal": TermoConfig("unequal", ("unequal", "uneven*"), "variabilidade"),
    "diverse": TermoConfig("diverse", ("divers*",), "variabilidade"),
    "mutable": TermoConfig("mutable", ("mutab*",), "variabilidade"),
    "multiform": TermoConfig("multiform", ("multiform*",), "variabilidade"),
    "heterogeneous": TermoConfig("heterogeneous", ("heterogene*",), "variabilidade"),
}


@dataclass
class RelatorioMapeamento:
    padroes_sem_categoria: list[str] = field(default_factory=list)
    formas_sem_categoria: list[str] = field(default_factory=list)
    padroes_sem_polaridade: list[str] = field(default_factory=list)

    def para_dataframe(self):
        import pandas as pd
        rows = (
            [{"tipo": "padrao_sem_categoria", "valor": v}
             for v in self.padroes_sem_categoria]
            + [{"tipo": "forma_sem_categoria", "valor": v}
               for v in self.formas_sem_categoria]
            + [{"tipo": "padrao_sem_polaridade", "valor": v}
               for v in self.padroes_sem_polaridade]
        )
        return pd.DataFrame(rows) if rows else pd.DataFrame(
            columns=["tipo", "valor"])


def campo_como_dict(termos: dict[str, TermoConfig] | None = None) -> dict[str, list[str]]:
    src = termos or SEARCH_TERMS
    return {t.canonical_term: list(t.patterns) for t in src.values()}


def todos_search_patterns(termos: dict[str, TermoConfig] | None = None) -> list[str]:
    src = termos or SEARCH_TERMS
    out: list[str] = []
    for t in src.values():
        out.extend(t.patterns)
    return out


# Mapa padrão -> canónico (para recuperar etiqueta a partir do padrão).
_PADRAO_PARA_CANONICO: dict[str, str] = {}


def forma_e_no(forma: str) -> bool:
    """True se a forma pertence morfologicamente ao nó textur*."""
    w = str(forma).lower().strip().strip("*")
    if not w:
        return False
    if w in set().union(*NOS.values()):
        return True
    return w.startswith("textur")


def parece_padrao_no(s: str) -> bool:
    s = str(s).strip().lower()
    if s in NODE_PATTERNS:
        return True
    return forma_e_no(s)


def stem_de_padrao(padrao: str) -> str:
    """uniform* -> uniform; *varying -> varying."""
    p = str(padrao).strip().lower()
    if p.startswith("*") and p.endswith("*"):
        return p.strip("*")
    if p.startswith("*"):
        return p[1:]
    if p.endswith("*"):
        return p[:-1]
    return p


def registar_campo(campo: dict[str, list[str]]) -> None:
    global _PADRAO_PARA_CANONICO
    _PADRAO_PARA_CANONICO = {}
    for etq, pads in campo.items():
        for p in pads:
            _PADRAO_PARA_CANONICO[p.strip().lower()] = etq


def mapa_padrao_para_canonico(padrao_ou_termos=None):
    """Sobrecarga: str → canónico | None; dict TermoConfig → mapa completo."""
    if padrao_ou_termos is None:
        if _PADRAO_PARA_CANONICO:
            return dict(_PADRAO_PARA_CANONICO)
        return {p: t.canonical_term for t in SEARCH_TERMS.values() for p in t.patterns}
    if isinstance(padrao_ou_termos, str):
        return _PADRAO_PARA_CANONICO.get(padrao_ou_termos.strip().lower())
    src = padrao_ou_termos
    if not src:
        return {}
    primeiro = next(iter(src.values()))
    if isinstance(primeiro, TermoConfig):
        return {p: t.canonical_term for t in src.values() for p in t.patterns}
    # dict[str, list[str]]
    m: dict[str, str] = {}
    for etq, pads in src.items():
        for p in pads:
            m[str(p)] = str(etq)
    return m


def termos_dos_dois_polos(
        termos: dict[str, TermoConfig] | None = None) -> dict[str, TermoConfig]:
    src = termos or SEARCH_TERMS
    return {
        k: v for k, v in src.items()
        if v.dominio == "campo_lexical"
        and v.polaridade in ("estabilidade", "variabilidade")
    }


def completar_polaridade(termos: dict[str, TermoConfig]) -> dict[str, TermoConfig]:
    out: dict[str, TermoConfig] = {}
    for etq, t in termos.items():
        pol = t.polaridade
        if pol is None and t.dominio == "campo_lexical":
            if etq in SEARCH_TERMS and SEARCH_TERMS[etq].polaridade:
                pol = SEARCH_TERMS[etq].polaridade
            else:
                pol = polaridade(etq)
        out[etq] = TermoConfig(t.canonical_term, t.patterns, pol, t.dominio)
    return out


def carregar_termos_ficheiro(caminho) -> dict[str, TermoConfig]:
    """Formato: 'etiqueta [: E|V] = padrao1, padrao2'."""
    texto = Path(caminho).read_text(encoding="utf-8")
    out: dict[str, TermoConfig] = {}
    for linha in texto.splitlines():
        linha = linha.split("#", 1)[0].strip()
        if not linha or "=" not in linha:
            continue
        esq, pads = linha.split("=", 1)
        esq = esq.strip()
        pol: str | None = None
        dominio = "campo_lexical"
        if ":" in esq:
            etq, polo = (x.strip() for x in esq.split(":", 1))
            p = polo.upper()
            if p.startswith("E"):
                pol = "estabilidade"
            elif p.startswith("V"):
                pol = "variabilidade"
            else:
                pol = None
                dominio = "controlo"
        else:
            etq = esq
        patterns = tuple(p.strip() for p in pads.split(",") if p.strip())
        out[etq] = TermoConfig(etq, patterns, pol, dominio)
    return completar_polaridade(out)


def _rx_padrao(p: str) -> re.Pattern:
    esq = p.startswith("*")
    dir_ = p.endswith("*")
    nucleo = re.escape(p.strip("*"))
    if esq and dir_:
        star = r"[\w\-]{0,20}"
        return re.compile("^" + star + nucleo + star + "$")
    if esq:
        return re.compile(r"^(?:[\w\-]*-)?" + nucleo + r"$")
    if dir_:
        return re.compile("^" + nucleo + r"\w{0,20}$")
    return re.compile("^" + nucleo + "$")


def canonical_de_forma(forma: str, campo: dict[str, list[str]]) -> str:
    """Etiqueta canónica cujo padrão casa a forma — nunca um padrão do nó."""
    f = str(forma).lower()
    for etq, pads in campo.items():
        if parece_padrao_no(etq):
            continue
        for p in pads:
            partes = p.split()
            if len(partes) == 1:
                if _rx_padrao(partes[0]).match(f):
                    return etq
            elif f == p.lower():
                return etq
    return stem_de_padrao(forma)


def carregar_campo_termos(caminho: Path,
                          campo_ref: dict[str, list[str]] | None = None
                          ) -> dict[str, list[str]]:
    """Lê 'etiqueta [: polo] = padrao1, padrao2'.

    Se a etiqueta for um padrão do nó (ex.: «textur* = uniform*»),
    reinterpreta os padrões da direita como o campo lexical.
    """
    campo: dict[str, list[str]] = {}
    avisos = []
    for linha in caminho.read_text(encoding="utf-8").splitlines():
        linha = linha.split("#", 1)[0].strip()
        if not linha or "=" not in linha:
            continue
        etq, pads_s = linha.split("=", 1)
        etq = etq.strip()
        polo = None
        if ":" in etq:
            etq, polo = (x.strip() for x in etq.split(":", 1))
        pads = [p.strip() for p in pads_s.split(",") if p.strip()]
        if not pads:
            continue

        if parece_padrao_no(etq):
            avisos.append(
                f"etiqueta '{etq}' e padrao do no — a usar os padroes "
                f"da direita como campo: {pads}")
            for p in pads:
                if parece_padrao_no(p):
                    continue
                can = stem_de_padrao(p)
                # preferir chave do léxico de referência se existir
                if campo_ref:
                    for etq_ref, prefs in campo_ref.items():
                        if any(stem_de_padrao(x) == can or x == p
                               for x in prefs):
                            can = etq_ref
                            break
                campo.setdefault(can, [])
                if p not in campo[can]:
                    campo[can].append(p)
                if polo:
                    _aplicar_polo(can, polo)
            continue

        if parece_padrao_no(etq) or forma_e_no(etq):
            raise SystemExit(
                f"Campo lexical ilegal: etiqueta '{etq}' e um padrao do no. "
                f"Use p.ex. 'uniform = uniform*' (o no vem da matriz KWIC).")

        campo[etq] = pads
        if polo:
            _aplicar_polo(etq, polo)

    for a in avisos:
        print(f"      AVISO: {a}", flush=True)

    assert_campo_sem_no(campo)
    registar_campo(campo)
    return campo


def _aplicar_polo(etq: str, polo: str) -> None:
    if polo.upper().startswith("E"):
        POLO_ESTABILIDADE.add(etq)
    elif polo.upper().startswith("V"):
        POLO_VARIABILIDADE.add(etq)


def assert_campo_sem_no(campo: dict[str, list[str]]) -> None:
    for etq in campo:
        if parece_padrao_no(etq) or forma_e_no(etq):
            raise SystemExit(
                f"Assercao falhou: canonical_term '{etq}' pertence ao no "
                f"(NODE_PATTERNS / textur*). Separe no e campo lexical.")


def assert_output_sem_no(canonical_terms) -> None:
    for can in canonical_terms:
        if can is None or (isinstance(can, float) and can != can):
            continue
        if parece_padrao_no(str(can)) or forma_e_no(str(can)):
            raise SystemExit(
                f"Assercao falhou no output: canonical_term='{can}' "
                f"e um padrao do no.")


def eixo_semantico(tipo: str) -> str:
    sinc = {"uniform", "homogeneous", "heterogeneous", "diverse",
            "varied", "unequal", "irregular", "multiform", "mutable"}
    diac = {"static", "invariable", "unvarying", "immutable", "unchanging",
            "constant", "varying", "changing", "stable", "steady", "sustained",
            "monotonous", "consistent", "regular"}
    if tipo in sinc and tipo in diac:
        return "ambos"
    if tipo in sinc:
        return "homogeneidade_sincronica"
    if tipo in diac:
        return "invariancia_diacronica"
    return "ambos"


def polaridade(tipo: str, negado=False, *, inverter_negada: bool = False):
    if tipo in POLO_ESTABILIDADE:
        base = "estabilidade"
    elif tipo in POLO_VARIABILIDADE:
        base = "variabilidade"
    else:
        return None
    # apenas negação directa (bool True ou string 'directo') inverte
    neg_dir = negado is True or negado == "directo"
    if inverter_negada and neg_dir:
        base = "variabilidade" if base == "estabilidade" else "estabilidade"
    return base


RE_EXT_FICHEIRO = re.compile(
    r"\.(pdf|txt|docx?|html?|xhtml|rtf|odt|epub|xml|tei)(?:\b|$)", re.I)


def parece_caminho_ficheiro(valor: str) -> bool:
    s = str(valor or "").strip()
    if not s or s.lower() in {"nan", "none"}:
        return False
    if RE_EXT_FICHEIRO.search(s):
        return True
    # file: URI
    if s.lower().startswith("file:"):
        return True
    return False


def doc_id_de_caminho(caminho: str) -> str:
    """SHA-256 estável do nome de ficheiro normalizado (não do contexto)."""
    import hashlib
    from pathlib import PurePosixPath
    s = str(caminho or "").strip().replace("\\", "/")
    if not s:
        return hashlib.sha256(b"").hexdigest()[:16]
    nome = PurePosixPath(s).name
    if not nome or "." not in nome:
        # directório ou caminho sem extensão: usar caminho completo
        chave = s.lower()
    else:
        chave = nome.lower()
    return hashlib.sha256(chave.encode("utf-8", errors="replace")).hexdigest()[:16]


def escolher_coluna_fonte(bruto, col_pedida: int) -> int:
    """Devolve índice 1-based da coluna com caminhos de ficheiro.

    Se a coluna pedida não tiver extensões, procura automaticamente.
    """
    ncols = bruto.shape[1]
    if 1 <= col_pedida <= ncols:
        serie = bruto.iloc[:, col_pedida - 1].astype(str)
        frac = serie.map(parece_caminho_ficheiro).mean()
        if frac >= 0.2:
            return col_pedida

    melhor, melhor_frac = None, 0.0
    for i in range(ncols):
        serie = bruto.iloc[:, i].astype(str)
        frac = float(serie.map(parece_caminho_ficheiro).mean())
        if frac > melhor_frac:
            melhor, melhor_frac = i + 1, frac
    if melhor is None or melhor_frac < 0.15:
        raise SystemExit(
            "Nenhuma coluna da matriz contem caminhos de ficheiro "
            "(extensao .pdf/.txt/...). Verifique --col-src; a coluna "
            f"pedida foi {col_pedida} e parece um directorio/raiz.")
    if melhor != col_pedida:
        print(f"      AVISO: --col-src={col_pedida} nao tem ficheiros "
              f"(frac={serie.map(parece_caminho_ficheiro).mean() if 1 <= col_pedida <= ncols else 0:.2f}); "
              f"a usar coluna {melhor} (frac={melhor_frac:.2f})", flush=True)
    return melhor
