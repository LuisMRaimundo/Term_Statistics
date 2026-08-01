#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Gera ``dados/dicionario_colunas.md`` a partir das estruturas vivas do código.

Fonte:
  - ``textura.exportacao.COLUNAS_HITS_PRIORIDADE`` / ``DESCRICAO_COLUNAS_HITS``
  - ``textura.revisao.REVISAO_VOCABULARIO``

Uso (raiz do repositório)::

    python utilitarios/gera_dicionario_colunas.py
"""

from __future__ import annotations

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from textura.exportacao import (  # noqa: E402
    COLUNAS_HITS_PRIORIDADE,
    DESCRICAO_COLUNAS_HITS,
)
from textura.revisao import (  # noqa: E402
    REVISAO_EXACTAS,
    REVISAO_PREFIXOS,
)

DESTINO = RAIZ / "dados" / "dicionario_colunas.md"


def render() -> str:
    linhas = [
        "# Dicionário de colunas — `8_Concordancia`",
        "",
        "Gerado por `utilitarios/gera_dicionario_colunas.py` a partir de",
        "`textura.exportacao.COLUNAS_HITS_PRIORIDADE`. **Não editar à mão** —",
        "volte a correr o gerador após alterar a lista viva.",
        "",
        "Âmbito: colunas do pipeline NEAR (`textura_near` / `textura.pipeline`).",
        "Colunas pós-hoc `qa_*` de `textura_concordancia_qa.py` estão documentadas",
        "na docstring desse script, não aqui.",
        "",
        "## Colunas (ordem de exportação)",
        "",
        "| Coluna | Descrição |",
        "|---|---|",
    ]
    for col in COLUNAS_HITS_PRIORIDADE:
        desc = DESCRICAO_COLUNAS_HITS.get(col, "*(sem descrição)*")
        linhas.append(f"| `{col}` | {desc} |")

    linhas += [
        "",
        "## Vocabulário `revisao_sugerida`",
        "",
        "Etiquetas exactas:",
        "",
    ]
    for t in sorted(REVISAO_EXACTAS):
        linhas.append(f"- `{t}`")
    linhas += [
        "",
        "Prefixos (seguidos de detalhe após `:`):",
        "",
    ]
    for t in sorted(REVISAO_PREFIXOS):
        linhas.append(f"- `{t}*`")
    linhas += [
        "",
        "Vários valores na mesma célula separam-se por `; `.",
        "",
    ]
    return "\n".join(linhas)


def main() -> int:
    DESTINO.parent.mkdir(parents=True, exist_ok=True)
    texto = render()
    DESTINO.write_text(texto, encoding="utf-8", newline="\n")
    print(f"wrote {DESTINO} ({len(COLUNAS_HITS_PRIORIDADE)} columns)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
