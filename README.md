# TEXTURA — Term Statistics

Corpus pipeline for NEAR co-occurrence mining around *textur\**, human
adjudication, association statistics, and APA-oriented concordance
appendices.

**Canonical local path:** `E:\PYTHON CODES\Term statistics`  
**Remote:** https://github.com/LuisMRaimundo/Term_Statistics

## Documentation

| Document | Language | Contents |
|---|---|---|
| [`MANUAL_TECNICO.md`](MANUAL_TECNICO.md) | PT | Usage, tutorial, lexicon, CLI, pitfalls |
| [`TECHNICAL_MANUAL.md`](TECHNICAL_MANUAL.md) | EN | Architecture, statistics, schema near ≥ 2 |
| [`GUIA_REVISAO_FASE1.md`](GUIA_REVISAO_FASE1.md) | PT | What to edit in `8_Concordancia` |
| [`dados/dicionario_colunas.md`](dados/dicionario_colunas.md) | PT | Concordance column dictionary |
| [`dados/lexicos/INVENTARIO_FASE2.md`](dados/lexicos/INVENTARIO_FASE2.md) | PT | Single-source TSV lexicons |

## Phase-2 frequency figures

The legacy chart `_g_freq_token.png` (one colour per canonical term) is
unchanged. Analysis also writes `_g_freq_eixos.png` / `.svg`: all realised
lexical families grouped by the author’s curation table
[`dados/lexicos/eixos_curadoria.tsv`](dados/lexicos/eixos_curadoria.tsv)
(retained / related / excluded). Caption sidecar:
`_g_freq_eixos_legenda.txt`. Appendix flow table: sheet + TSV
`curadoria_fluxo`.

```bat
python textura_analise.py --xlsx resultado_near.xlsx
python textura_analise.py --xlsx resultado_near.xlsx --modo-tese
```

`--modo-tese` (default off) removes in-image title/footer; the Word
caption lives in the sidecar. GUI: checkbox in «Analisar Excel…».

## Quick start

```bat
cd /d "E:\PYTHON CODES\Term statistics"
python textura_gui.py
```

Or CLI: search → review `*_near.xlsx` → `textura_analise.py` → optional
`textura_apendice.py`. See the Portuguese manual §5.
