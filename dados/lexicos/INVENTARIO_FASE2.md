# Inventário de léxicos duplicados — Phase 2 discovery

Gerado antes de qualquer movimento. A aceitação «cada lista aparece
exactamente uma vez» verifica-se por diff/grep contra este inventário.

## Duplicados a eliminar (mesma lista em ≥2 sítios)

| Léxico | Locais (pré-Phase-2) | Destino canónico |
|---|---|---|
| `NOS` (paradigmas do nó) | `textura/config.py`, `textura_lexico.py` | `dados/lexicos/nos.tsv` |
| `POLO_ESTABILIDADE` | `textura/estatistica.py`, `textura_lexico.py` | `dados/lexicos/polo_estabilidade.tsv` |
| `POLO_VARIABILIDADE` | `textura/estatistica.py`, `textura_lexico.py` | `dados/lexicos/polo_variabilidade.tsv` |
| `ABREVIATURAS` | `textura/config.py`, `textura_query.py` | `dados/lexicos/abreviaturas.tsv` |
| `COPULAS` | `textura/config.py`, `textura_query.py` | `dados/lexicos/copulas.tsv` |
| `RELACOES_NUCLEARES` | `textura/config.py`, `textura_triagem.py` | `dados/lexicos/relacoes_nucleares.tsv` |
| `DOMINIO_JANELA_LEXICO` / `DOMAIN_LEXICON` | `textura/lexico.py`, `textura_concordancia_qa.py` | `dados/lexicos/dominio_janela.tsv` |
| Domínios válidos (taxonomia) | `textura_triagem.DOMINIOS_VALIDOS` (+ labels janela/QA) | `dados/lexicos/dominio_taxonomia.tsv` |
| Path→domínio | `dominios.tsv` (raiz) | `dados/lexicos/dominios_path.tsv` (+ stub/compat na raiz) |

## Listas únicas (mover para `dados/lexicos/` sem duplicar)

| Léxico | Local actual | Destino |
|---|---|---|
| `NEGACAO` | `textura/config.py` | `dados/lexicos/negacao.tsv` |
| `GRADUACAO` / `MODALIZACAO` | `textura/config.py` | `dados/lexicos/graduacao.tsv` |
| `MODALIDADE` | `textura/config.py` | `dados/lexicos/modalidade.tsv` |
| `RELACOES_NAO_NUCLEARES` | `textura_triagem.py` | `dados/lexicos/relacoes_nao_nucleares.tsv` |
| `FALSOS_AMIGOS_FORMAS` | `textura_triagem.py` | `dados/lexicos/falsos_amigos.tsv` |

## Fora de âmbito desta fase (não são word-lists partilhadas)

| Símbolo | Motivo |
|---|---|
| `CAMPO` / `SEARCH_TERMS` | Campo lexical adjudicado (padrões); fica em código/`--termos` |
| `NODE_PATTERNS` | Derivável de `NOS` / prefixo textur* |
| `NON_TEXTURAL_HEADS`, `TEXTURE` (QA regex) | Classificador de padrões QA — Phase 3 language registry |
| `LIGATURES`, `FILL_*` | Normalização/UI, não taxonomia de domínio |
| `RELACOES_VALIDAS` (`textura_analise`) | União nuclear∪não-nuclear — passar a derivar dos TSV |

## Formato TSV

- UTF-8, `#` comentários, cabeçalho na primeira linha não-comentário.
- Listas simples: coluna `token` (uma entrada por linha).
- `nos.tsv`: `lingua` \\t `forma`
- `dominio_janela.tsv`: `dominio` \\t `pista`
- `dominios_path.tsv`: `padrao` \\t `dominio` (como o actual `dominios.tsv`)
- `falsos_amigos.tsv`: `forma` \\t `motivo_exclusao`
- `dominio_taxonomia.tsv`: `dominio` \\t `fonte` (`path`\|`janela`\|`ambos`)

## Estado pós-movimento

Fonte única: `textura/lexico.py` carrega `dados/lexicos/*.tsv`.
Consumidores reexportam (sem literais). Aceitação:
`tests/test_lexicos_fonte_unica.py`.

### Precedência `dominios.tsv` (não silenciosa)

| Situação | Comportamento |
|---|---|
| Só raiz com regras | Usa raiz + `DeprecationWarning` (nomeia `dados/lexicos/dominios_path.tsv`) |
| Só canónico | Usa `dados/lexicos/dominios_path.tsv` |
| Ambos idênticos | Usa canónico + aviso para remover a raiz |
| Ambos diferem | `LexicoError` — migração consciente obrigatória |

A raiz enviada no repositório é **só comentários** (sem regras), para não
sombrear o canónico. Personalizações antigas na raiz continuam a ser
honradas ou a falhar alto se divergirem do canónico.
