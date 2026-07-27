# Migração — taxonomia de relações e colunas

## Colunas renomeadas / aliases

| Antiga | Nova | Notas |
|---|---|---|
| `caminho` | `caminho_ficheiro` | `caminho` mantido como alias de leitura |
| `caminho_dep` | `percurso_dep` | `caminho_dep` mantido como alias |
| `modalizado` (graduação) | `graduado` | `modalizado` passa a significar modalidade |
| `relacao` | `relacao_sintactica` | `relacao` mantido como alias |
| — | `nuclear`, `orientacao`, `governante`, `motivo_exclusao`, `fonte_classificacao`, `n_nos_janela`, `forma_em_composto`, `eixo`, `polaridade_base` | novos |

## Valores de `relacao_sintactica`

| Valor antigo (heurística / mapa) | Valor novo | Nuclear? |
|---|---|---|
| `atributiva` / `caracterizacao_atributiva` | `atributiva` | sim |
| `predicativa` / `caracterizacao_predicativa` | `predicativa` | sim |
| `especificacao_preposicional` | reavaliar com spaCy → tipicamente `nominal_genitiva` ou `incidental` | conforme árvore |
| `definicao_metalinguistica` | `predicativa_secundaria` ou `indeterminada` | conforme árvore |
| `indeterminada` | `indeterminada` | não |
| — | `nominal_composto`, `nominal_genitiva`, `adverbial` | sim |
| — | `incidental`, `adverbial_verbal`, `adverbial_de_grau`, `coordenada` | não |

**Importante:** o mapa antigo que colapsava especificação preposicional em `atributiva` foi suprimido. As categorias finas são gravadas tal qual.

## Folhas novas

- `9_Excluidas` — linhas com `nuclear=False`
- `Dominios_por_rever` — documentos sem entrada em `dominios.tsv`
- `Duplicados` — só snippets de `contexto` exactamente iguais (não caminhos/títulos)

## Reclassificar Excel antigo

```text
python utilitarios/reclassificar_concordancia.py --xlsx resultado_antigo.xlsx --saida reclassificado.xlsx
```

Requer spaCy + modelo (`run.bat`). Não volta a correr a extracção NEAR; reclassifica a folha `8_Concordancia` existente.

## Flags CLI relevantes

- `--fases {1,2,ambas}` — omissão `1` (só extracção); `ambas` ≡ `--sem-revisao`
- `--sintaxe spacy` (omissão) / `--sintaxe heuristica`
- `--incluir-nao-nucleares` — estatística sem filtro nuclear
- `--inverter-polaridade-negada` — omissão: desligada
- `--dominios dominios.tsv` / `--incluir-dominio mir_visao`

## Regressões 2.ª iteração (R1–R10)

- **R1:** `textur* = uniform*` no ficheiro de termos é reinterpretado como campo `{uniform: [uniform*]}`. Etiqueta correcta: `uniform = uniform*`.
- **R2:** `--col-src` é validado; se apontar para directorio/raiz, escolhe-se automaticamente a coluna com extensões de ficheiro. `doc_id` deriva do nome do ficheiro.
- **R3:** `motivo_exclusao=janela_sobreposta` + `n_janelas_fundidas`.
- **R6:** `negado ∈ {directo, indirecto, nao}`.
- **R10:** removidos `relacao`/`caminho_dep`/`atribuicao`/`caminho` duplicados na exportação.
