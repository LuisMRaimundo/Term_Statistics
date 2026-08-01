# Dicionário de colunas — `8_Concordancia`

Gerado por `utilitarios/gera_dicionario_colunas.py` a partir de
`textura.exportacao.COLUNAS_HITS_PRIORIDADE`. **Não editar à mão** —
volte a correr o gerador após alterar a lista viva.

Âmbito: colunas do pipeline NEAR (`textura_near` / `textura.pipeline`).
Colunas pós-hoc `qa_*` de `textura_concordancia_qa.py` estão documentadas
na docstring desse script, não aqui.

## Colunas (ordem de exportação)

| Coluna | Descrição |
|---|---|
| `source_matrix_row` | N.º de linha na matriz KWIC de origem |
| `texture_occurrence_id` | ID estável da ocorrência mestra de textura |
| `match_id` | ID do hit NEAR dentro da ocorrência |
| `hit_key` | Chave de deduplicação exacta do hit |
| `grupo_passagem_id` | Grupo de passagem/janela sobreposta |
| `candidato_duplicado` | Etiquetas C#### / P#### / J#### de possível duplicado |
| `no` | Forma do nó (texture / textura / …) na janela |
| `termo_tipo` | Etiqueta do tipo lexical do campo (canonical bucket) |
| `canonical_term` | Termo canónico adjudicado do campo lexical |
| `query_pattern` | Padrão (truncatura) que casou o termo |
| `termo_forma` | Forma lexical do termo na janela |
| `matched_form` | Forma exactamente casada no contexto |
| `n_palavras` | N.º de tokens do match multiword (se aplicável) |
| `distancia` | Distância em tokens entre termo e nó |
| `lado` | Lado do termo relativamente ao nó (esq/dir) |
| `negado` | Negação no escopo (nao / directo / indirecto / vazio) |
| `graduado` | Presença de graduação (more/rather/…) |
| `modalizado` | Presença de modalidade/evidencialidade |
| `relacao_sintactica` | Classe sintáctica (taxonomia nuclear/não-nuclear) |
| `polaridade_base` | Polaridade antes de inversão por negação |
| `polaridade` | Polaridade efectiva (estabilidade / variabilidade) |
| `eixo` | Eixo semântico (homogeneidade_sincronica / …) |
| `censurado_esq` | Contexto truncado à esquerda na matriz |
| `censurado_dir` | Contexto truncado à direita na matriz |
| `idx_no` | Índice token do nó na janela |
| `idx_termo` | Índice token do termo na janela |
| `off_no` | Offset de carácter do nó no contexto |
| `off_termo` | Offset de carácter do termo no contexto |
| `n_nos_janela` | N.º de formas do nó na janela |
| `forma_em_composto` | Match dentro de composto hifenizado |
| `caminho_ficheiro` | Caminho/fonte documental (valor de dados, não path OS) |
| `doc_id` | Identificador estável do documento |
| `url` | URL se presente na matriz |
| `contexto` | Janela textual exportada (evidência) |
| `motivo_exclusao` | Motivo de exclusão / não-nuclear |
| `nuclear` | TRUE = entra na análise pós-revisão |
| `fonte_classificacao` | dependencias (spaCy) ou heuristica |
| `n_janelas_fundidas` | Janelas fundidas por sobreposição |
| `revisao_sugerida` | Etiquetas de revisão automática (ver vocabulário) |
| `nucleo_da_propriedade` | Núcleo nominal da propriedade na árvore |
| `orientacao` | Direcção da relação (termo_sobre_no / …) |
| `governante` | Governante sintáctico reportado |
| `percurso_dep` | Percurso de dependência spaCy |
| `dominio` | Domínio documental (triagem path / revisão) |
| `dominio_janela` | Domínio sugerido por pistas na janela |
| `revisto_por_humano` | Marca de revisão humana |
| `nota_revisao` | Nota livre do revisor |

## Vocabulário `revisao_sugerida`

Etiquetas exactas:

- `atributiva_coordenada`
- `atributiva_via_conj`
- `genitiva_por_complemento`

Prefixos (seguidos de detalhe após `:`):

- `associativa_com_nao_textural:*`
- `coordenacao_heterogenea:*`
- `dominio_janela:*`

Vários valores na mesma célula separam-se por `; `.
