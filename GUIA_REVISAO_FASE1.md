# Guia rápido — revisão do Excel (fase 1)

Este guia descreve **apenas o passo 2** do fluxo:

1. Pesquisar / extrair NEAR → gera o Excel (`*_near.xlsx`)  
2. **Rever o Excel** (este documento)  
3. Analisar o Excel revisto → estatística e gráficos  
4. Apêndice DOCX → concordância legível (fase 3)

A análise (passo 3) **só usa as linhas que deixar como nucleares** após a sua revisão. Não há desduplicação automática: isso é feito aqui, à mão.

---

## 1. Abrir o ficheiro certo

Abra o Excel da **extracção NEAR** (ex.: `COMPÒSITA_near.xlsx`, `resultado_near.xlsx`) — **não** o Excel de `Results` da pesquisa simples.

O ficheiro certo tem a folha `0_Instrucoes` e a folha `8_Concordancia`.

Folhas principais:

| Folha | Para quê |
|---|---|
| `0_Instrucoes` | Metadados da extracção (`schema_near`, `n_hits`, `n_ocorrencias`) e comando da fase 2 |
| `8_Concordancia` | **Folha de trabalho (hits NEAR)** — edite aqui |
| `8_Concordancia_Hits` | Cópia explícita do nível hits (só leitura / arquivo) |
| `8_Concordancia_Ocorrencias` | 1 linha por linha da matriz (`texture_occurrence_id`) — **não editar** |
| `Duplicados` | Grupos de passagem/janela sobreposta (informativo) |
| `Dominios_por_rever` | Ficheiros sem domínio adjudicado |
| `9_Excluidas` | Linhas já marcadas como não nucleares na extracção |
| `Config_lexico` | Tipos lexicais, polaridade e eixo automáticos |
| `Manifesto_corpus` | Lista de fontes / matriz (não editar) |

### Duas contagens defensáveis

| Contagem | Onde | Pergunta |
|---|---|---|
| **N_hits** | linhas nucleares em `8_Concordancia` | Quantos matches lexicais? |
| **N_ocorrencias** | `texture_occurrence_id` únicos (folha `8_Concordancia_Ocorrencias` / meta) | Quantas ocorrências de textura (linhas da matriz) têm a propriedade? |

Não apague hits só porque partilham passagem: `continuous` e `homogeneous` na mesma ocorrência são dois hits legítimos, uma ocorrência.

Folhas que começam por `_` (ex.: `_Fase1_snapshot`) são técnicas e estão ocultas — **não as altere**.

---

## 2. Onde trabalhar: `8_Concordancia`

Cada linha = um **hit NEAR** (termo do campo perto do nó) dentro de uma
ocorrência mestra da matriz.

Identificadores (não editar): `source_matrix_row`, `texture_occurrence_id`,
`match_id`, `hit_key`, `grupo_passagem_id`.

### Células a amarelo = editáveis

Só deve alterar as colunas a amarelo. As restantes são evidência da extracção.

| Coluna amarela | O que fazer |
|---|---|
| `relacao_sintactica` | Corrigir a relação sintáctica (lista pendente) |
| `nuclear` | `TRUE` = entra na análise; `FALSE` = fica de fora |
| `polaridade` | `estabilidade` ou `variabilidade` |
| `eixo` | `homogeneidade_sincronica`, `invariancia_diacronica` ou `ambos` |
| `negado` | `nao`, `directo` ou `indirecto` |
| `candidato_duplicado` | Já vem preenchido (passagem/janela); pode anotar |
| `dominio` | Domínio documental (texto livre / adjudicação) |
| `motivo_exclusao` | Porque excluiu ou duvidou (texto livre) |
| `revisto_por_humano` | A sua identificação / marca de revisão |
| `nota_revisao` | Comentário livre sobre a linha |

### Não alterar (evidência)

Não mexa em: `source_matrix_row`, `texture_occurrence_id`, `match_id`, `hit_key`, `canonical_term`, `matched_form`, `query_pattern`, `contexto`, `distancia`, `lado`, `doc_id`, `caminho_ficheiro`, `no`, offsets, percursos de dependência, etc.

Se a forma ou o termo canónico estiverem errados, **exclua a linha** (`nuclear=FALSE`) em vez de os reescrever.

---

## 3. Ordem de revisão recomendada

Siga esta ordem; evita retrabalho.

### Passo A — Filtrar e ler o contexto

1. Active o filtro automático na linha de cabeçalho.  
2. Leia sempre `contexto` + `matched_form` + `no` antes de decidir.  
3. Confirme que a distância (`distancia`, `lado`) faz sentido para a leitura.

### Passo B — Desduplicação manual

1. Abra a folha `Duplicados` (se existir).  
2. Em `8_Concordancia`, filtre `candidato_duplicado` ≠ (vazio).  
3. Para cada grupo repetido, **conserve uma linha** e nas restantes:
   - `nuclear` → `FALSE`
   - `motivo_exclusao` → ex.: `duplicado_manual` ou `citacao_repetida`
4. Opcional: apagar a linha inteira (efeito equivalente a excluí-la da análise).

Tipos típicos em `Duplicados` / `candidato_duplicado`:

- `contexto_repetido` / `citacao_entre_doc_ids` — mesmo texto sob ficheiros diferentes  
- `janela_sobreposta` — janelas KWIC sobrepostas no mesmo documento  
- `mesmo_doc_id_varios_caminhos` — o mesmo `doc_id` com vários caminhos  

### Passo C — Relação sintáctica e nuclearidade

1. Corrija `relacao_sintactica` com a lista pendente.  
2. Ajuste `nuclear` em conformidade:
   - relações **nucleares** → em princípio `TRUE`
   - relações **não nucleares** / incidentais → `FALSE`

**Valores nucleares** (entram na análise se `nuclear=TRUE`):

- `atributiva`
- `predicativa`
- `predicativa_secundaria`
- `nominal_composto`
- `nominal_genitiva`
- `adverbial`

**Valores não nucleares** (em geral `nuclear=FALSE`):

- `incidental`
- `adverbial_verbal`
- `adverbial_de_grau`
- `coordenada`
- `indeterminada`

Se estiver indeciso: deixe `indeterminada`, `nuclear=FALSE`, e explique em `nota_revisao`.

### Passo D — Polaridade, eixo e negação

1. `polaridade`: `estabilidade` vs `variabilidade` (conforme o tipo e o sentido no contexto).  
2. `eixo`:
   - `homogeneidade_sincronica` — contraste no mesmo momento / superfície  
   - `invariancia_diacronica` — permanência / mudança no tempo  
   - `ambos` — quando as duas leituras são legítimas  
3. `negado`:
   - `nao` — sem negação  
   - `directo` — negação explícita do termo (*not uniform*, *non-static*…)  
   - `indirecto` — negação atenuada / de escopo mais largo  

Consulte `Config_lexico` para ver a proposta automática; pode divergir dela se o contexto o exigir.

### Passo E — Domínio e exclusões temáticas

1. Folha `Dominios_por_rever`: ficheiros ainda sem domínio.  
2. Em `8_Concordancia`, preencha `dominio` nas linhas desses ficheiros (ou marque exclusão).  
3. Se a linha for ruído (OCR, metatexto, fora de domínio):  
   - `nuclear` → `FALSE`  
   - `motivo_exclusao` → ex.: `metatexto`, `ruido_ocr`, `fora_de_dominio`

### Passo F — Registo da revisão

Em **cada linha que alterar** (ou, no mínimo, em todas as que decidir manter):

1. `revisto_por_humano` → o seu nome / iniciais / data  
2. `nota_revisao` → breve justificação quando a decisão não for óbvia  

Isto alimenta a taxa de concordância automático vs humano na fase 2.

---

## 4. Valores admitidos (listas pendentes)

Nas colunas com lista, use só os valores da lista (o Excel rejeita outros).

| Coluna | Valores |
|---|---|
| `relacao_sintactica` | ver secção 3C |
| `nuclear` | `TRUE` / `FALSE` |
| `polaridade` | `estabilidade` / `variabilidade` / (vazio) |
| `eixo` | `homogeneidade_sincronica` / `invariancia_diacronica` / `ambos` / (vazio) |
| `negado` | `nao` / `directo` / `indirecto` |

---

## 5. Checklist antes de gravar

- [ ] Revistos os candidatos em `Duplicados` / `candidato_duplicado`  
- [ ] `nuclear` coerente com `relacao_sintactica`  
- [ ] Polaridade e eixo conferidos nos casos duvidosos  
- [ ] Domínios em falta tratados (`Dominios_por_rever`)  
- [ ] `revisto_por_humano` preenchido nas linhas alteradas  
- [ ] Nenhuma alteração a `canonical_term` / `matched_form` / `contexto`  
- [ ] Ficheiro gravado (mesmo nome ou cópia; a fase 2 aponta para este ficheiro)

---

## 6. Depois da revisão — fase 2

Na GUI: **«3. Analisar Excel…»** e escolha o ficheiro revisto.

Ou na linha de comando (pasta do projecto):

```text
python textura_analise.py --xlsx "CAMINHO\para\o_seu_ficheiro_revisto.xlsx"
```

A análise:

- exige a folha `0_Instrucoes` (prova de que passou pela fase 1);  
- valida os valores das colunas editáveis contra a taxonomia;  
- usa **só** as linhas com `nuclear=TRUE` tal como as deixou.

---

## 7. Erros frequentes

| Situação | O que fazer |
|---|---|
| Valor «fora da taxonomia» | Escolher da lista pendente; não inventar rótulos |
| Quer corrigir a forma casada | Excluir a linha; não editar `matched_form` |
| Duas linhas iguais | Manter uma; `nuclear=FALSE` na outra |
| Fase 2 recusa o ficheiro | Confirmar que existe `0_Instrucoes` e que gravou o Excel |
| Esqueceu-se de marcar a revisão | Preencher `revisto_por_humano` antes de analisar |

---

## 8. Resumo numa frase

**Amarelo = pode editar; `nuclear=TRUE` = entra na estatística; duplicados e erros resolvem-se na revisão, não na análise.**
