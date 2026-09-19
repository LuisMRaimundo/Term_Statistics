# TEXTURA — Manual técnico

Inventário lexicográfico–corpus de *texture* / *textura* e do campo lexical
adjudicado (homogeneidade, estase, imobilidade, …).

Este documento é o manual de uso em português: explanação simples, explanação
para especialistas, tutorial e exemplos. O protocolo formal em inglês
(estatística, schema, bibliografia de métodos) continua em
[`TECHNICAL_MANUAL.md`](TECHNICAL_MANUAL.md). A revisão humana do Excel está em
[`GUIA_REVISAO_FASE1.md`](GUIA_REVISAO_FASE1.md).

**Código canónico:** `E:\PYTHON CODES\Term statistics`  
**Schema:** near ≥ 2  
**Data deste texto:** 2026-09-19

---

## Índice

1. [Explanação 1 — simples, directa, objectiva](#1-explanação-1--simples-directa-objectiva)
2. [Explanação 2 — técnica para especialistas](#2-explanação-2--técnica-para-especialistas)
3. [O que o programa faz e o que não faz](#3-o-que-o-programa-faz-e-o-que-não-faz)
4. [Peças do sistema](#4-peças-do-sistema)
5. [Tutorial](#5-tutorial)
6. [Exemplos de aplicação](#6-exemplos-de-aplicação)
7. [Como escrever o léxico](#7-como-escrever-o-léxico)
8. [Línguas, acentos e nós](#8-línguas-acentos-e-nós)
9. [Folhas do Excel](#9-folhas-do-excel)
10. [Contagens que se podem defender](#10-contagens-que-se-podem-defender)
11. [Referência CLI](#11-referência-cli)
12. [Armadilhas frequentes](#12-armadilhas-frequentes)
13. [Glossário](#13-glossário)

---

## 1. Explanação 1 — simples, directa, objectiva

Tem uma grande folha Excel com frases em que aparece a palavra *texture*
(ou *textura*). Isso é a **matriz**.

Escreve uma lista de palavras do campo (por exemplo *homogeneous*,
*estática*, *stasis*). O programa procura essas palavras **perto** de
*texture* na mesma frase e, de preferência, quando uma caracteriza a outra
(*homogeneous texture*, *textura estatica*).

O resultado é outra folha Excel. Cada linha é um **encontro**. O programa
marca sozinho o que parece uso real (**nuclear**) e o que parece
acidental. Você abre o Excel, corrige o que estiver errado, e só depois
pede as contas e os gráficos.

Três regras de ouro:

1. **Não analise o Excel da pesquisa simples** (`resultado_pesquisa.xlsx`).
   Analise o da extracção (`*_near.xlsx`) ou o da análise (`*_analise.xlsx`).
2. **Só edite as células amarelas** em `8_Concordancia`. Não reescreva a
   palavra encontrada: se estiver errada, marque `nuclear = FALSE`.
3. **O padrão tem de parecer-se com o que está na matriz.** Se a matriz
   não tem acentos, `estátic*` mesmo assim encontra `estatica`. Mas
   `invari` sem asterisco só encontra a palavra exacta `invari`, não
   *invariant*. `estátic*` não é o inglês *static*.

Fluxo em uma linha:

`matriz → pesquisar → rever Excel → analisar → (opcional) apêndice Word`

---

## 2. Explanação 2 — técnica para especialistas

### 2.1 Objecto e ontologia a três níveis

O objecto não é o PDF original. É uma matriz KWIC (`Neighbor Contexts`,
sem cabeçalho por omissão) em que cada linha é uma **ocorrência mestra de
textura**. A distância NEAR mede-se no **contexto integral** (coluna 15),
não nas colunas de vizinhança 1–5 / 7–11.

| Nível | Identificador | Pergunta estatística |
|---|---|---|
| Ocorrência | `texture_occurrence_id` = `{doc_id}::ROW_{source_matrix_row}` | Em quantas ocorrências de textura se observa a propriedade *P*? |
| Hit | `hit_key` / `match_id` (`M001`…) | Quantos casamentos lexicais? |
| Janela | `contexto` | Só exibição; não é observação independente |

Dois hits na mesma ocorrência (`homogeneous` e `stasis` na mesma linha da
matriz) são dois *N_hits* e **uma** *N_ocorrencia*. Janelas KWIC
deslocadas da mesma passagem não devem inflacionar o *N*: a extracção
marca `janela_sobreposta` / `grupo_passagem_id`.

### 2.2 Duas portas, um cânone de análise

`textura_search.py` avalia uma consulta booleana (`AND OR NOR NOT NEAR/n * ?`)
com filtros opcionais de mesma frase e relação heurística. Escreve
`Results` e o ficheiro de termos adjudicados. **Não** é input de
`textura_analise.py`.

`textura_near.py` relê a matriz + o léxico, emparelha nó×termo dentro de
`NEAR/n`, classifica sintaxe (spaCy por omissão; EN se `--lingua todas`),
atribui `canonical_term` e escreve o schema-2 (`8_Concordancia`, Hits,
Ocorrências, `9_Excluidas`).

A análise (fase 2) lê o livro near, **sincroniza Hits a partir de
Concordância** (salvo `--so-leitura`) e conta só `nuclear = TRUE` após
`revisto_por_humano` / override. Gráficos nativos em
`6_Graficos_barras` e `6_Graficos_familias` são a fonte do gráfico
legado por termo; os `_*_g_freq_token.xlsx` são cópia. A fase 2 escreve
ainda `_g_freq_eixos.png` / `.svg` (famílias realizadas agrupadas pelo
eixo de **curadoria** em `dados/lexicos/eixos_curadoria.tsv`) e a folha
`curadoria_fluxo`. Esse eixo de curadoria **não** é a coluna `eixo` do
Excel (homogeneidade síncrona / invariância diacrónica).

### 2.3 Casamento lexical e diacríticos

O tokenizador aceita `À-ÿ`. O casamento é insensível a maiúsculas e,
desde 2026-09-05, **dobra diacríticos** (NFD + remoção de `Mn`, mais
ligaduras `æ/œ/ß`) em:

- pesquisa: `textura_query._wildcard_para_regex` / `forma_casa_padrao`
- extracção: `textura.tokenizacao._rx_palavra` / `_casa_em` / `indices_no`
- canónico: `textura_lexico._rx_padrao` / `canonical_de_forma`
- filtro do nó na matriz: `pipeline` compara `sem_diacriticos(NODE)`

Assim `estátic*` casa `estática` e `estatica`, e a etiqueta no Excel fica
`estátic`, não a forma de superfície. **Não** há stemming
interlingue: `estátic*` ≠ `static`; `homogéneo` (exacto) ≠ `homogeneous`
(esse cai em `homoge*`).

A primeira etiqueta do léxico cujo padrão casa ganha. Por isso
`homoge*` absorve `homogeneo` / `homogeneidade` antes de `homogéneo`.

### 2.4 Língua da execução (`--lingua`)

A língua é da **corrida**, não da linha. `NOS` (`dados/lexicos/nos.tsv`)
define o paradigma do nó por código (`en`: texture, textures, textural, …;
`pt`: textura, texturas, texturais, textural, …).

| Valor | Efeito |
|---|---|
| `en` | Só linhas cujo nó está no paradigma EN |
| `pt` / `fr` / `de` | Só o paradigma dessa língua + modelo spaCy do registo |
| `todas` (omissão) | União dos NOS; spaCy **por forma do nó** (`textura*` → PT, `texture`/`textural` → EN) |

A pesquisa booleana com `textur*` já atravessa línguas. A GUI e
`textura_search.py` e `textura_near.py` usam `--lingua todas` e
`--near 8` por omissão, para o inventário misto não perder *textura
estatica*. Os goldens EN/PT/FR continuam a fixar `--lingua en|pt|fr`
e o modelo spaCy correspondente.
Para tese numa só língua, force `--lingua pt` (ou `fr`) — não `todas` —
para a sintaxe não ser classificada com o modelo EN.

### 2.5 Nuclearidade, polaridade, eixo, família

`nuclear` deriva da relação sintáctica ∈ `RELACOES_NUCLEARES` (atributiva,
predicativa, genitiva, …) e de regras de triagem (incidental, metatexto,
citação repetida, coordenação, janela sobreposta). A revisão humana
prevalece.

Polaridade (`:E` / `:V` no léxico) e eixo
(`homogeneidade_sincronica` / `invariancia_diacronica` / `ambos`) são
função de `canonical_term` no campo adjudicado. Testes que cruzam eixo ou
polaridade com o termo saem com *V* de Cramér = 1: não é falha, é
colinearidade esperada.

Família terminológica: se o léxico tem
`etiqueta:E / homogeneity = homoge*`, esse nome ganha. Senão, heurística
de substantivo sobre as formas nucleares (não usar o adjectivo mais
frequente).

### 2.6 Associação e inferência (fase 2)

Medidas 2×2 hit-dentro-de-NEAR vs banda de referência (`--banda` >
`--near`): OR, *G*², MI, logDice, etc. Testes de contingência com
Monte Carlo quando a tabela é esparsa; BH nas famílias de testes.
`--plano-a-priori` (JSON/YAML) trava desduplicação / nulo de polaridade /
relações e **ganha** à CLI. `--kappa-cego` calcula κ de Cohen em
`hit_key` (nuclear + relação) → folha `16_Kappa`.

Unidade de desduplicação (`--desduplicacao`) é decisão de protocolo, não
optimização: `nenhuma` (omissão da análise de frequência por hit),
`contexto`, `ocorrencia`, `obra_termo`, …

### 2.7 Gráfico de frequência por eixo de curadoria

O gráfico legado `_g_freq_token.png` (uma cor por termo, via
`cor_canonical`) mantém-se. Um segundo gráfico,
`barras_agrupadas_por_eixo`, mostra **todas** as famílias lexicais
realizadas nas atribuições nucleares, agrupadas por eixo de curadoria
(`invariancia`, `campo_figurativo`, `semelhanca_interna`,
`consistencia_entre_exemplares`, `variabilidade`, `por_classificar`).

A decisão editorial (`retido` / `relacionado` / `excluido`) não usa cor:
barra sólida = retido; preenchimento mais claro com contorno = relacionado;
hachura (`///`) com fundo branco = excluído. Só há duas cores na figura
(ocorrências vs documentos). Termos com *N* ≤ 2 dentro de cada eixo
colapsam em `outros (k)`; a lista vai para
`_g_freq_eixos_legenda.txt`.

`--modo-tese` (omissão: desligado) tira título, subtítulo e rodapé da
imagem — a legenda vive no Word — e grava PNG a 300 dpi + SVG. A GUI tem
a mesma opção na caixa «Analisar Excel…».

A frequência **não** equivale à força de associação ao núcleo; as
medidas de associação constam do Apêndice.

---

## 3. O que o programa faz e o que não faz

**Faz:** ler a matriz KWIC; casar um campo lexical pré-mapeado junto do
nó; separar hit vs ocorrência; ajudar a rever; contar; testar associação
e polaridade; desenhar barras / famílias / eixos de curadoria / nuvens /
Sankey; exportar concordância DOCX.

**Não faz:** re-OCR; recuperar offsets no PDF original; detectar língua
por linha; fundir radicais de línguas diferentes (`static` / `estátic`);
substituir a adjudicação humana do campo.

---

## 4. Peças do sistema

```
textura_gui.py           orquestração no ambiente gráfico
textura_search.py        fase 1a — consulta booleana → Results + termos
textura_near.py          fase 1b — extracção schema-2
textura_analise.py       fase 2  — estatística e gráficos
textura_apendice.py      fase 3  — DOCX
textura_freq.py          relatório filtrável N_hits × documentos
textura_doctor.py        checklist antes da análise
textura_concordancia_qa.py  segunda opinião / QA
```

Caminho típico de ficheiros (pasta de trabalho, não o código):

```
TEXTURA_TUDO_MATRIZ_v12.xlsx
        │
        ▼
resultado_pesquisa.xlsx
resultado_pesquisa_termos_adjudicados.txt
        │
        ▼
resultado_pesquisa_near.xlsx          ← rever aqui
        │
        ▼
resultado_pesquisa_near_analise.xlsx  + _g_freq_*.xlsx/.png
                                        + _g_freq_eixos.png/.svg
                                        + _g_freq_eixos_legenda.txt
                                        + curadoria_fluxo.tsv
```

Raiz do código: `E:\PYTHON CODES\Term statistics`. Não use cópias no
Ambiente de Trabalho como fonte do programa.

---

## 5. Tutorial

Pré-requisitos: Python 3.10+, dependências do projecto, modelos spaCy
`en_core_web_sm` (e `pt_core_news_sm` / `fr_core_news_sm` se for correr
língua única). Trabalhe numa pasta só para a corrida
(ex.: `C:\Users\lmr20\Desktop\TESTE`).

### 5.1 Pela GUI (caminho habitual)

1. Abra `python textura_gui.py` a partir da pasta do código.
2. **Matriz:** o `.xlsx` KWIC (`TEXTURA_TUDO_MATRIZ_v12.xlsx`).
3. **Consulta** (exemplo misto EN/PT):

   ```
   textur* NEAR/8 (homoge* OR homogéneo OR immob* OR imóvel OR stasis OR estátic* OR invari* OR static*)
   ```

4. Deixe *mesma frase* e *relação sintáctica* ligados, salvo teste.
5. **Língua do nó (NEAR):** `todas` numa matriz mista; `pt` só em
   português; `en` só em inglês.
6. Indique a pasta/ficheiro de saída. Execute a pesquisa. A GUI chama a
   extracção NEAR a seguir.
7. Abra `*_near.xlsx` → folha `8_Concordancia`. Siga
   [`GUIA_REVISAO_FASE1.md`](GUIA_REVISAO_FASE1.md): células amarelas,
   `nuclear`, `motivo_exclusao`, `revisto_por_humano`.
8. Na GUI: **Analisar Excel…** sobre o `*_near.xlsx` (não sobre
   `resultado_pesquisa.xlsx`).
9. Leia `1_Resumo`, `6_Graficos_barras`, `6_Graficos_familias`,
   `curadoria_fluxo`, `0_Avisos`. Confira se Concordância e Hits têm o
   mesmo *N*. Ao lado do Excel: `_g_freq_token.png` (legado) e
   `_g_freq_eixos.png` (curadoria; com `--modo-tese` também `.svg` e
   `_g_freq_eixos_legenda.txt`).

### 5.2 Pela linha de comandos (reproduzível)

Na pasta do código:

```bat
python textura_search.py ^
  --xlsx "C:\Users\lmr20\Desktop\TESTE\TEXTURA_TUDO_MATRIZ_v12.xlsx" ^
  --consulta "textur* NEAR/8 (homoge* OR estátic* OR stasis OR immob* OR invari* OR static*)" ^
  --saida "C:\Users\lmr20\Desktop\TESTE\resultado_pesquisa.xlsx" ^
  --lingua todas
```

Isto escreve `resultado_pesquisa.xlsx`, o `.txt` de termos, e (por
omissão) chama o NEAR com o mesmo `--lingua` e o `--near` da consulta.

Revisão no Excel, depois:

```bat
python textura_analise.py --xlsx "C:\Users\lmr20\Desktop\TESTE\resultado_pesquisa_near.xlsx"
python textura_analise.py --xlsx "C:\Users\lmr20\Desktop\TESTE\resultado_pesquisa_near.xlsx" --modo-tese
```

Relatório de frequência à parte (opcional):

```bat
python textura_freq.py --xlsx "C:\Users\lmr20\Desktop\TESTE\resultado_pesquisa_near_analise.xlsx" --nuclear true --unidade ambas
```

### 5.3 Verificar se a corrida é coerente

| Verificação | Onde |
|---|---|
| Léxico lido | `Config_lexico` e o `.txt` ao lado |
| Tipos atestados | folha `Termos` da pesquisa (`atestado = sim`) |
| Espelho | `8_Concordancia` e `8_Concordancia_Hits`: mesmas linhas e mesmo *N* nuclear |
| Exclusões | `9_Excluidas` = linhas com `nuclear = FALSE` |
| Etiqueta certa | `estátic*` → coluna `canonical_term` = `estátic`, não `estaticas` |
| Avisos esperados | `0_Avisos`: colinearidade eixo/polaridade; termos sem `:E`/`:V` |
| Curadoria de tese | `_g_freq_eixos.png` + `curadoria_fluxo`; TSV `eixos_curadoria.tsv` |

Se Concordância e Hits divergirem e **não** quiser que a análise
reescreva Hits, use `textura_analise.py --so-leitura`.

---

## 6. Exemplos de aplicação

### Exemplo A — Campo inglês *homogeneity* (uma língua)

**Léxico** (`uniforme_en.txt`):

```
homoge:E / homogeneity = homoge*
stasis:E / stasis = stasis
immob:E / immobility = immob*
static:E / stasis = static*
```

**Consulta:** `textur* NEAR/8 (homoge* OR stasis OR immob* OR static*)`  
**NEAR isolado:** `--lingua en --near 8`

Esperado: `homogeneous`, `homogeneity`, `homogenous`, `immobility`,
`stasis`, `static texture`. Família do gráfico = *homogeneity* / *stasis*
/ *immobility* (nomes do `/ familia`, não o adjectivo mais frequente).

### Exemplo B — Matriz mista EN+PT sem diacríticos (o caso TESTE / v12)

A `TEXTURA_TUDO_MATRIZ_v12.xlsx` tem ~58 000 linhas. O contexto (col. 15)
está **sem acentos**, mas o nó inclui `texture` e `textura`.

**Léxico:**

```
homoge:E / homogeneity = homoge*
homogéneo:E / homogeneity = homogéneo
estátic:E / stasis = estátic*
static:E / stasis = static*
stasis:E / stasis = stasis
immob:E / immobility = immob*
imóvel:E / immobility = imóvel
invari:E / invariance = invari*
```

**Corrida:** pesquisa + NEAR com `--lingua todas`.

O que deve acontecer:

| Padrão | Casa na v12 | Etiqueta no Excel |
|---|---|---|
| `homoge*` | *homogeneous*, *homogeneidade*, *homogeneo* | `homoge` |
| `estátic*` | *estatica*, *estaticas*, *estaticidade* | `estátic` |
| `static*` | *static*, *statically* | `static` |
| `invari*` | *invariant*, *invariance*, *invariancias* | `invari` |
| `homogéneo` / `imóvel` | só se a forma (dobrada) existir e não tiver já caído em `homoge*` / `immob*` | a etiqueta da linha |

O que **não** deve acontecer: `nnvaria*` / `unvaria*` (typo); `invari`
sem `*` (fica cego a *invariant*); `--lingua en` no NEAR (corta nós
`textura` / `texturas` e perde *textura estatica*).

### Exemplo C — Só português, sintaxe PT

```bat
python textura_near.py ^
  --xlsx MATRIZ.xlsx --termos campo_pt.txt ^
  --near 8 --lingua pt --saida saida_pt_near.xlsx --so-extrair
```

Usa `pt_core_news_sm` e o paradigma `textura` / `texturas` / `texturais`.
Não misture este livro com o EN na mesma análise sem um passo explícito
de fusão (e uma coluna `fonte` / língua).

### Exemplo D — Plano a priori e segundo revisor

`plano.yaml`:

```yaml
desduplicacao: contexto
nulo_polaridade: lexico
```

```bat
python textura_analise.py --xlsx livro_near.xlsx --plano-a-priori plano.yaml --kappa-cego revisor_b.xlsx
```

O plano ganha à CLI. κ vai para `16_Kappa` e `1_Resumo`. Sem estes
ficheiros, o resumo mostra `—`.

### Exemplo E — Frequência filtrada depois da análise

```bat
python textura_freq.py --xlsx livro_analise.xlsx --nuclear true --termo homoge --unidade familia
```

Reutiliza os gráficos já implementados (`barras_horizontais_agrupadas`).
Não reabre a matriz.

---

## 7. Como escrever o léxico

Uma linha por tipo. Codificação UTF-8.

```
etiqueta [:E|:V] [/ nome_da_familia] = padrao1, padrao2
```

| Peça | Função | Exemplo |
|---|---|---|
| etiqueta | Canónico no Excel e nos gráficos de token | `estátic` |
| `:E` / `:V` | Polaridade estabilidade / variabilidade | `homoge:E` |
| `/ familia` | Nome da família (substantivo) | `/ homogeneity` |
| `*` | Truncatura à direita (até 20 caracteres de palavra) | `homoge*` |
| `?` | Um carácter (pesquisa booleana) | raro no NEAR |
| vários padrões | OR dentro do tipo | `homogene*, homogenous` |

Regras práticas:

- Um padrão **exacto** (`homogéneo`, `invari`, `stasis`) só casa essa
  forma (depois de dobrar acentos).
- `homoge*` já cobre o PT sem acento (`homogeneidade`). A linha
  `homogéneo = homogéneo` é redundante na v12; não é erro.
- Não invente radicais (`nnvaria*`, `unvaria*` se o corpus tem *invari-*).
- Não ponha o nó no campo (`textur* = …` do lado esquerdo é rejeitado).
- A ordem das linhas conta: o primeiro padrão que casa atribui a
  etiqueta.

O ficheiro gerado pela pesquisa (`*_termos_adjudicados.txt`) pode ser
editado à mão **antes** de voltar a extrair o NEAR.

A curadoria de tese (eixo + retido/relacionado/excluído) é outro
ficheiro: `dados/lexicos/eixos_curadoria.tsv`. Não substitui o léxico
de padrões; só classifica as famílias **já realizadas** no gráfico
`_g_freq_eixos`. Ver §2.7.

---

## 8. Línguas, acentos e nós

### Acentos

O código **reconhece** acentos e **dobra-os** no casamento. `é` e `e`
contam como a mesma letra para decidir se o padrão acerta. A forma
gravada em `matched_form` é a do corpus (`estatica`, não `estática`).

Isto serve EN (sem acentos) e PT/FR/ES (com acentos ou com acentos
retirados no OCR). Não traduz entre línguas.

### Nós (`NOS`)

A extracção NEAR **primeiro** filtra as linhas da matriz pelo valor da
coluna do nó (col. 6). `textur*` na consulta de pesquisa não substitui
esse filtro. Por isso `--lingua` muda o *N* mesmo com o mesmo léxico.

`textural` está em EN e em PT: janelas PT cujo nó KWIC é `textural`
sobrevivem a `--lingua en`. Janelas cujo nó é `textura` não.

### Modelos spaCy

| `--lingua` | Modelo (omissão) | Estado |
|---|---|---|
| `en` / `todas` | `en_core_web_sm` | validado |
| `pt` | `pt_core_news_sm` | validado |
| `fr` | `fr_core_news_sm` | validado |
| `de` | `de_core_news_sm` | não validado |

`todas` une os nós e **despacha o spaCy pela forma do nó**: `textura` /
`texturas` / `texturais` usam `pt_core_news_sm`; `texture` / `textural`
(partilhados) ficam no modelo EN. Não há detecção pelo texto da janela.
Para um inventário só FR (o nó escreve-se *texture*), use `--lingua fr`.

---

## 9. Folhas do Excel

### Pesquisa (`resultado_pesquisa.xlsx`)

| Folha | Uso |
|---|---|
| `Results` | Hits da consulta (não analisar) |
| `Termos` | Cada etiqueta: formas atestadas, *N*, sim/não |
| `Resumo` | Consulta, filtros, totais |
| `Config_lexico` | Campo usado |

### Extracção / análise (`*_near.xlsx`, `*_analise.xlsx`)

| Folha | Uso |
|---|---|
| `0_Instrucoes` | Comando, schema, *N* |
| `Config_lexico` | Etiqueta, padrões, polaridade, eixo |
| `8_Concordancia` | **Trabalho:** um hit por linha; amarelo = editável |
| `8_Concordancia_Hits` | Espelho do nível hit |
| `8_Concordancia_Ocorrencias` | Uma linha da matriz; só leitura |
| `9_Excluidas` | `nuclear = FALSE` |
| `Duplicados` | Passagens sobrepostas / vários caminhos |
| `1_Resumo` | Indicadores da fase 2 |
| `2_Frequencias` … `15_Perfis` | Estatística |
| `6_Graficos_barras` | Token canónico (fonte do gráfico legado) |
| `6_Graficos_familias` | Família substantivo |
| `curadoria_fluxo` | Brutas vs nucleares por termo + eixo/decisão de curadoria |
| `16_Kappa` | Só se passou `--kappa-cego` |
| `0_Avisos` | Colinearidades, gráficos falhados, polaridade em falta |

Dicionário de colunas: [`dados/dicionario_colunas.md`](dados/dicionario_colunas.md).

---

## 10. Contagens que se podem defender

| Símbolo | Definição |
|---|---|
| *N_hits* | Linhas nucleares em `8_Concordancia` |
| *N_ocorrencias* | `texture_occurrence_id` distintos com pelo menos um hit nuclear |
| *n_documentos* | Documentos (`doc_id`) com hit nuclear daquele termo/família |
| *N_dedup* | *N_hits* após a regra de `--desduplicacao` (se não for `nenhuma`) |

Declare sempre a unidade. Não some *N_hits* de duas extracções com
`--lingua` diferentes e chame a isso um corpus único, sem o dizer.

A pesquisa (`Results`) e o NEAR **não** têm de ter o mesmo *N*: a
pesquisa conta janelas da consulta heurística; o NEAR relê a matriz com
spaCy, filtro de nó e um hit por (ocorrência × tipo). Compare-os como
diagnóstico, não como dois ouro.

---

## 11. Referência CLI

Cwd recomendado: `E:\PYTHON CODES\Term statistics`.

### `textura_search.py`

| Flag | Papel |
|---|---|
| `--xlsx` | Matriz (obrigatório) |
| `--consulta` | Expressão booleana |
| `--saida` | Livro `Results` |
| `--termos` | Fecha o campo (senão deriva da consulta) |
| `--lingua` | Para o NEAR delegado. Omissão: `todas` |
| `--extrair-near` / `--sem-extrair-near` | Delegar ou não a `textura_near.py` |
| `--permitir-outra-frase` | Desliga o filtro de frase |
| `--sem-sintaxe` | Não exige relação heurística |
| `--col-no` `--col-ctx` `--col-src` | Omissão 6, 15, 12 |

### `textura_near.py`

| Flag | Papel |
|---|---|
| `--xlsx` `--termos` `--saida` | Matriz, léxico, livro schema-2 |
| `--near` | Raio (omissão **8**; a pesquisa propaga o da consulta se houver NEAR/n) |
| `--banda` | Banda de referência (omissão 12; > `--near`) |
| `--lingua` | Omissão **`todas`** |
| `--sintaxe` | `spacy` / `heuristica` |
| `--so-extrair` | Só fase 1 |
| `--modelo` | Override do spaCy |

### `textura_analise.py`

| Flag | Papel |
|---|---|
| `--xlsx` | Livro near (revisto) |
| `--so-leitura` | Não reescreve Hits a partir de Concordância |
| `--lexico` | Termos (`etiqueta / familia = …`) |
| `--unidade` | `canonical` / `familia` / `ambas` (frequências) |
| `--plano-a-priori` | JSON/YAML; prevalece |
| `--kappa-cego` | Segundo revisor → `16_Kappa` |
| `--desduplicacao` | Ver §2.6 |
| `--modo-tese` | `_g_freq_eixos` sem título na imagem; PNG 300 dpi + SVG (omissão: desligado) |

### `textura_freq.py`

`--xlsx` `--nuclear true|false|ambos` `--termo` `--forma` `--unidade`
`--saida` `--etiqueta` `--lexico`

---

## 12. Armadilhas frequentes

1. **Analisar `resultado_pesquisa.xlsx`.** A análise recusa ou produz
   lixo de schema. Use `*_near.xlsx`.
2. **`--lingua en` numa matriz com `textura`.** Corta nós `textura` /
   `texturas`. A omissão (GUI, pesquisa e NEAR) é `todas`. Só use `en`
   quando quiser um inventário só inglês.
3. **Padrão sem `*`.** `invari` ≠ `invari*`.
4. **Typo no radical.** `nnvaria*` / `unvaria*` = zero eterno.
5. **Achar que `estátic*` apanha *static*.** Não apanha. Ponha os dois
   padrões se o campo for bilingue.
6. **Zero hits ⇒ exclusão.** Se a etiqueta não está em `9_Excluidas` nem
   em Concordância, o padrão nunca casou (léxico vs grafia), não foi
   excluído.
7. **Editar `canonical_term` / `matched_form`.** Quebra IDs. Exclua a
   linha.
8. **Correr `carregar_base` / análise sem `--so-leitura`** quando Hits e
   Concordância foram deliberadamente divergentes (snapshot). A análise
   realinha Hits.
9. **Tratar `_g_freq_*.xlsx` como original.** São cópia; edite as folhas
   `6_Graficos_*` do livro `*_analise.xlsx`.
10. **κ e plano vazios.** Não foram pedidos na corrida; o *N* é de um
    só revisor.
11. **Ler a frequência como associação.** `_g_freq_eixos` conta
    atribuições nucleares. A força de associação está em `9_Associacao`
    / Apêndice. Revise `eixos_curadoria.tsv` se um termo estiver em
    `por_classificar`.

---

## 13. Glossário

| Termo | Sentido neste projecto |
|---|---|
| Matriz / KWIC | Livro mestre `Neighbor Contexts`; uma linha = uma ocorrência de textura |
| Nó | Forma de *textur\** na coluna 6 |
| Campo lexical | Propriedades adjudicadas (homoge, estátic, …), não o nó |
| Padrão | Expressão com `*` / `?` que casa tokens |
| Etiqueta / canónico | Balde em que a forma entra (`estátic`) |
| Família | Nome substantival do balde (`homogeneity`) |
| Hit | Um casamento campo×nó numa ocorrência |
| Nuclear | Hit que entra nas contas após regras + revisão |
| Incidental | Hit rejeitado (aposto, metatexto, fora de relação, …) |
| NEAR/*n* | Distância máxima em tokens, mesma frase por omissão |
| Dobragem de acentos | `é`≡`e` só para casar; a forma mostrada não muda |
| `--lingua` | Filtro do paradigma do nó + modelo spaCy da corrida |
| Eixo (pipeline) | Coluna `eixo` do Excel: homogeneidade síncrona / invariância diacrónica |
| Eixo de curadoria | Tabela `eixos_curadoria.tsv` (retido / relacionado / excluído) para o gráfico de tese |
| `--modo-tese` | Figura de eixos sem título/rodapé; a legenda vai no Word |

---

## Ver também

- [`TECHNICAL_MANUAL.md`](TECHNICAL_MANUAL.md) — arquitectura, fórmulas, CLI EN, bibliografia de métodos
- [`GUIA_REVISAO_FASE1.md`](GUIA_REVISAO_FASE1.md) — o que editar no Excel
- [`dados/dicionario_colunas.md`](dados/dicionario_colunas.md) — colunas de `8_Concordancia`
- [`dados/lexicos/INVENTARIO_FASE2.md`](dados/lexicos/INVENTARIO_FASE2.md) — TSV de léxicos
