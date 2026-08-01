# Referências (APA 7)

Artefactos do subsistema de referências do apêndice de concordância.

| Ficheiro | Fase | Notas |
|---|---|---|
| `inventario.tsv` | 0 | Uma linha por `doc_id` (heurística `tipo_provavel` só indicativa) |
| `referencias_rascunho.tsv` | 1 | Extracção com colunas `evidencia_*` (rascunho; nunca inventa campos) |

Comando Fase 1:

```text
python utilitarios/gera_referencias.py --extrair --xlsx SEU_near.xlsx --raiz-corpus "E:\todos os textos"
```

Opcional: `--permitir-web` só resolve DOIs *já presentes* no PDF via doi.org/Crossref (`evidencia=doi_org`).
| `referencias_revisto.tsv` | 2 | Edição humana obrigatória |
| `referencias_apa7.tsv` / `.md` | 3 | Formatação APA 7 + citação curta |

**Nunca inventar metadados.** Campo sem evidência fica vazio e `verificar=sim`.
