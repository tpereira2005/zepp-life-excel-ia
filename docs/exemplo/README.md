# Exemplo fictício

Dados **inventados** e determinísticos (gerados por `zepp_life_excel_ia/sintetico.py`), para experimentar o projeto sem conta Zepp Life. Não correspondem a nenhuma pessoa.

| Ficheiro | O que é |
|---|---|
| [`exemplo-transformacao.xlsx`](exemplo-transformacao.xlsx) | O livro Excel gerado: Painel, resumos, medidas manuais, medições e critérios. |
| [`exemplo-dados-ia.json`](exemplo-dados-ia.json) | O JSON para IA correspondente (esquema 2.0.0). |
| [`exemplo-exportacao-zepp.json`](exemplo-exportacao-zepp.json) | A exportação de onde partiu, no formato devolvido pelo leitor Zepp. |
| [`exemplo-medidas-manuais.csv`](exemplo-medidas-manuais.csv) | O formato do CSV opcional de medidas manuais iniciais (`csv_medidas_iniciais` no `config.json`). |

O exemplo cobre 150 dias e 20 semanas de medidas, com:

- pesagens só de peso (excluídas, com motivo) e repesagens no mesmo dia;
- dias sem pesagem e duas semanas sem medidas manuais (tracejado nos gráficos);
- uma semana marcada "Sem medição: esqueci-me".

## Gerar de novo

```bat
.venv\Scripts\python scripts\gerar_demo.py
```

Com o Microsoft Excel instalado, o exemplo é confirmado no Excel real. As imagens do README capturam-se a partir do PowerShell:

```powershell
.\scripts\capturar-folhas.ps1 -Livro docs\exemplo\exemplo-transformacao.xlsx -Pasta docs\imagens -Intervalos 'Painel!A1:Q40','Painel!A41:Q98','Resumo Semanal!A1:V14','Medidas Manuais!A1:G14'
```

`capturar-folhas.ps1` não substitui ficheiros: se já existir uma imagem com o mesmo nome, grava `-2`, `-3`, …; capture para uma pasta vazia e copie depois.
