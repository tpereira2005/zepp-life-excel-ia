# Exemplo fictício

Dados **inventados** e determinísticos (gerados por `zepp_life_excel_ia/sintetico.py`), para experimentar o projeto sem conta Zepp Life. Não correspondem a nenhuma pessoa.

| Ficheiro | O que é |
|---|---|
| [`exemplo-transformacao.xlsx`](exemplo-transformacao.xlsx) | O livro Excel gerado: Painel, resumos, fases, medidas manuais, medições e critérios. |
| [`exemplo-dados-ia.json`](exemplo-dados-ia.json) | O JSON para IA correspondente (esquema 2.1.0). |
| [`exemplo-dados-ia-resumo.json`](exemplo-dados-ia-resumo.json) | A versão compacta do JSON (agregados e leituras dos últimos 28 dias). |
| [`exemplo-config.json`](exemplo-config.json) | A configuração usada, com as opções `sexo`, `ano_nascimento`, `objetivo` e `fases`. |
| [`exemplo-exportacao-zepp.json`](exemplo-exportacao-zepp.json) | A exportação de onde partiu, no formato devolvido pelo leitor Zepp. |
| [`exemplo-medidas-manuais.csv`](exemplo-medidas-manuais.csv) | O formato do CSV opcional de medidas manuais iniciais (`csv_medidas_iniciais` no `config.json`). |

O exemplo cobre 150 dias e 20 semanas de medidas, com:

- pesagens só de peso (excluídas, com motivo) e repesagens no mesmo dia;
- dias sem pesagem e duas semanas sem medidas manuais (tracejado nos gráficos);
- uma semana marcada "Sem medição: esqueci-me";
- uma descida de peso seguida de uma subida lenta, para mostrar "desde o mínimo" e duas fases (*Défice* e *Manutenção*).

## Gerar de novo

```bat
.venv\Scripts\python scripts\gerar_demo.py
```

Com o Microsoft Excel instalado, o exemplo é confirmado no Excel real. As imagens do README capturam-se a partir do PowerShell:

```powershell
.\scripts\capturar-folhas.ps1 -Livro docs\exemplo\exemplo-transformacao.xlsx -Pasta docs\imagens -Intervalos @('Painel!A1:Q46','Painel!A47:Q104','Resumo Semanal!A1:Y14','Medidas Manuais!A1:G14','Fases!A1:S7')
```

Como os dados do exemplo acabam em junho de 2026, o Painel mostra o aviso vermelho de leituras em atraso. Para as imagens do README, o Painel foi capturado a partir de uma cópia temporária com a célula auxiliar `Painel!S6` (dias desde a última leitura) posta a 0.

`capturar-folhas.ps1` não substitui ficheiros: se já existir uma imagem com o mesmo nome, grava `-2`, `-3`, …; capture para uma pasta vazia e copie depois.
