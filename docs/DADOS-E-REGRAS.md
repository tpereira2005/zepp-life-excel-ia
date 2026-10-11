# Dados e regras

## Origem

O leitor (`zepp-reader.exe`) pede à API Zepp os registos de peso do membro `-1` (titular da conta), página a página, e devolve-os com o instante (com desvio horário), as métricas da balança/aplicação e `ProfileID`. Registos com `weightType ≠ 0` são ignorados pelo leitor (valores de peso inválidos).

> `ProfileID = -1` é o membro **pedido** à API, não um atributo verificado de cada registo. A alternativa por altura (`altura_perfil_sem_id_cm`) só serve exportações antigas que não tinham esse campo.

## Validação de cada descarga

Antes de qualquer cálculo, `python -m zepp_life_excel_ia validar-exportacao` (código em `exportacao.py`) recusa a nova descarga se:

- estiver vazia;
- algum registo não tiver data com fuso horário ou peso positivo;
- houver dois registos no mesmo instante UTC;
- faltar alguma pesagem da descarga anterior (comparação por instante UTC). Com `-PermitirRemocoes` isto passa a aviso.

Uma descarga recusada não substitui nada.

## Seleção das leituras

Por ordem:

1. **Antes do início** — data local < `inicio_transformacao`: contada, não listada.
2. **Composição incompleta** — falta algum de: IMC, gordura, água, massa óssea, metabolismo, músculo, gordura visceral, massa proteica, proteína, idade corporal, pontuação. Na prática são pesagens só de peso, sem bioimpedância (por exemplo, com meias ou sem contacto dos pés).
3. **Valor implausível** — fora de limites largos. Apanham erros de leitura, não julgam a composição:

   | Campo | Limites | Campo | Limites |
   |---|---|---|---|
   | Peso | 30–250 kg | Massa muscular | 15–150 kg |
   | IMC | 10–60 | Proteína | 5–35 % |
   | Gordura | 2–70 % | Gordura visceral | 1–59 |
   | Água | 25–80 % | Metabolismo basal | 700–4500 kcal/dia |
   | Massa óssea | 0,5–8 kg | Idade corporal | 10–100 anos |
   | | | Pontuação | 0–100 |

4. **Outro perfil**.

Cada pesagem excluída desde o início aparece, com o motivo, na folha *Critérios e Qualidade* e em `pesagens_excluidas` no JSON.

Instantes repetidos (mesmo instante UTC) interrompem a atualização. Duas pesagens à mesma hora local na mudança de hora de outubro são **distintas** (desvios `+01:00` e `+00:00`).

## Protocolo de pesagem

Definido em `config.json` e mostrado no Excel e no JSON. A hora do relógio **não** é usada como filtro: se a pesagem é sempre feita ao acordar, uma leitura às 11h num domingo continua a ser comparável.

## Agregações

| Conceito | Definição |
|---|---|
| Semana | Domingo a sábado, hora local. |
| Média do período | Média aritmética das leituras aceites (por leitura, não por dia). Repesagens no mesmo dia contam ambas; *Dias com medição* mostra quantos dias distintos houve. |
| Variação | Média do período − média do período anterior. Gordura em pontos percentuais (pp). |
| Período parcial | Começa antes do início da transformação ou termina depois da última leitura. |
| Média móvel de 7 dias | Média das leituras em ]t − 7 dias, t]. |
| Desde o início (Painel) | Média dos últimos 7 dias − média dos primeiros 7 dias (reduz o ruído diário de uma leitura isolada). |
| Ritmo recente | (Média dos últimos 7 dias − média da janela de há 3 semanas) / 3, em unidade por semana. |
| Última leitura | A de maior instante UTC (coluna oculta *Instante UTC*), não a de maior hora local: na mudança de hora de outubro duas pesagens podem ter a mesma hora local. |
| Massa gorda | Peso × gordura / 100. |
| Massa proteica | Derivada pelo leitor: peso × proteína / 100 (não vem da aplicação). |

## Medidas manuais

- Uma linha por semana; cintura ao nível do umbigo e pescoço, idealmente ao domingo.
- O valor registado é a média de 3 leituras arredondada a 0,1 cm; as leituras individuais, quando escritas nas observações no formato `Leituras — cintura: a - b - c; pescoço: x - y - z`, são extraídas para o JSON.
- Variação da cintura = medição − medição anterior disponível (semanas sem medição são saltadas).
- Estados: *Completa*, *Incompleta* (só uma medida), *Sem medição* (sem valores mas com observação, ex. férias), *Por preencher*.

## Gráficos

- Peso, gordura e músculo: pontos = leituras; linha = média móvel de 7 dias.
- Cintura e pescoço: linha contínua entre medições consecutivas; **tracejado** sobre semanas sem medição; pontos apenas em medições reais; nada depois da última medição.
- Eixos Y com amplitude mínima (peso 6 kg, gordura 4 pp, músculo 3 kg, cintura 10 cm, pescoço 3 cm), margem de 12 % e 5–7 marcas redondas, para não ampliar ruído de medição.
- Escalas e tracejados são calculados quando o livro é gerado; ao editar no Excel, só mudam depois de recalcular.

## Onde está no código

| Regra | Ficheiro |
|---|---|
| Métricas, campos obrigatórios, limites de plausibilidade, indicadores | `zepp_life_excel_ia/metricas.py` |
| Validação da descarga | `zepp_life_excel_ia/exportacao.py` |
| Fuso horário, seleção e motivos de exclusão | `zepp_life_excel_ia/leituras.py` |
| Semanas, CSV inicial, preservação e validação das medidas manuais | `zepp_life_excel_ia/medidas.py` |
| Períodos, variações, médias móveis e indicadores | `zepp_life_excel_ia/resumos.py` |
| Amplitudes mínimas e escalas dos gráficos | `zepp_life_excel_ia/excel/folha_painel.py`, `excel/graficos.py` |

Os cálculos em `resumos.py` são a fonte única: o JSON usa-os diretamente e o Excel (com fórmulas nativas) é confirmado contra eles em cada atualização.

## Limitações

Composição corporal, idade e pontuação são estimativas da aplicação por bioimpedância, sensíveis à hidratação e às condições de medição. Não são diagnósticos clínicos.
