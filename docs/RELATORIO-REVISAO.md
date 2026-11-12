# Relatório das revisões 2.0.0 a 2.2.0 (para revisão por outro agente)

Este relatório descreve o que mudou, porquê, e como foi verificado, por ordem cronológica. Os valores pessoais da verificação com dados reais estão num relatório privado fora do repositório (`privado/notas/VERIFICACAO-DADOS-REAIS.md`).

> **Nomes atuais.** As secções 2.0.x referem a estrutura dessa altura: o pacote `zepp_transformacao/` é hoje `zepp_life_excel_ia/`, `excel.py` é o pacote `excel/` (um módulo por folha), `leituras.METRICAS` passou para `metricas.py` e a validação da exportação passou do PowerShell para `exportacao.py`. Ver a secção 2.1.0 e [ARQUITETURA.md](ARQUITETURA.md).

> **A revisão mais recente (2.2.0) está no fim deste documento.**

## Âmbito aprovado

Plano em 6 fases, aprovado na íntegra: (0) cópias e base de controlo, (1) correções críticas no atualizador, (2) substituição da biblioteca proprietária por Python/openpyxl, (3) conteúdo e design do Excel, (4) JSON 2.0.0, (5) testes, (6) arrumação e documentação. Decisões do utilizador: médias por leitura (mantidas) com "Dias com medição"; sem filtro por hora (protocolo: pesagem ao acordar, logo depois de ir à casa de banho); dados pessoais fora do Git; JSON 2.0.0 sem chaves antigas; repositório público.

## Problemas encontrados e correção

| # | Problema | Correção | Onde verificar |
|---|---|---|---|
| A | Biblioteca de geração proprietária ("internal evaluation and testing"), runtimes com caminhos fixos | Reescrita com openpyxl; `.venv` + `requirements.txt` | `zepp_transformacao/excel.py`, `INSTALAR.cmd` |
| B | Bloqueio preso com a janela de erro aberta | Libertação no `finally` antes de qualquer pausa | `scripts/atualizar-zepp.ps1` (bloco `finally`); teste 4 e 6 de `tests/atualizador.Tests.ps1` |
| C | Pastas temporárias acumuladas | Apagadas após sucesso; máximo de 5 falhas guardadas | idem; teste 1 e 8 |
| D | Verificação de perdas com precedência errada; `ProfileID` tautológico | Comparação por instante UTC de todas as leituras antigas; documentação honesta | `atualizar-zepp.ps1`; teste 3; `docs/DADOS-E-REGRAS.md` |
| E | Duplicados por hora local (mudança de hora) | Chave = instante UTC; JSON com `data_hora_utc` e `desvio_utc` | `leituras.py`; `test_mudanca_de_hora_nao_gera_duplicado` |
| F | Medidas manuais por posição; datas em texto em formato EUA | Leitura por cabeçalho; dd/mm/aaaa; números com vírgula | `medidas.py`; `test_medidas.py` |
| G | Testes antigos com caminhos absolutos e valores fixos | Novo conjunto com dados fictícios; antigos arquivados (privado) | `tests/` |
| H | "Desde o início" com uma leitura isolada | Média dos primeiros vs últimos 7 dias | Painel; `resumos.indicadores` |
| I | Massa proteica descrita como da aplicação | Origem "derivada" | `leituras.METRICAS` |
| J/K | Ruído de vírgula flutuante; chaves = cabeçalhos | Arredondamento a 3 casas; snake_case | `ia_json.py`; `schema/` |
| — | Variação de cintura 0,0 aparecia em branco no Resumo Semanal | Fórmula avaliada pelo Excel real (0,0) | comparação com a referência (abaixo) |

## Como foi testado

1. **Cópia de segurança** completa (ZIP + manifesto SHA-256) antes de qualquer alteração, e retrato de referência do Excel/JSON anteriores (privado).
2. **pytest** (31 testes, dados fictícios): seleção e motivos, fuso/DST, duplicados, semanas, CSV, leitura por cabeçalho, prioridades de preservação (edição, apagamento, data alterada), novas semanas/meses, variações, médias móveis, indicadores, eixos, estrutura do livro, tracejados (`dPt`) e `dispBlanksAs=span`, JSON validado pelo JSON Schema, ausência de chaves sensíveis e do identificador do dispositivo, hash Excel↔JSON, e **validação no Microsoft Excel real** (todas as células de resumos, médias móveis, variações e cartões do Painel comparadas com Python; zero erros de fórmula).
3. **Testes do atualizador** (`tests/atualizador.Tests.ps1`, leitor simulado e pastas privadas de teste): primeira atualização, nova pesagem com histórico, exportação que perde pesagens (recusada, Excel inalterado), bloqueio libertado após falha, Excel aberto, atualização simultânea, recálculo sem descarga, Excel ilegível (recusado sem perda), esquecer conta.
4. **Dados reais** com `RECALCULAR-SEM-DESCARREGAR` (sem contactar a Xiaomi): todas as leituras e todos os valores das métricas idênticos à referência; todas as semanas manuais (data, cintura, pescoço, observações) idênticas; todos os resumos semanais e mensais iguais à referência a 1e-3, exceto a correção do 0,0 acima; JSON válido; hash correto; verificado no Excel real.
5. **Revisão visual** no Excel real (capturas das folhas e gráficos).

## Correções após revisão externa (2.0.2)

| Ponto | Causa | Correção | Teste |
|---|---|---|---|
| Última pesagem errada na mudança de hora | `INDEX(MATCH(MAX(Data/hora)))` pela hora local devolvia a primeira de duas pesagens à mesma hora local | Coluna oculta *Instante UTC*; Painel usa `MATCH(MAX(Instante UTC))`; janelas Python usam min/max da hora local, como o Excel | `test_ultima_pesagem_na_mudanca_de_hora` (validado no Excel real; falha com o código anterior) |
| Erro com todas as datas manuais vazias | `min()` sobre lista vazia na escala do eixo de datas | Usar o início das semanas como alternativa | `test_gera_sem_nenhuma_data_manual` (validado no Excel real; falha com o código anterior) |
| Verificação no Excel intermitente | Chamadas COM rejeitadas/leituras parciais com o Excel ocupado | `scripts/excel-com.ps1` (IOleMessageFilter, esperas por `Ready`), folhas esperadas obrigatórias, até 3 tentativas em Python | `test_comparacao_acusa_folhas_em_falta`; 12 verificações seguidas com dados reais, 12 sem falhas |

## Não feito / limitações

- O leitor Go não foi recompilado localmente (Go não instalado); o executável anterior continua em uso. A compilação e os testes Go correm no CI.
- A descarga real (`ATUALIZAR-EXCEL.cmd`) não foi executada nesta revisão para não acionar autenticação Xiaomi sem o utilizador presente.
- `VisceralFat`/`BasalMetabolism` continuam truncados para inteiro no leitor (documentado).

## Comandos para reproduzir (versão atual)

```bat
INSTALAR.cmd
.venv\Scripts\python -m pip install -r requirements-dev.txt
.venv\Scripts\python -m ruff check .
.venv\Scripts\python -m pytest -rs
powershell -NoProfile -ExecutionPolicy Bypass -File tests\atualizador.Tests.ps1
.venv\Scripts\python scripts\gerar_demo.py
.venv\Scripts\python scripts\comparar_saidas.py antes.xlsx antes.json depois.xlsx depois.json
```

## Reorganização do código (2.1.0)

Objetivo: facilitar a manutenção sem mudar o resultado.

| Problema | Alteração | Onde |
|---|---|---|
| Endereços das células repetidos entre o gerador e a verificação | Cada folha regista `Esperado(folha, célula, valor)` ao escrever a fórmula; a verificação percorre a lista | `excel/folha_*.py`, `verificacao.py` |
| Cada métrica definida em 4–5 sítios | Registo único `Metrica` + listas de ordem (`COLUNAS_MEDICOES`, `RESUMO`, `MEDIAS_MOVEIS`) | `metricas.py` |
| `excel.py` com 632 linhas e uma função de ~270 | Pacote `excel/` com estilos, gráficos e um módulo por folha; colunas pelos cabeçalhos | `excel/` |
| Dicionários com chaves em texto | `Periodo`, `Indicador`, `IndicadorManual`, `Calculos` | `resumos.py` |
| Ponto de entrada ligado ao argparse | `gerar.gerar()` + `__main__.py` com subcomandos | `gerar.py`, `__main__.py` |
| Validação da exportação em PowerShell, sem testes unitários | `exportacao.validar` com 8 testes; `Invoke-Python` único no atualizador | `exportacao.py`, `scripts/atualizar-zepp.ps1` |
| Conversões de datas do Excel repetidas | `datas.py` | |
| Sem verificador de código | ruff no CI; `pyproject.toml` | |

**Como foi verificado:** antes de alterar, geraram-se referências (Excel + JSON) com o código anterior para três casos — dados reais, dados fictícios e um caso com duas pesagens à mesma hora local na mudança de hora. Depois de cada passo, um comparador verificou cada célula (valor, fórmula, número, letra, preenchimento, bordas, alinhamento), células unidas, larguras, alturas, painéis fixos, tabelas, validações, formatação condicional, o XML dos gráficos e dos desenhos, e todo o JSON (exceto data de geração, versão e hash): **0 diferenças** nos três casos. A nova lista de células esperadas cobre as mesmas células que a verificação anterior (ex.: 227 × 4 em Medições, 34 × 19 no Resumo Semanal) e foi confirmada no Excel real. Testes: 42 pytest, testes do atualizador, ruff sem avisos. O comparador usado ficou no repositório: `scripts/comparar_saidas.py`.

Nota encontrada durante o trabalho: no Windows PowerShell 5.1, redirecionar o stderr do Python com `*>>` embrulhava cada linha num registo de erro ("python.exe : ERRO: …") partido pela largura da consola, pelo que a mensagem de erro não era encontrada. O atualizador passou a gravar cada linha como texto simples.


## Indicadores da fase atual, JSON compacto e proteção dos dados (2.2.0)

Revisão crítica completa do projeto, seguida de um plano por prioridades aprovado na íntegra pelo utilizador, incluindo seis decisões: (D1) mínimo/máximo automáticos e fases opcionais em `config.json`; (D2) estimativa pela fita com `sexo` em `config.json`; (D3) manter o JSON completo e acrescentar um ficheiro compacto; (D4) guardar a palavra-passe passa a ser uma escolha; (D5) retenção do histórico (30 dias + uma cópia por mês); (D6) manter a regra "médias por leitura" para repesagens.

### Problemas encontrados e o que foi feito

| # | Problema | Alteração | Onde |
|---|---|---|---|
| 1 | O Painel só comparava com o início: com uma descida seguida de uma subida, a recuperação desde o ponto mais baixo não aparecia em lado nenhum | Mínimo da média de 7 dias do peso (valor e data) e variações "desde o mínimo" de peso, massa gorda, massa magra e cintura | `resumos.extremos`, `resumos.indicadores`, `excel/folha_painel.py` |
| 2 | Sem massa magra nem repartição da variação do peso | Métrica derivada `massa_magra` no registo; `Composicao` (desde o início e desde o mínimo) | `metricas.py`, `leituras.py`, `resumos._composicao` |
| 3 | Ritmo recente com dois pontos (média de 7 dias agora vs. há 3 semanas) | Regressão linear das leituras dos últimos 28 dias; coluna auxiliar oculta + `SLOPE` no Excel | `resumos.ritmo_semanal`, `excel/folha_medicoes.py` |
| 4 | Cintura e pescoço pouco aproveitados | Cintura ÷ altura; gordura pela fita (Marinha dos EUA, só sexo masculino); mínimo da cintura | `resumos.gordura_fita`, `excel/folha_resumos.py` |
| 5 | Sem noção de fases | `fases` opcionais em `config.json`; folha *Fases*; chave `fases` no JSON | `config.py`, `resumos.fases`, `excel/folha_fases.py` |
| 6 | "Últimos 7 dias" ancorados na última leitura sem aviso de dados antigos | Subtítulo do Painel a vermelho após 3 dias; `qualidade_dados.dias_desde_ultima_leitura`; aviso na consola | `excel/folha_painel.py`, `ia_json.py`, `gerar.py` |
| 7 | JSON grande demais para muitas IA | Ficheiro compacto derivado do completo (tabelas `{colunas, linhas}`, leituras dos últimos 28 dias) | `ia_json.construir_resumo` |
| 8 | Leituras de bioimpedância anómalas passavam sem sinal | Assinaladas (nunca excluídas) face à mediana das vizinhas a ±7 dias | `resumos.atipicas`, `excel/folha_criterios.py` |
| 9 | Risco teórico: leitura exatamente no limite de uma janela de 7 dias podia dar valores diferentes no Excel e no Python e bloquear a atualização | Meio segundo de margem em todos os critérios de data/hora das fórmulas | `excel/estilos.py` (`depois_de`, `ate`, `antes_de`, `desde`) |
| 10 | As medidas manuais só existiam no Excel gerado; apagar o Excel fazia-as voltar ao CSV inicial sem aviso | Cópia `privado/medidas-manuais.csv` em cada atualização, usada (com aviso) se o Excel não existir | `medidas.escrever_copia/ler_copia`, `gerar._medidas_manuais` |
| 11 | Histórico e tentativas falhadas sem limite | 30 dias completos + a última cópia de cada mês; tentativas falhadas com mais de 30 dias apagadas | `atualizar-zepp.ps1` (`Limpar-Historico`) |
| 12 | Palavra-passe sempre guardada | Pergunta ao iniciar sessão; `-EsquecerPalavraPasse` | `atualizar-zepp.ps1` |
| 13 | Nada garantia que o leitor (que recebe a palavra-passe) era o mesmo de sempre | Hash SHA-256 registado na primeira utilização e confirmado antes de cada execução; o CI publica o hash | `atualizar-zepp.ps1` (`Confirmar-Leitor`), `.github/workflows/ci.yml` |
| 14 | Pesagem apagada na Zepp obrigava a repetir com `-PermitirRemocoes` | Código de saída 2 em `validar-exportacao`; o atualizador pergunta (resposta `SIM`) | `exportacao.PesagensRemovidas`, `__main__.py`, `atualizar-zepp.ps1` |
| 15 | Caminhos do atualizador sem teste; CI só em Python 3.12 | 6 cenários novos; CI em 3.11, 3.12 e 3.14 | `tests/atualizador.Tests.ps1`, `ci.yml` |

### O que **não** mudou (regras decididas com o utilizador)

Semanas de domingo a sábado em hora local; médias por leitura (repesagens contam todas); última leitura pelo instante UTC; os valores do Excel ganham ao CSV, incluindo campos apagados; `null` nunca é zero; sem filtro por hora. A folha *Medidas Manuais* (a única que o utilizador edita) não foi alterada. Nenhuma leitura passou a ser excluída: as atípicas são só assinaladas.

### Como foi testado

1. **Cópia de segurança antes de tudo:** ZIP de `config.json`, `Excel e JSON/` e `privado/` (392 ficheiros) com manifesto SHA-256 em `privado/backups/`, relido e conferido ficheiro a ficheiro. Trabalho num ramo próprio (`melhorias-2.2.0`).
2. **Valores anteriores inalterados:** antes de alterar o código geraram-se referências (Excel + JSON) para dados reais, fictícios e mudança de hora; depois, todos os valores que existiam no JSON antigo foram comparados com o novo. Diferenças: só `versao_esquema`, `gerador`, o hash do Excel, a descrição dos indicadores e o ritmo (mudança deliberada). Nenhuma outra, nos três casos.
3. **Fórmulas novas confirmadas no Microsoft Excel real** contra o Python (zero diferenças, zero erros de fórmula): dados reais com a configuração real; dados reais com uma configuração de teste (sexo, fases, incluindo uma fase futura sem leituras); dados fictícios com fases; mudança de hora. Cada célula nova tem `Esperado`: massa magra, coluna auxiliar do ritmo, mínimo e data do mínimo, todos os cartões, cintura ÷ altura, gordura pela fita e todas as colunas da folha *Fases*.
4. **pytest:** 56 testes (42 anteriores ajustados + 14 novos em `tests/test_indicadores.py`): massa magra fecha a conta com o peso; mínimo ignora os primeiros 6 dias e resolve empates pela leitura mais antiga; ritmo é o declive e respeita os mínimos; fórmula da fita (valor de referência e casos sem resultado); mínimo da cintura; atípicas assinaladas e não excluídas; qualidade dos dados; fases (JSON validado pelo esquema, folha, Excel real); configuração inválida; JSON compacto (esquema próprio, menos de 1/3 do tamanho, agregados iguais aos do completo); cópia das medidas manuais (ida e volta, reposição, prioridade do Excel, cópia ilegível); e **leituras exatamente a 7, 14, 21 e 28 dias da última**, confirmadas no Excel real.
5. **Atualizador:** 15 cenários com leitor simulado e pastas temporárias (9 anteriores + falha a meio da substituição com reposição, Excel apagado com reposição pela cópia CSV, retenção do histórico, leitor alterado recusado, sessão expirada com palavra-passe guardada, esquecer só a palavra-passe).
6. **Dados reais** com `RECALCULAR-SEM-DESCARREGAR` (sem contactar a Xiaomi): ver o relatório privado `privado/notas/`.
7. **Revisão visual** no Excel real (capturas do Painel, Fases, Resumo Semanal e Critérios), com dados reais e com o exemplo fictício.
8. `ruff check .` sem avisos.

### Não feito / limitações

- A descarga real (`ATUALIZAR-EXCEL.cmd`) não foi executada: exigiria a sessão Xiaomi do utilizador. Os caminhos de autenticação novos (pergunta sobre a palavra-passe, pergunta sobre pesagens em falta) são interativos e foram verificados por leitura do código; os caminhos não interativos equivalentes têm teste. O fluxo de verificação Xiaomi no navegador continua sem teste automático.
- O leitor Go não foi recompilado localmente (Go não instalado). O hash registado na primeira utilização fixa o executável atual, seja ele qual for: confirmar uma vez que corresponde ao publicado pelo CI.
- A estimativa pela fita só existe para sexo masculino (a fórmula feminina exige a anca).
- As fórmulas do Excel continuam a só ser confirmadas onde há Microsoft Excel (no CI esse passo é ignorado com aviso).
- As notas dos cartões (texto) não têm valor esperado; os números que mostram vêm de células que têm.
- Semanas novas continuam a nascer com a data de domingo pré-preenchida.
- Os pacotes do leitor Go herdados do projeto original e não usados continuam no repositório.
