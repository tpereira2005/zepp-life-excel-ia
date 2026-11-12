# Registo de alterações

## 2.2.0 — 2026-11-11

Indicadores que mostram a fase atual (e não só a comparação com o início), JSON compacto para IA e mais proteção dos dados. Os valores que já existiam no Excel e no JSON não mudaram (confirmado valor a valor com dados reais, fictícios e mudança de hora), com uma exceção deliberada: o **ritmo recente**.

**Indicadores e Excel**

- **Mínimo do peso**: a média de 7 dias mais baixa (a contar do 7.º dia de leituras) e a sua data; novo cartão *Peso desde o mínimo*.
- **Massa magra** (peso − massa gorda): coluna em Medições Corporais e nos resumos; cartões *Massa gorda/magra desde o início* e *desde o mínimo*, com a percentagem da variação do peso que foi gordura.
- **Ritmo recente** passa a ser o declive da regressão linear das leituras dos últimos 28 dias (antes: diferença entre duas médias de 7 dias, mais sensível a uma semana atípica). Exige 8 leituras em pelo menos 14 dias.
- **Cintura**: mínimo, data do mínimo e *Cintura desde o mínimo*; **cintura ÷ altura** no Resumo Semanal e no Painel.
- **Gordura pela fita** (método da Marinha dos EUA: cintura, pescoço e altura), uma estimativa independente da balança. Só com `"sexo": "masculino"` em `config.json` (a fórmula feminina exige a anca).
- **Fases** opcionais em `config.json` (nome, data de início, objetivo): nova folha *Fases* com as médias de 7 dias no início e no fim de cada fase, variações de peso, gordura, massa gorda, massa magra e cintura, e kg/semana.
- **Aviso de leituras em atraso**: o subtítulo do Painel fica a vermelho quando a última leitura tem mais de 3 dias.
- Painel com 16 cartões (4 filas); os cartões do pescoço e *Músculo desde o início* deram lugar aos novos (o músculo desde o início passa para a nota do cartão *Músculo*; o pescoço mantém o gráfico e a nota do último cartão).
- **Leituras atípicas** assinaladas (nunca excluídas) na folha *Critérios e Qualidade*, que ganha também um bloco *Qualidade dos dados*.
- Fórmulas com janelas de datas (médias móveis, primeiros/últimos 7 dias) com meio segundo de margem: uma leitura exatamente no limite da janela deixa de poder dar um resultado diferente no Excel e no Python e bloquear a atualização.

**JSON (esquema 2.1.0, só chaves novas)**

- `contexto` (altura, sexo, ano de nascimento, objetivo, data de referência), `fases`, `qualidade_dados` (dias desde a última leitura, cobertura, leituras atípicas).
- `indicadores`: massa gorda e massa magra, valores no mínimo do peso, mínimo/máximo da média de 7 dias, `composicao_da_variacao`, mínimo da cintura, `cintura_altura`, `gordura_fita_percentagem`.
- `massa_magra_kg` em cada medição; `massa_magra_media_kg` nos resumos; `cintura_altura` e `gordura_fita_percentagem` nas medidas manuais e no resumo semanal.
- Novo ficheiro **`dados-zepp-life-ia-resumo.json`** (cerca de 10 % do tamanho): os mesmos agregados em formato tabular e só as leituras dos últimos 28 dias. Esquema próprio em `schema/dados-zepp-life-ia-resumo.schema.json`.

**Atualização e segurança**

- **Cópia das medidas manuais** em `privado\medidas-manuais.csv` a cada atualização; se o Excel for apagado, as medidas são repostas a partir dela (antes voltavam a ser só as do CSV inicial).
- **Retenção do histórico**: todas as cópias dos últimos 30 dias e uma por mês das mais antigas (`-ManterHistoricoDias`; 0 = nunca apagar). Tentativas falhadas com mais de 30 dias são apagadas.
- **Palavra-passe opcional**: ao iniciar sessão pergunta se a guarda; `-EsquecerPalavraPasse` apaga só a palavra-passe.
- **Leitor fixado por SHA-256** (`bin\zepp-reader.exe.sha256`, criado na primeira utilização): um executável diferente é recusado antes de receber qualquer segredo. O CI publica o hash junto do executável.
- Pesagens desaparecidas na Zepp: o atualizador pergunta (é preciso escrever `SIM`) em vez de obrigar a repetir com `-PermitirRemocoes`; sem janela interativa continua a recusar.

**Testes e CI**

- 56 testes pytest (14 novos) e 15 cenários do atualizador (6 novos: reposição após falha a meio da substituição, cópia CSV, retenção, leitor alterado, sessão expirada com palavra-passe guardada, esquecer só a palavra-passe).
- CI também em Python 3.11 e 3.14.

## 2.1.0 — 2026-10-11

Reorganização interna do código. O Excel (células, fórmulas, estilos, gráficos) e o JSON gerados são idênticos aos da versão 2.0.3, confirmado célula a célula em três casos (dados reais, fictícios e mudança de hora).

- Pacote renomeado para `zepp_life_excel_ia`; linha de comandos `python -m zepp_life_excel_ia gerar | validar-exportacao`.
- `metricas.py`: registo único das métricas (antes repetidas em 4–5 sítios).
- `excel/`: o antigo `excel.py` (632 linhas, uma função de ~270) dividido em estilos, gráficos e um módulo por folha; letras de coluna obtidas pelos cabeçalhos.
- Verificação no Excel sem endereços repetidos: cada folha regista o valor esperado onde escreve a fórmula (`Esperado`); `verificacao.comparar` só percorre essa lista.
- `resumos.py`: `Periodo`, `Indicador`, `IndicadorManual` e `Calculos` (dataclasses) em vez de dicionários com chaves em texto.
- `gerar.gerar()` é uma função normal (já não recebe os argumentos da linha de comandos).
- Validação da exportação passou do PowerShell para Python (`exportacao.py`, com testes); o atualizador usa uma única função para chamar o Python e passa a mostrar o motivo exato das recusas.
- `datas.py`: conversão de datas do Excel num só sítio (estava repetida em 4).
- `pyproject.toml` (substitui `pytest.ini`; versão num só sítio) e ruff no CI.
- Documentação revista: README (pastas, validação, desenvolvimento, índice), instalação (atualizar, problemas comuns), utilização (opções do atualizador, registos), regras (validação, limites, onde está no código), JSON (versões, validação) e arquitetura (manutenção); `docs/exemplo/README.md`.
- `scripts/comparar_saidas.py`: compara dois pares Excel+JSON célula a célula, para confirmar alterações estruturais.

## 2.0.3 — 2026-10-11

- Projeto renomeado para **zepp-life-excel-ia** (pasta e repositório).
- Nova organização das pastas pessoais (fora do Git): `Excel e JSON/` só com os resultados; `privado/` com `entradas/`, `exportacao-zepp/`, `historico/`, `temporario/`, `backups/`, `arquivo/` e `notas/`. Substitui `inputs/`, `outputs/zepp-life/`, `arquivo/` e `docs/privado/`.
- `atualizar-zepp.ps1`: novo parâmetro `-PastaResultados`; `-PastaDados` aponta para `privado/`.

## 2.0.2 — 2026-10-11

- Mudança de hora: o Painel identifica a última pesagem pelo instante UTC (nova coluna oculta *Instante UTC* em Medições Corporais), não pela hora local. Antes, duas pesagens à mesma hora local (01:00–02:00 do último domingo de outubro) podiam mostrar o peso errado como "último".
- As janelas de 7 dias e de 4 semanas usam o mínimo/máximo da hora local, tal como as fórmulas do Excel.
- O livro é gerado mesmo que todas as datas das medidas manuais estejam vazias.
- Painel com um único espaçamento (~27 px no ecrã de referência): margens laterais, espaços entre cartões na horizontal e na vertical, antes dos gráficos e entre gráficos; largura total ajustada a uma janela de ~1858 px.
- Verificação no Excel mais robusta: filtro de mensagens COM (repete chamadas rejeitadas enquanto o Excel está ocupado), espera explícita até o Excel estar pronto, confirmação de que todas as folhas esperadas foram lidas e até 3 tentativas para falhas intermitentes da automação (nunca para diferenças de valores).

## 2.0.1 — 2026-10-11

- Painel alargado para ocupar um monitor largo (~1900 px a 100 %), com valores maiores nos cartões e margens laterais iguais aos espaços entre cartões.
- Linhas fixas: nas tabelas simples (Medidas Manuais, Medições Corporais, Critérios) só o cabeçalho fica fixo, sem linha vertical; nos resumos largos a coluna A continua fixa e o título, o subtítulo e a nota começam na coluna B, para a linha do Excel não os atravessar.
- Acentos corretos na janela do atualizador.
- Verificação/captura no Excel robustas quando já há outro Excel aberto.

## 2.0.0 — 2026-10-11

Revisão completa. As medidas manuais e todas as leituras existentes foram preservadas.

### Corrigido
- O bloqueio da atualização ficava preso enquanto a janela de erro estava aberta (a pausa acontecia antes de o libertar).
- As pastas temporárias nunca eram apagadas após sucesso e acumulavam cópias de dados pessoais.
- A verificação de "pesagens perdidas" não protegia leituras com altura diferente da altura de referência (precedência de `-and`/`-or` no PowerShell). Agora compara todos os instantes UTC.
- Duplicados detetados pela hora local: duas pesagens na mudança de hora de outubro bloqueariam a atualização. Agora usa o instante UTC.
- As medidas manuais eram lidas por posição de coluna; datas escritas como texto eram interpretadas no formato americano.
- Uma variação de cintura de exatamente 0,0 cm aparecia em branco no Resumo Semanal (diferença do motor de cálculo antigo face ao Excel).
- A massa proteica era descrita como valor da aplicação; é derivada pelo leitor.
- Ruído de vírgula flutuante no JSON.
- O CSV da exportação usava vírgula como separador e vírgula decimal ao mesmo tempo.

### Alterado
- Gerador reescrito em Python com **openpyxl** (MIT), substituindo uma biblioteca proprietária e os runtimes externos com caminhos fixos. Instalação reprodutível com `INSTALAR.cmd`.
- Cálculos feitos em Python (fonte única para o JSON) e **verificados no Microsoft Excel real** em cada atualização.
- Painel: médias de 7 dias, variação desde o início pela média dos primeiros/últimos 7 dias, ritmo das últimas 4 semanas, data da última leitura e da atualização.
- Gráficos: leituras diárias com média móvel de 7 dias; rótulos curtos; gráficos ancorados às células; tracejados por ponto, sem folha oculta.
- Resumos: colunas Início, Fim, Leituras, Dias com medição e Período parcial explícito; cintura e pescoço mensais.
- Critérios e Qualidade: protocolo de pesagem, limites de plausibilidade, contagens e lista de pesagens excluídas com motivo.
- JSON 2.0.0 com JSON Schema (ver [docs/JSON-PARA-IA.md](docs/JSON-PARA-IA.md)).
- Configuração pessoal em `config.json` (fora do Git).
- Estrutura do projeto reorganizada; scripts antigos arquivados.

### Acrescentado
- Testes automáticos com dados fictícios (pytest e testes PowerShell do atualizador com leitor simulado) e CI no GitHub Actions, incluindo compilação do leitor Go.
- Exemplo fictício e imagens de demonstração.
- Opção `-PermitirRemocoes` para aceitar pesagens apagadas de propósito na Zepp Life.

## 1.x — 2026-08 a 2026-10

Versões iniciais com gerador em Node.js e pós-processamento OOXML.
