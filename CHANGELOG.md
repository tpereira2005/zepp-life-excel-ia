# Registo de alterações

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
