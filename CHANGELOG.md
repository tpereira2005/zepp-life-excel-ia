# Registo de alterações

## 2.0.1 — 2026-10-11

- Painel alargado para ocupar um monitor largo (~1900 px a 100 %), com valores maiores nos cartões.
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
