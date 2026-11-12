# Arquitetura

```
├── ATUALIZAR-EXCEL.cmd, RECALCULAR-SEM-DESCARREGAR.cmd, ESQUECER-CONTA.cmd, INSTALAR.cmd
├── config.example.json          → copiar para config.json (privado)
├── scripts/
│   ├── atualizar-zepp.ps1        orquestração: bloqueio, autenticação, chamadas ao Python, substituição, histórico
│   ├── excel-com.ps1             apoio COM comum: filtro de mensagens (Excel ocupado), esperas, abertura, novas tentativas
│   ├── verificar-excel-real.ps1  abre uma cópia no Microsoft Excel (COM), recalcula e devolve valores
│   ├── capturar-folhas.ps1       imagens PNG de folhas (revisão de design / README)
│   ├── gerar_demo.py             exemplo fictício em docs/exemplo
│   └── comparar_saidas.py        compara dois pares Excel+JSON célula a célula (alterações estruturais)
├── zepp_life_excel_ia/           pacote Python (python -m zepp_life_excel_ia gerar | validar-exportacao)
│   ├── __main__.py               linha de comandos
│   ├── gerar.py                  gerar(): lê, calcula uma vez, escreve Excel, verifica, escreve JSON
│   ├── config.py                 leitura de config.json (inclui as opções sexo, ano_nascimento, objetivo e fases)
│   ├── metricas.py               registo único das métricas (JSON, colunas Excel, resumos, limites)
│   ├── exportacao.py             leitura e validação de uma nova descarga (duplicados, pesagens perdidas)
│   ├── leituras.py               conversão de fuso, seleção e motivos de exclusão
│   ├── medidas.py                medidas manuais: leitura por cabeçalho, CSV inicial, cópia de segurança, prioridades, validação
│   ├── resumos.py                cálculos (fonte única): Periodo, Indicador, Extremo, Composicao, ResumoFase, Qualidade, Calculos
│   ├── datas.py                  números de série de datas do Excel
│   ├── xlsx_bruto.py             leitura direta do OOXML (valores guardados, sem adivinhar formatos)
│   ├── excel/                    livro openpyxl, uma folha por módulo
│   │   ├── __init__.py           construir(): cria as folhas e devolve a lista de células esperadas
│   │   ├── estilos.py            cores, letra, títulos, cabeçalhos, tabelas, ignoredErrors
│   │   ├── graficos.py           gráficos diários (média móvel) e manuais (tracejados), eixos legíveis
│   │   ├── folha_painel.py       cartões e gráficos (medidas do ecrã de referência)
│   │   ├── folha_resumos.py      resumos semanal e mensal (cada coluna definida num só sítio)
│   │   ├── folha_fases.py        folha Fases (só quando config.json define fases)
│   │   ├── folha_manuais.py      medidas manuais editáveis
│   │   ├── folha_medicoes.py     leituras com massa gorda, massa magra, médias de 7 dias e a coluna auxiliar do ritmo
│   │   └── folha_criterios.py    dicionário, regras, qualidade dos dados, leituras atípicas e pesagens excluídas
│   ├── verificacao.py            corre o Excel real e compara a lista de células esperadas
│   ├── ia_json.py                JSON 2.1.0 e a versão compacta (derivada do completo)
│   └── sintetico.py              dados fictícios determinísticos
├── schema/                       JSON Schema do JSON para IA (completo e compacto)
├── leitor/                       leitor Go (derivado de SmartScaleConnect)
├── tests/                        pytest + testes PowerShell do atualizador (dados fictícios)
├── pyproject.toml                versão (lida de __init__.py), dependências, configuração do pytest e do ruff
└── docs/
```

Pastas pessoais (ignoradas pelo Git):

```
├── Excel e JSON/          resultados: analise-historico-zepp-life.xlsx, dados-zepp-life-ia.json, dados-zepp-life-ia-resumo.json
├── privado/
│   ├── medidas-manuais.csv  cópia de segurança das medidas manuais (atualizada em cada atualização)
│   ├── entradas/          CSV das medidas iniciais, exportação oficial completa da Zepp
│   ├── exportacao-zepp/   última descarga da API (JSON e CSV)
│   ├── historico/         versões anteriores, uma pasta por atualização (30 dias completos + a última de cada mês)
│   ├── temporario/        pasta de trabalho de cada atualização; guarda as 5 últimas tentativas falhadas (máx. 30 dias)
│   ├── .atualizacao.lock  bloqueio exclusivo durante uma atualização
│   ├── backups/           ZIPs de segurança com manifesto SHA-256
│   ├── arquivo/           código e tentativas da versão 1.x, material das sessões Codex
│   └── notas/             relatórios privados (dados reais)
├── bin/zepp-reader.exe    leitor compilado (+ zepp-reader.exe.sha256, o hash registado na primeira utilização)
├── .venv/                 ambiente Python
└── config.json            configuração pessoal
```

## Fluxo

1. `atualizar-zepp.ps1` obtém um bloqueio exclusivo (`privado/.atualizacao.lock`), confirma que o Excel está fechado e cria `privado/temporario/<data>_<id>/`.
2. O atualizador confirma que o SHA-256 do leitor é o registado (`Confirmar-Leitor`). O leitor Go recebe o pedido (token, conta ou código OAuth) **por stdin** e devolve `{token, weights}` por stdout. Nenhum segredo vai para argumentos, ficheiros ou registos.
3. `python -m zepp_life_excel_ia validar-exportacao` recusa uma descarga vazia, inválida ou com duplicados (código 1). Se o único problema forem pesagens anteriores em falta (por instante UTC), sai com código 2 e o atualizador pergunta ao utilizador antes de repetir com `--permitir-remocoes`.
4. `python -m zepp_life_excel_ia gerar` lê as medidas manuais do Excel atual (ou da cópia CSV, se o Excel não existir), calcula tudo em Python, escreve o livro com fórmulas nativas e verifica-o no Excel real; depois escreve o JSON com o SHA-256 do livro, a versão compacta e a nova cópia CSV das medidas manuais.
5. O atualizador confirma o hash, substitui os ficheiros com `File.Replace` (cópia anterior em `privado/historico/`) e, em caso de falha a meio, repõe os que já tinha substituído. No fim arruma o histórico (`Limpar-Historico`).

O PowerShell chama o Python sempre pela mesma função (`Invoke-Python`): grava toda a saída como texto em `validacao.log` / `gerador.log`, mostra as linhas de resumo e, se o comando falhar, lança a mensagem da linha `ERRO:`. No Windows PowerShell 5.1 não se redireciona o stderr nativo com `*>` (embrulha cada linha num registo de erro partido à largura da consola).
6. O bloqueio é libertado **antes** de qualquer pausa; a pasta temporária é apagada em caso de sucesso e guardada (máx. 5) em caso de falha.

## Decisões

- **openpyxl em vez de bibliotecas proprietárias**: licença MIT, instalação reprodutível, Excel nativo editável. Como o openpyxl não calcula fórmulas, os valores são calculados em Python e confirmados pelo Excel real.
- **Fórmulas vivas** nos resumos e no Painel para que edições manuais se reflitam de imediato; os dados corporais não devem ser editados à mão.
- **Tracejados por ponto (`dPt`)** com `dispBlanksAs=span`, em vez de séries auxiliares numa folha oculta.
- **`ignoredErrors`** acrescentado ao XML depois de gravar (o openpyxl não o escreve) para evitar avisos de "fórmula inconsistente" que são intencionais.
- **Endereços das células num só sítio**: cada folha, ao escrever uma fórmula, regista o valor que o Python espera (`Esperado`). A verificação percorre essa lista, por isso mudar o layout não exige mudar a verificação.
- **Registo único de métricas** (`metricas.py`): acrescentar ou corrigir uma métrica faz-se num só sítio; as folhas e o JSON escolhem quais mostrar e por que ordem.
- **Leitura OOXML direta** das medidas manuais, para preservar exatamente o que está guardado.
- **Margem de meio segundo nos critérios de data** (`estilos.depois_de`, `ate`, `antes_de`, `desde`): as funções `*IFS` convertem o critério em texto com 15 algarismos antes de comparar, o que podia trocar o lado de uma leitura exatamente no limite de uma janela (por exemplo, 7 dias certos antes de outra) e dar um valor diferente do Python. Como as leituras têm segundos inteiros, meio segundo de margem elimina o problema sem mudar a regra.
- **Valores auxiliares em colunas ocultas** em vez de fórmulas repetidas: o mínimo do peso e a sua data ficam em `Painel!S4:S5` (usados por cinco cartões) e os dias até à última leitura numa coluna oculta de Medições Corporais (o eixo X de `SLOPE` no ritmo). Todos têm valor esperado e são confirmados no Excel real.
- **Empates tratados da mesma forma nos dois lados**: o mínimo da média de 7 dias aceita valores até 1e-6 acima do mínimo e fica com a leitura mais antiga, no Python (`resumos.extremos`) e no Excel (`MINIFS`), para um empate numérico nunca dar datas diferentes.
- **JSON compacto derivado do completo** (`ia_json.construir_resumo` recebe o dicionário já construído): não há segundo caminho de cálculo que possa divergir.
- **Leituras atípicas só assinaladas**: excluir leituras mudaria as médias de forma pouco transparente; assinalar deixa a decisão a quem analisa.

## Manutenção

### Acrescentar ou corrigir uma métrica

1. Em `metricas.py`, declarar a `Metrica` (campo da exportação, chave JSON, nome, unidade, significado, origem; se aparecer no Excel, `coluna`, `formato` e `largura`; se aparecer nos resumos, `resumo`, `cabecalho_resumo` e `formato_resumo`).
2. Incluí-la em `COLUNAS_MEDICOES` (folha Medições Corporais) e/ou `RESUMO` (resumos), na posição desejada.
3. Atualizar `schema/dados-zepp-life-ia.schema.json` (chaves novas em `medicoes_corporais` e nos resumos) e, se o formato mudar, `versao_esquema` em `ia_json.py`.
4. Correr os testes: o teste no Excel real confirma automaticamente as novas células, porque cada folha regista os valores esperados.

### Acrescentar uma coluna aos resumos

Em `excel/folha_resumos.py`, acrescentar uma `Coluna` (cabeçalho, fórmula, formato, valor esperado, cor do grupo). As variações referem as colunas das médias pela letra e há uma verificação (`assert`) para a ordem não mudar sem aviso.

### Mudar o Painel

As medidas estão em constantes no início de `excel/folha_painel.py` (larguras, alturas, linhas dos cartões e dos gráficos), calculadas para um ecrã com ~1858 px úteis. Os cartões são `Cartao(rótulo, fórmula, formato, cor, nota, valor esperado)`. As notas são expressões Excel; para pôr um número numa nota use `estilos.numero_formula` (FIXED), não `TEXT()`, que depende do idioma do Excel. Critérios de data/hora em `*IFS` escrevem-se com `estilos.depois_de/ate/antes_de/desde`. Depois de mudar, ver o resultado no Excel real com `scripts/capturar-folhas.ps1` (correr a partir do PowerShell, não do Git Bash: usa a área de transferência).

### Acrescentar um indicador

1. Calculá-lo em `resumos.py` (e guardá-lo em `Calculos`, `Indicador` ou numa dataclass própria).
2. Mostrá-lo no Excel com uma fórmula e registar o `Esperado` com o valor do Python — se a fórmula e o Python divergirem, a atualização pára.
3. Acrescentá-lo a `ia_json.py` e ao esquema (`schema/`); uma chave nova sobe o número intermédio de `VERSAO_ESQUEMA`.
4. Cobrir com um teste em `tests/test_indicadores.py`, de preferência com `excel_real=True` para o caso-limite.

### Alterações estruturais (sem mudar o resultado)

1. Antes de mexer, gerar uma referência com o código atual (por exemplo com os dados fictícios: `sintetico.criar_projeto` e `python -m zepp_life_excel_ia gerar ... --sem-excel-real --hoje AAAA-MM-DD`).
2. Depois da alteração, gerar de novo com os mesmos dados e a mesma data.
3. `python scripts/comparar_saidas.py antes.xlsx antes.json depois.xlsx depois.json` tem de indicar 0 diferenças.

### Publicar uma versão

1. Subir `__version__` em `zepp_life_excel_ia/__init__.py` (o `pyproject.toml` lê-a daí) e descrever as alterações no `CHANGELOG.md`.
2. `ruff check .`, `pytest` e `tests\atualizador.Tests.ps1` sem falhas.
3. Se o aspeto mudou: `python scripts/gerar_demo.py` e capturar as imagens do README com `scripts/capturar-folhas.ps1`.

