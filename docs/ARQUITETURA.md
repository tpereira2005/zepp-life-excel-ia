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
│   ├── config.py                 leitura de config.json
│   ├── metricas.py               registo único das métricas (JSON, colunas Excel, resumos, limites)
│   ├── exportacao.py             leitura e validação de uma nova descarga (duplicados, pesagens perdidas)
│   ├── leituras.py               conversão de fuso, seleção e motivos de exclusão
│   ├── medidas.py                medidas manuais: leitura por cabeçalho, CSV inicial, prioridades, validação
│   ├── resumos.py                cálculos (fonte única): Periodo, Indicador, Calculos
│   ├── datas.py                  números de série de datas do Excel
│   ├── xlsx_bruto.py             leitura direta do OOXML (valores guardados, sem adivinhar formatos)
│   ├── excel/                    livro openpyxl, uma folha por módulo
│   │   ├── __init__.py           construir(): cria as folhas e devolve a lista de células esperadas
│   │   ├── estilos.py            cores, letra, títulos, cabeçalhos, tabelas, ignoredErrors
│   │   ├── graficos.py           gráficos diários (média móvel) e manuais (tracejados), eixos legíveis
│   │   ├── folha_painel.py       cartões e gráficos (medidas do ecrã de referência)
│   │   ├── folha_resumos.py      resumos semanal e mensal (cada coluna definida num só sítio)
│   │   ├── folha_manuais.py      medidas manuais editáveis
│   │   ├── folha_medicoes.py     leituras com massa gorda e médias de 7 dias
│   │   └── folha_criterios.py    dicionário, regras e pesagens excluídas
│   ├── verificacao.py            corre o Excel real e compara a lista de células esperadas
│   ├── ia_json.py                JSON 2.0.0
│   └── sintetico.py              dados fictícios determinísticos
├── schema/                       JSON Schema do JSON para IA
├── leitor/                       leitor Go (derivado de SmartScaleConnect)
├── tests/                        pytest + testes PowerShell do atualizador (dados fictícios)
├── pyproject.toml                versão (lida de __init__.py), dependências, configuração do pytest e do ruff
└── docs/
```

Pastas pessoais (ignoradas pelo Git):

```
├── Excel e JSON/          resultados: analise-historico-zepp-life.xlsx, dados-zepp-life-ia.json
├── privado/
│   ├── entradas/          CSV das medidas iniciais, exportação oficial completa da Zepp
│   ├── exportacao-zepp/   última descarga da API (JSON e CSV)
│   ├── historico/         versões anteriores do Excel/JSON/exportação, uma pasta por atualização
│   ├── temporario/        pasta de trabalho de cada atualização; guarda as 5 últimas tentativas falhadas
│   ├── .atualizacao.lock  bloqueio exclusivo durante uma atualização
│   ├── backups/           ZIPs de segurança com manifesto SHA-256
│   ├── arquivo/           código e tentativas da versão 1.x, material das sessões Codex
│   └── notas/             relatórios privados (dados reais)
├── bin/zepp-reader.exe    leitor compilado
├── .venv/                 ambiente Python
└── config.json            configuração pessoal
```

## Fluxo

1. `atualizar-zepp.ps1` obtém um bloqueio exclusivo (`privado/.atualizacao.lock`), confirma que o Excel está fechado e cria `privado/temporario/<data>_<id>/`.
2. O leitor Go recebe o pedido (token, conta ou código OAuth) **por stdin** e devolve `{token, weights}` por stdout. Nenhum segredo vai para argumentos, ficheiros ou registos.
3. `python -m zepp_life_excel_ia validar-exportacao` recusa uma descarga vazia, inválida, com duplicados ou que perca pesagens anteriores (por instante UTC).
4. `python -m zepp_life_excel_ia gerar` lê as medidas manuais do Excel atual, calcula tudo em Python, escreve o livro com fórmulas nativas e verifica-o no Excel real; depois escreve o JSON com o SHA-256 do livro.
5. O atualizador confirma o hash, substitui os ficheiros com `File.Replace` (cópia anterior em `privado/historico/`) e, em caso de falha, repõe os anteriores.

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

## Manutenção

### Acrescentar ou corrigir uma métrica

1. Em `metricas.py`, declarar a `Metrica` (campo da exportação, chave JSON, nome, unidade, significado, origem; se aparecer no Excel, `coluna`, `formato` e `largura`; se aparecer nos resumos, `resumo`, `cabecalho_resumo` e `formato_resumo`).
2. Incluí-la em `COLUNAS_MEDICOES` (folha Medições Corporais) e/ou `RESUMO` (resumos), na posição desejada.
3. Atualizar `schema/dados-zepp-life-ia.schema.json` (chaves novas em `medicoes_corporais` e nos resumos) e, se o formato mudar, `versao_esquema` em `ia_json.py`.
4. Correr os testes: o teste no Excel real confirma automaticamente as novas células, porque cada folha regista os valores esperados.

### Acrescentar uma coluna aos resumos

Em `excel/folha_resumos.py`, acrescentar uma `Coluna` (cabeçalho, fórmula, formato, valor esperado, cor do grupo). As variações referem as colunas das médias pela letra e há uma verificação (`assert`) para a ordem não mudar sem aviso.

### Mudar o Painel

As medidas estão em constantes no início de `excel/folha_painel.py` (larguras, alturas, linhas dos cartões e dos gráficos), calculadas para um ecrã com ~1858 px úteis. Os cartões são `Cartao(rótulo, fórmula, formato, cor, nota, valor esperado)`. Depois de mudar, ver o resultado no Excel real com `scripts/capturar-folhas.ps1` (correr a partir do PowerShell, não do Git Bash: usa a área de transferência).

### Alterações estruturais (sem mudar o resultado)

1. Antes de mexer, gerar uma referência com o código atual (por exemplo com os dados fictícios: `sintetico.criar_projeto` e `python -m zepp_life_excel_ia gerar ... --sem-excel-real --hoje AAAA-MM-DD`).
2. Depois da alteração, gerar de novo com os mesmos dados e a mesma data.
3. `python scripts/comparar_saidas.py antes.xlsx antes.json depois.xlsx depois.json` tem de indicar 0 diferenças.

### Publicar uma versão

1. Subir `__version__` em `zepp_life_excel_ia/__init__.py` (o `pyproject.toml` lê-a daí) e descrever as alterações no `CHANGELOG.md`.
2. `ruff check .`, `pytest` e `tests\atualizador.Tests.ps1` sem falhas.
3. Se o aspeto mudou: `python scripts/gerar_demo.py` e capturar as imagens do README com `scripts/capturar-folhas.ps1`.

