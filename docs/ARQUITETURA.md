# Arquitetura

```
├── ATUALIZAR-EXCEL.cmd, RECALCULAR-SEM-DESCARREGAR.cmd, ESQUECER-CONTA.cmd, INSTALAR.cmd
├── config.example.json          → copiar para config.json (privado)
├── scripts/
│   ├── atualizar-zepp.ps1        orquestração: bloqueio, autenticação, validação, substituição, histórico
│   ├── verificar-excel-real.ps1  abre uma cópia no Microsoft Excel (COM), recalcula e devolve valores
│   ├── capturar-folhas.ps1       imagens PNG de folhas (revisão de design / README)
│   └── gerar_demo.py             exemplo fictício em docs/exemplo
├── zepp_transformacao/           pacote Python
│   ├── config.py                 leitura de config.json
│   ├── leituras.py               conversão de fuso, seleção e motivos de exclusão
│   ├── medidas.py                medidas manuais: leitura por cabeçalho, CSV inicial, prioridades, validação
│   ├── xlsx_bruto.py             leitura direta do OOXML (valores guardados, sem adivinhar formatos)
│   ├── resumos.py                cálculos (fonte única): resumos, médias móveis, indicadores
│   ├── excel.py                  livro openpyxl: tabelas, fórmulas vivas, formatação, gráficos
│   ├── verificacao.py            compara o Excel real com os cálculos Python
│   ├── ia_json.py                JSON 2.0.0
│   ├── sintetico.py              dados fictícios determinísticos
│   └── gerar.py                  ponto de entrada (CLI)
├── schema/                       JSON Schema do JSON para IA
├── leitor/                       leitor Go (derivado de SmartScaleConnect)
├── tests/                        pytest + testes PowerShell do atualizador
└── docs/
```

Pastas privadas (ignoradas pelo Git): `inputs/`, `outputs/`, `arquivo/`, `docs/privado/`, `bin/`, `.venv/`, `config.json`.

## Fluxo

1. `atualizar-zepp.ps1` obtém um bloqueio exclusivo (`outputs/zepp-life/.atualizacao.lock`), confirma que o Excel está fechado e cria `temporario/<data>_<id>/`.
2. O leitor Go recebe o pedido (token, conta ou código OAuth) **por stdin** e devolve `{token, weights}` por stdout. Nenhum segredo vai para argumentos, ficheiros ou registos.
3. A exportação é validada (vazia, duplicados, pesagens perdidas por instante UTC).
4. `python -m zepp_transformacao.gerar` lê as medidas manuais do Excel atual, calcula tudo em Python, escreve o livro com fórmulas nativas e verifica-o no Excel real; depois escreve o JSON com o SHA-256 do livro.
5. O atualizador confirma o hash, substitui os ficheiros com `File.Replace` (cópia anterior em `historico/`) e, em caso de falha, repõe os anteriores.
6. O bloqueio é libertado **antes** de qualquer pausa; a pasta temporária é apagada em caso de sucesso e guardada (máx. 5) em caso de falha.

## Decisões

- **openpyxl em vez de bibliotecas proprietárias**: licença MIT, instalação reprodutível, Excel nativo editável. Como o openpyxl não calcula fórmulas, os valores são calculados em Python e confirmados pelo Excel real.
- **Fórmulas vivas** nos resumos e no Painel para que edições manuais se reflitam de imediato; os dados corporais não devem ser editados à mão.
- **Tracejados por ponto (`dPt`)** com `dispBlanksAs=span`, em vez de séries auxiliares numa folha oculta.
- **`ignoredErrors`** acrescentado ao XML depois de gravar (o openpyxl não o escreve) para evitar avisos de "fórmula inconsistente" que são intencionais.
- **Leitura OOXML direta** das medidas manuais, para preservar exatamente o que está guardado.
