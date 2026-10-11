# Utilização

## Atualizar com novas pesagens

1. Abra a Zepp Life no telemóvel e sincronize a balança.
2. Guarde e **feche** o Excel.
3. Execute `ATUALIZAR-EXCEL.cmd` e aguarde a mensagem verde.

O atualizador:

1. reutiliza a sessão guardada (ou autentica com a conta guardada);
2. descarrega **todas** as pesagens e recusa a exportação se estiver vazia, tiver duplicados ou **perder pesagens anteriores** (comparação por instante UTC);
3. gera o Excel e o JSON numa pasta temporária, preservando as medidas manuais do Excel atual;
4. abre o livro no Microsoft Excel (cópia só de leitura), recalcula e confirma que todos os valores coincidem com os cálculos em Python;
5. confirma que o hash do Excel coincide com o JSON;
6. substitui os ficheiros e guarda as versões anteriores em `outputs\zepp-life\historico\<data-hora>\`.

Se algo falhar, nada é substituído (ou é reposto) e a janela indica o motivo e o registo em `outputs\zepp-life\temporario\`. São guardadas as 5 tentativas falhadas mais recentes.

Se apagou pesagens de propósito na Zepp Life, a proteção contra perdas vai interromper a atualização. Nesse caso execute:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\atualizar-zepp.ps1 -PermitirRemocoes
```

## Medidas manuais

Na folha **Medidas Manuais**, preencha as células amarelas: data da medição, cintura (ao nível do umbigo), pescoço e observações. A data tem de pertencer à semana da coluna A (o Excel avisa). Números como `80,5` ou datas como `05/10/2026` escritos como texto também são aceites.

Depois, guarde, feche e execute `RECALCULAR-SEM-DESCARREGAR.cmd` para atualizar o JSON e os gráficos (escalas e tracejados) sem contactar a Xiaomi.

Regras de preservação:

- os valores do Excel têm sempre prioridade sobre o CSV inicial, **incluindo campos apagados de propósito**;
- a chave é a semana (coluna A); as colunas são lidas pelo nome do cabeçalho;
- não mude os nomes das folhas nem dos cabeçalhos de Medidas Manuais;
- formatação pessoal fora das células amarelas não é preservada (o livro é regenerado).

## O que contém o Excel

| Folha | Conteúdo |
|---|---|
| **Painel** | Última leitura, médias de 7 dias, variação desde o início, ritmo das últimas 4 semanas, cintura e pescoço; gráficos. |
| **Resumo Mensal** / **Resumo Semanal** | Leituras, dias com medição, período parcial, médias e variações; cintura e pescoço. |
| **Medidas Manuais** | Registo semanal editável com variação e estado. |
| **Medições Corporais** | Todas as leituras aceites, massa gorda e médias móveis de 7 dias. Tabela com filtros. |
| **Critérios e Qualidade** | Dicionário de campos, regras de seleção, contagens e pesagens excluídas com motivo. |

## Partilhar com uma IA

Partilhe apenas `dados-zepp-life-ia.json` (contém dados corporais pessoais). É uma fotografia da última atualização: edições posteriores no Excel só entram depois de `RECALCULAR-SEM-DESCARREGAR.cmd`. Ver [JSON-PARA-IA.md](JSON-PARA-IA.md).
