# Utilização

## Atualizar com novas pesagens

1. Abra a Zepp Life no telemóvel e sincronize a balança.
2. Guarde e **feche** o Excel.
3. Execute `ATUALIZAR-EXCEL.cmd` e aguarde a mensagem verde.

O atualizador:

1. reutiliza a sessão guardada (ou autentica com a conta guardada);
2. descarrega **todas** as pesagens e recusa a descarga se estiver vazia, tiver pesagens inválidas ou duplicadas, ou **perder pesagens anteriores** (comparação por instante UTC);
3. gera o Excel e o JSON numa pasta temporária, preservando as medidas manuais do Excel atual;
4. abre o livro no Microsoft Excel (cópia só de leitura), recalcula e confirma que cada célula com fórmula relevante coincide com os cálculos em Python;
5. confirma que o hash do Excel coincide com o JSON;
6. substitui os ficheiros e guarda as versões anteriores em `privado\historico\<data-hora>\`.

Se algo falhar, nada é substituído (ou é reposto) e a janela indica o motivo e o caminho do registo. Ver [Registos e erros](#registos-e-erros).

## Recalcular sem descarregar

Depois de editar medidas manuais (ou quando quiser regenerar o livro), guarde e feche o Excel e execute `RECALCULAR-SEM-DESCARREGAR.cmd`. Usa a última exportação guardada e não contacta a Xiaomi. Também acrescenta as semanas e os meses que entretanto começaram.

## Opções do atualizador

Os três `.cmd` chamam `scripts\atualizar-zepp.ps1`. Para usar opções, execute-o diretamente:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\atualizar-zepp.ps1 -PermitirRemocoes
```

| Opção | Efeito |
|---|---|
| `-SoExcel` | Recalcular sem descarregar (o que faz `RECALCULAR-SEM-DESCARREGAR.cmd`). |
| `-PermitirRemocoes` | Aceitar uma descarga com menos pesagens do que a anterior (se as apagou de propósito na Zepp Life). |
| `-SemExcelReal` | Não abrir o Microsoft Excel para verificar (mais rápido; só para diagnóstico). |
| `-EsquecerConta` | Apagar a conta e a sessão guardadas (o que faz `ESQUECER-CONTA.cmd`). |
| `-SemPausa` | Não esperar por Enter no fim (útil em tarefas agendadas). |

## Medidas manuais

Na folha **Medidas Manuais**, preencha as células amarelas: data da medição, cintura (ao nível do umbigo), pescoço e observações. A data tem de pertencer à semana da coluna A (o Excel avisa). Números como `80,5` ou datas como `05/10/2026` escritos como texto também são aceites.

Para guardar as leituras individuais (por exemplo, três medições da cintura), escreva-as nas observações neste formato — são extraídas para o JSON:

```
Leituras — cintura: 82,5 - 82,6 - 82,5; pescoço: 37 - 37,1 - 37
```

Regras de preservação:

- os valores do Excel têm sempre prioridade sobre o CSV inicial, **incluindo campos apagados de propósito**;
- a chave é a semana (coluna A); as colunas são lidas pelo nome do cabeçalho, por isso a ordem pode mudar;
- não mude os nomes das folhas nem dos cabeçalhos de Medidas Manuais;
- formatação pessoal fora das células amarelas não é preservada (o livro é regenerado);
- uma semana com cintura ou pescoço tem de ter data; valores fora de 1–300 cm (cintura) ou 1–150 cm (pescoço) são recusados.

## O que contém o Excel

| Folha | Conteúdo |
|---|---|
| **Painel** | Última leitura, médias de 7 dias, variação desde o início, ritmo das últimas 4 semanas, cintura e pescoço; gráficos de peso, gordura, músculo, cintura e pescoço. Desenhado para um monitor largo a 100 %. |
| **Resumo Mensal** / **Resumo Semanal** | Leituras, dias com medição, período parcial, médias e variações; cintura e pescoço. A primeira coluna fica fixa ao deslocar para a direita. |
| **Medidas Manuais** | Registo semanal editável com variação e estado. |
| **Medições Corporais** | Todas as leituras aceites, massa gorda e médias móveis de 7 dias. Tabela com filtros (as colunas ocultas *Dia* e *Instante UTC* servem as fórmulas). |
| **Critérios e Qualidade** | Dicionário de campos, regras de seleção, contagens e pesagens excluídas com motivo. |

As fórmulas são nativas: ao editar medidas manuais, variações, estados e o Painel atualizam-se logo. As escalas dos gráficos e os tracejados só mudam depois de recalcular.

## Onde ficam os ficheiros

| Pasta | Conteúdo |
|---|---|
| `Excel e JSON\` | `analise-historico-zepp-life.xlsx` e `dados-zepp-life-ia.json` — os únicos ficheiros a abrir ou partilhar. |
| `privado\entradas\` | CSV das medidas iniciais e a exportação oficial completa da Zepp (referência). |
| `privado\exportacao-zepp\` | Última descarga da API (`zepp-completo.json` e `.csv`). |
| `privado\historico\` | Uma pasta por atualização com as versões anteriores. |
| `privado\temporario\` | Tentativas falhadas (as 5 mais recentes) com os registos. |
| `privado\backups\` | Cópias de segurança completas (ZIP + manifesto SHA-256). |

Tudo isto é pessoal e fica fora do Git. Como os backups estão dentro do projeto, copie `privado\backups\` de vez em quando para o OneDrive ou um disco externo.

## Registos e erros

Cada atualização trabalha em `privado\temporario\<data-hora>_<id>\`:

- `validacao.log` — validação da descarga (motivo da recusa, se houver);
- `gerador.log` — geração e verificação no Excel (inclui as diferenças, se o Excel calcular algo diferente do esperado).

Em caso de sucesso a pasta é apagada; em caso de falha fica guardada e a janela mostra o motivo e o caminho do registo.

## Partilhar com uma IA

Partilhe apenas `dados-zepp-life-ia.json` (contém dados corporais pessoais). É uma fotografia da última atualização: edições posteriores no Excel só entram depois de `RECALCULAR-SEM-DESCARREGAR.cmd`. Ver [JSON-PARA-IA.md](JSON-PARA-IA.md).
