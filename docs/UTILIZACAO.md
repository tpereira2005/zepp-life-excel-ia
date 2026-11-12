# Utilização

## Atualizar com novas pesagens

1. Abra a Zepp Life no telemóvel e sincronize a balança.
2. Guarde e **feche** o Excel.
3. Execute `ATUALIZAR-EXCEL.cmd` e aguarde a mensagem verde.

O atualizador:

1. reutiliza a sessão guardada (ou autentica com a conta guardada);
2. descarrega **todas** as pesagens e recusa a descarga se estiver vazia ou tiver pesagens inválidas ou duplicadas; se **faltarem pesagens anteriores** (comparação por instante UTC), mostra quais e só continua se escrever `SIM`;
3. gera o Excel e os dois JSON numa pasta temporária, preservando as medidas manuais do Excel atual;
4. abre o livro no Microsoft Excel (cópia só de leitura), recalcula e confirma que cada célula com fórmula relevante coincide com os cálculos em Python;
5. confirma que o hash do Excel coincide com o JSON;
6. substitui os ficheiros, guarda as versões anteriores em `privado\historico\<data-hora>\` e atualiza a cópia das medidas manuais (`privado\medidas-manuais.csv`);
7. arruma o histórico: ficam todas as cópias dos últimos 30 dias e, das mais antigas, a última de cada mês.

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
| `-PermitirRemocoes` | Aceitar sem perguntar uma descarga com menos pesagens do que a anterior (se as apagou de propósito na Zepp Life). |
| `-ManterHistoricoDias N` | Dias em que todas as cópias do histórico são mantidas (por omissão 30); `0` nunca apaga. |
| `-EsquecerPalavraPasse` | Apagar só a palavra-passe guardada; a sessão continua até expirar. |
| `-SemExcelReal` | Não abrir o Microsoft Excel para verificar (mais rápido; só para diagnóstico). |
| `-EsquecerConta` | Apagar a conta e a sessão guardadas (o que faz `ESQUECER-CONTA.cmd`). |
| `-SemPausa` | Não esperar por Enter no fim nem fazer perguntas (útil em tarefas agendadas): pesagens em falta são recusadas. |

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

### Cópia de segurança das medidas manuais

As medidas manuais são os únicos dados que não se podem voltar a descarregar. Em cada atualização ficam também em `privado\medidas-manuais.csv` (texto simples: semana, data, cintura, pescoço, observações). Se o Excel for apagado ou perdido, a atualização seguinte repõe as medidas a partir dessa cópia e avisa. Enquanto o Excel existir, é sempre ele que conta. Se o Excel existir mas estiver ilegível, a atualização pára sem substituir nada: recupere uma versão de `privado\historico\`.

## O que contém o Excel

| Folha | Conteúdo |
|---|---|
| **Painel** | 16 cartões em 4 filas e 5 gráficos (ver abaixo). Desenhado para um monitor largo a 100 %. |
| **Resumo Mensal** / **Resumo Semanal** | Leituras, dias com medição, período parcial, médias e variações (incluindo massa gorda e massa magra); cintura e pescoço; no semanal, cintura ÷ altura e, se configurado, gordura pela fita. A primeira coluna fica fixa ao deslocar para a direita. |
| **Fases** *(se definidas em `config.json`)* | Uma linha por fase: médias de 7 dias no início e no fim, variações de peso, gordura, massa gorda, massa magra e cintura, kg/semana e objetivo. |
| **Medidas Manuais** | Registo semanal editável com variação e estado. |
| **Medições Corporais** | Todas as leituras aceites, massa gorda, massa magra e médias móveis de 7 dias. Tabela com filtros (três colunas ocultas servem as fórmulas: *Dias até à última*, *Dia* e *Instante UTC*). |
| **Critérios e Qualidade** | Dicionário de campos, regras, qualidade dos dados, leituras atípicas (assinaladas, não excluídas) e pesagens excluídas com motivo. |

### Cartões do Painel

| Fila | Cartões |
|---|---|
| Peso | Último peso · média de 7 dias · **desde o início** · **desde o mínimo** (a nota mostra o valor e a data do mínimo) |
| Tendência e balança | **Ritmo recente** (tendência dos últimos 28 dias, kg/semana) · gordura (média de 7 dias) · gordura desde o início · músculo (média de 7 dias; a nota mostra a variação desde o início) |
| Composição | Massa gorda e massa magra **desde o início** e **desde o mínimo do peso**; a nota diz que percentagem da variação do peso foi gordura |
| Fita métrica | Última cintura · desde o início · **desde o mínimo** · cintura ÷ altura (ou *gordura pela fita*, se `sexo` for `masculino`) |

"Desde o mínimo" responde à pergunta que "desde o início" esconde: quanto se recuperou desde o ponto mais baixo. Se o peso ainda está a descer, o mínimo é o valor atual e o cartão mostra cerca de zero.

O subtítulo fica **vermelho** quando a última leitura tem mais de 3 dias: os valores "atuais" referem-se a essa leitura, não ao dia de hoje.

As fórmulas são nativas: ao editar medidas manuais, variações, estados e o Painel atualizam-se logo. As escalas dos gráficos e os tracejados só mudam depois de recalcular.

## Onde ficam os ficheiros

| Pasta | Conteúdo |
|---|---|
| `Excel e JSON\` | `analise-historico-zepp-life.xlsx`, `dados-zepp-life-ia.json` e `dados-zepp-life-ia-resumo.json` — os únicos ficheiros a abrir ou partilhar. |
| `privado\medidas-manuais.csv` | Cópia de segurança das medidas manuais, atualizada em cada atualização. |
| `privado\entradas\` | CSV das medidas iniciais e a exportação oficial completa da Zepp (referência). |
| `privado\exportacao-zepp\` | Última descarga da API (`zepp-completo.json` e `.csv`). |
| `privado\historico\` | Uma pasta por atualização com as versões anteriores: todas as dos últimos 30 dias e a última de cada mês anterior. |
| `privado\temporario\` | Tentativas falhadas (as 5 mais recentes, nunca com mais de 30 dias) com os registos. |
| `privado\backups\` | Cópias de segurança completas (ZIP + manifesto SHA-256). |

Tudo isto é pessoal e fica fora do Git. Como os backups estão dentro do projeto, copie `privado\backups\` de vez em quando para o OneDrive ou um disco externo.

## Registos e erros

Cada atualização trabalha em `privado\temporario\<data-hora>_<id>\`:

- `validacao.log` — validação da descarga (motivo da recusa, se houver);
- `gerador.log` — geração e verificação no Excel (inclui as diferenças, se o Excel calcular algo diferente do esperado).

Em caso de sucesso a pasta é apagada; em caso de falha fica guardada e a janela mostra o motivo e o caminho do registo.

## Partilhar com uma IA

Partilhe `dados-zepp-life-ia-resumo.json` (cabe em qualquer IA: indicadores, fases, resumos e as leituras dos últimos 28 dias) ou, para uma análise leitura a leitura, `dados-zepp-life-ia.json`. Ambos contêm dados corporais pessoais. São uma fotografia da última atualização: edições posteriores no Excel só entram depois de `RECALCULAR-SEM-DESCARREGAR.cmd`. Ver [JSON-PARA-IA.md](JSON-PARA-IA.md).
