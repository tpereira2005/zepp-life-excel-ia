# Instalação

## Requisitos

- Windows 10 ou 11 com Windows PowerShell 5.1 (incluído no Windows).
- [Python 3.11 ou superior](https://www.python.org/downloads/) (de preferência com o lançador `py`).
- Microsoft Excel (recomendado). Sem Excel, o livro é gerado na mesma, mas a verificação no Excel real é ignorada com um aviso.
- Conta Xiaomi associada à Zepp Life (não apenas login Google/Apple).

## Passos

1. Clonar ou descarregar o repositório.
2. Executar `INSTALAR.cmd`. Cria `.venv\` com as dependências fixadas em `requirements.txt` (openpyxl, tzdata) e copia `config.example.json` para `config.json` se ainda não existir.
3. Editar `config.json`:

   | Campo | Significado |
   |---|---|
   | `inicio_transformacao` | Primeiro dia (AAAA-MM-DD) a considerar. |
   | `fuso_horario` | Fuso IANA das datas locais (por omissão `Europe/Lisbon`). |
   | `perfil_titular_id` | Membro Zepp lido (`-1` = titular da conta). |
   | `altura_perfil_sem_id_cm` | Só para exportações antigas sem identificador de perfil; `null` desativa. |
   | `protocolo_pesagem` | Texto livre que descreve como se pesa (aparece no Excel e no JSON). |
   | `csv_medidas_iniciais` | CSV opcional com medidas manuais antigas, relativo à pasta do projeto (por omissão `privado/entradas/medicoes-cintura-pescoco.csv`; formato no [exemplo](exemplo/exemplo-medidas-manuais.csv)). Só é usado para semanas que ainda não existem no Excel. |
   | `sexo` *(opcional)* | `"masculino"` ou `"feminino"`. Com `"masculino"` é calculada a *gordura pela fita* (método da Marinha dos EUA); a fórmula feminina exige a anca, que não é registada. Vai para o `contexto` do JSON. |
   | `ano_nascimento` *(opcional)* | Ano (ex.: `1990`); dá contexto à IA (idade aproximada). |
   | `objetivo` *(opcional)* | Texto livre; aparece no `contexto` do JSON. |
   | `fases` *(opcional)* | Lista de fases por ordem de data: `{"nome": "Défice", "inicio": "2026-01-04", "objetivo": "…"}`. Cada fase termina na véspera da seguinte. Cria a folha *Fases* e a chave `fases` do JSON. Exemplo em [`exemplo/exemplo-config.json`](exemplo/exemplo-config.json). |

4. Obter o leitor `bin\zepp-reader.exe`:
   - **Descarregar:** em *Actions → CI → execução mais recente → Artifacts → zepp-reader-windows-amd64* (é preciso ter sessão iniciada no GitHub). O artefacto traz `zepp-reader.exe.sha256` e o resumo da execução mostra o mesmo valor e o commit de origem.
   - **Compilar** (Go 1.24+):

     ```bat
     cd leitor
     set GOOS=windows& set GOARCH=amd64& set CGO_ENABLED=0
     go build -trimpath -ldflags "-s -w" -o ..\bin\zepp-reader.exe .\cmd\zepp-reader
     ```

5. Executar `ATUALIZAR-EXCEL.cmd`. Na primeira vez pede a conta e a palavra-passe Xiaomi e pergunta se quer guardar a palavra-passe.

### Confirmar o leitor

O leitor é o único programa que recebe a palavra-passe. Na primeira utilização o atualizador mostra o SHA-256 do executável e guarda-o em `bin\zepp-reader.exe.sha256`; a partir daí recusa qualquer executável diferente, antes de lhe enviar o que quer que seja. Compare o valor mostrado com o publicado pelo CI:

```powershell
(Get-FileHash bin\zepp-reader.exe -Algorithm SHA256).Hash.ToLower()
```

Se substituir o leitor de propósito (nova versão), apague `bin\zepp-reader.exe.sha256` e volte a confirmar o valor.

O resultado fica em `Excel e JSON\`; a exportação, o histórico e as tentativas falhadas ficam em `privado\` (ver [UTILIZACAO.md](UTILIZACAO.md#onde-ficam-os-ficheiros)).

## Primeira autenticação e verificação Xiaomi

A Xiaomi pode pedir uma verificação adicional (código por email). O atualizador abre o navegador; conclua a verificação e, na página final (mesmo que mostre erro 404), copie o endereço com `Ctrl+L`, `Ctrl+C`, regresse à janela e prima Enter. Não partilhe esse endereço.

A sessão fica sempre em `%LOCALAPPDATA%\ZeppLifeExcel`, cifrada com DPAPI para o seu utilizador Windows e numa pasta acessível só a esse utilizador. A palavra-passe só é guardada (da mesma forma) se responder `s` quando o atualizador pergunta:

- **`s`** — quando a sessão expirar, é renovada sozinha. Qualquer programa a correr com o seu utilizador Windows consegue, em teoria, decifrá-la.
- **`n`** — nada além da sessão fica guardado; quando esta expirar, a conta e a palavra-passe voltam a ser pedidas.

`ESQUECER-CONTA.cmd` apaga a sessão e a palavra-passe (não revoga sessões no servidor). Para apagar só a palavra-passe: `scripts\atualizar-zepp.ps1 -EsquecerPalavraPasse`.

## Atualizar o programa

```bat
git pull
INSTALAR.cmd
```

`INSTALAR.cmd` volta a instalar as dependências fixadas e não toca em `config.json`, `Excel e JSON\` nem `privado\`. As medidas manuais são preservadas na atualização seguinte do Excel. Ver as novidades em [CHANGELOG.md](../CHANGELOG.md).

## Ferramentas de desenvolvimento (opcional)

```bat
.venv\Scripts\python -m pip install -r requirements-dev.txt
```

Instala pytest, jsonschema e ruff, com as versões indicadas também em `pyproject.toml`. Ver a secção *Desenvolvimento* do [README](../README.md#desenvolvimento) e [ARQUITETURA.md](ARQUITETURA.md#manutenção).

## Mover para outro computador

Copie a pasta do projeto **incluindo** `config.json`, `Excel e JSON\`, `privado\` e `bin\` (dados pessoais e leitor), execute `INSTALAR.cmd` e volte a autenticar-se: a cifra DPAPI não é transferível entre computadores ou utilizadores.

## Problemas comuns

| Mensagem | O que fazer |
|---|---|
| *O ambiente Python não está instalado* | Executar `INSTALAR.cmd`. |
| *Falta config.json* | Copiar `config.example.json` para `config.json` e ajustar. |
| *Leitor Zepp não encontrado* | Colocar `zepp-reader.exe` em `bin\` (passo 4). |
| *O leitor Zepp foi alterado desde a última utilização* | O executável não é o que foi registado. Se o substituiu de propósito, apague `bin\zepp-reader.exe.sha256`; caso contrário, volte a descarregá-lo do CI. |
| *Guarde e feche o Excel antes de atualizar* | Fechar o livro em `Excel e JSON\` (o atualizador precisa de o substituir). |
| *Já existe uma atualização em curso* | Fechar a outra janela de atualização. |
| *A Xiaomi rejeitou a conta ou palavra-passe (código 70016)* | Confirmar que é uma Conta Xiaomi; após mudar a palavra-passe, executar `ESQUECER-CONTA.cmd`. |
