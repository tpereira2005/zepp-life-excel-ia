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

4. Obter o leitor `bin\zepp-reader.exe`:
   - **Descarregar:** em *Actions → CI → execução mais recente → Artifacts → zepp-reader-windows-amd64* (é preciso ter sessão iniciada no GitHub).
   - **Compilar** (Go 1.24+):

     ```bat
     cd leitor
     set GOOS=windows& set GOARCH=amd64& set CGO_ENABLED=0
     go build -trimpath -ldflags "-s -w" -o ..\bin\zepp-reader.exe .\cmd\zepp-reader
     ```

5. Executar `ATUALIZAR-EXCEL.cmd`. Na primeira vez pede a conta e a palavra-passe Xiaomi.

O resultado fica em `Excel e JSON\`; a exportação, o histórico e as tentativas falhadas ficam em `privado\` (ver [UTILIZACAO.md](UTILIZACAO.md#onde-ficam-os-ficheiros)).

## Primeira autenticação e verificação Xiaomi

A Xiaomi pode pedir uma verificação adicional (código por email). O atualizador abre o navegador; conclua a verificação e, na página final (mesmo que mostre erro 404), copie o endereço com `Ctrl+L`, `Ctrl+C`, regresse à janela e prima Enter. Não partilhe esse endereço.

A conta e a sessão ficam em `%LOCALAPPDATA%\ZeppLifeExcel`, cifradas com DPAPI para o seu utilizador Windows e numa pasta acessível só a esse utilizador. `ESQUECER-CONTA.cmd` apaga-as (não revoga sessões no servidor).

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
| *Guarde e feche o Excel antes de atualizar* | Fechar o livro em `Excel e JSON\` (o atualizador precisa de o substituir). |
| *Já existe uma atualização em curso* | Fechar a outra janela de atualização. |
| *A Xiaomi rejeitou a conta ou palavra-passe (código 70016)* | Confirmar que é uma Conta Xiaomi; após mudar a palavra-passe, executar `ESQUECER-CONTA.cmd`. |
