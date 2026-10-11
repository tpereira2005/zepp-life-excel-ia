# Leitor Zepp (`zepp-reader`)

Derivado de [SmartScaleConnect](https://github.com/AlexxIT/SmartScaleConnect) (Alexey Khit, MIT — ver [LICENSE](LICENSE)). Apenas `cmd/zepp-reader` é usado por este projeto; o resto do módulo original é mantido para compilar sem alterações.

## Contrato

- **Entrada (stdin, JSON):** um de `{"Token": "utilizador:apptoken"}`, `{"Account": "...", "Password": "..."}` ou `{"Code": "<código OAuth>"}`.
- **Saída (stdout, JSON):** `{"token": "...", "weights": [ {Date, Weight, BMI, BodyFat, ..., ProfileID} ]}`.
- **Erros:** mensagem em stderr e código 1. Se a Xiaomi exigir verificação, stderr contém `XIAOMI_VERIFICATION_URL=<url>`.
- Só leitura: nunca escreve nem apaga pesagens.

## Alterações locais face ao original

- `cmd/zepp-reader`: exportação só de leitura com segredos por stdin/stdout.
- `pkg/core/weight.go`: campos `ProteinRatio`, `StandBodyWeight`, `Impedance`, `EncryptImpedance`, `ProfileID`.
- `pkg/zepp`: `ProfileID` = membro pedido; erro em HTTP ≠ 200; deteção de paginação parada; erros de autenticação explícitos.
- `pkg/xiaomi/auth.go`: deteção de verificação adicional, CAPTCHA e credenciais rejeitadas.

## Limitações conhecidas

- `VisceralFat` e `BasalMetabolism` são convertidos para inteiro (truncados) pelo tipo do modelo original.
- `ProteinMass` é derivada (peso × proteína / 100), não devolvida pela API.
- Os corpos das respostas HTTP de cada página só são fechados no fim da exportação (impacto desprezável: poucas páginas).

## Compilar

```bat
set GOOS=windows& set GOARCH=amd64& set CGO_ENABLED=0
go build -trimpath -ldflags "-s -w" -o ..\bin\zepp-reader.exe .\cmd\zepp-reader
```
