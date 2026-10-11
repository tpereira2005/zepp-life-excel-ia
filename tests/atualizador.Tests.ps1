<#
  Testes do atualizador com um leitor Zepp falso e pastas temporárias.
  Nunca usa a pasta privada real nem os dados reais. Executar com:
    powershell -NoProfile -ExecutionPolicy Bypass -File tests\atualizador.Tests.ps1
#>
$ErrorActionPreference = 'Stop'
$raiz = Split-Path $PSScriptRoot -Parent
$python = Join-Path $raiz '.venv\Scripts\python.exe'
$script = Join-Path $raiz 'scripts\atualizar-zepp.ps1'
$t = Join-Path ([IO.Path]::GetTempPath()) ('zepp-teste-' + [guid]::NewGuid().ToString('N'))
$dados = Join-Path $t 'dados'; $resultados = Join-Path $t 'resultados'; $privado = Join-Path $t 'credenciais'
[void](New-Item -ItemType Directory -Path $t)
$falhas = 0
function Confirmar([bool]$condicao, [string]$descricao) {
    if ($condicao) { Write-Host "  ok  $descricao" -ForegroundColor Green }
    else { Write-Host "  FALHOU  $descricao" -ForegroundColor Red; $script:falhas++ }
}
function Resposta([int]$dias, [int]$remover = 0) {
    $saida = Join-Path $t 'resposta.json'
    & $python -c "import json,sys; sys.path.insert(0, r'$raiz'); from zepp_life_excel_ia import sintetico; w=sintetico.exportacao($dias)[${remover}:]; json.dump({'token':'utilizador:token-de-teste','weights':w}, open(r'$saida','w'))"
    if ($LASTEXITCODE -ne 0) { throw 'Falha ao criar a resposta de teste.' }
}
function Atualizar([string[]]$extra) {
    $saida = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script -SemPausa -SemExcelReal -Hoje 2026-05-10 `
        -PastaDados $dados -PastaResultados $resultados -PastaPrivada $privado -Leitor (Join-Path $t 'leitor.cmd') -Config (Join-Path $t 'config.json') @extra 2>&1
    return @{ Codigo = $LASTEXITCODE; Texto = ($saida | Out-String) }
}
function Hash([string]$caminho) { (Get-FileHash -LiteralPath $caminho -Algorithm SHA256).Hash }

try {
    & $python -c "import sys; sys.path.insert(0, r'$raiz'); from pathlib import Path; from zepp_life_excel_ia import sintetico; sintetico.criar_projeto(Path(r'$t'))"
    # Leitor falso: ignora o pedido em stdin e devolve a resposta preparada.
    Set-Content -LiteralPath (Join-Path $t 'leitor.cmd') -Value "@type `"%~dp0resposta.json`"`r`n@exit /b 0" -Encoding ASCII
    # Sessão fictícia guardada com DPAPI na pasta privada DE TESTE (nunca na real).
    [void](New-Item -ItemType Directory -Path $privado)
    ConvertTo-SecureString 'utilizador:sessao-ficticia' -AsPlainText -Force | Export-Clixml -LiteralPath (Join-Path $privado 'sessao.clixml')
    $livro = Join-Path $resultados 'analise-historico-zepp-life.xlsx'
    $ia = Join-Path $resultados 'dados-zepp-life-ia.json'

    Write-Host '1. Primeira atualização'
    Resposta 100
    $r = Atualizar @()
    Confirmar ($r.Codigo -eq 0) 'termina com sucesso'
    Confirmar ((Test-Path $livro) -and (Test-Path $ia)) 'cria Excel e JSON'
    Confirmar (Test-Path (Join-Path $privado 'sessao.clixml')) 'guarda a sessão na pasta privada de teste'
    Confirmar (-not ((Get-Content (Join-Path $privado 'sessao.clixml') -Raw).Contains('token-de-teste'))) 'sessão não fica em texto legível'
    Confirmar (@(Get-ChildItem (Join-Path $dados 'temporario') -Directory -ErrorAction SilentlyContinue).Count -eq 0) 'remove a pasta temporária após sucesso'
    $csv = Get-Content (Join-Path $dados 'exportacao-zepp\zepp-completo.csv') -TotalCount 1
    Confirmar ($csv -like '*;*') 'CSV usa ponto e vírgula'

    Write-Host '2. Nova pesagem'
    Resposta 110
    $r = Atualizar @()
    Confirmar ($r.Codigo -eq 0) 'termina com sucesso'
    $copias = @(Get-ChildItem (Join-Path $dados 'historico') -Directory)
    Confirmar ($copias.Count -eq 1 -and @(Get-ChildItem $copias[0].FullName).Count -eq 4) 'guarda as 4 versões anteriores em historico'

    Write-Host '3. Exportação que perde pesagens antigas'
    $antes = Hash $livro
    Resposta 110 3
    $r = Atualizar @()
    Confirmar ($r.Codigo -eq 1) 'é recusada'
    Confirmar ($r.Texto -match 'perdeu 3 pesagem') 'explica quantas pesagens se perderiam'
    Confirmar ((Hash $livro) -eq $antes) 'Excel fica inalterado'

    Write-Host '4. Bloqueio libertado depois de uma falha'
    Resposta 110
    $r = Atualizar @()
    Confirmar ($r.Codigo -eq 0) 'a atualização seguinte não fica bloqueada'

    Write-Host '5. Excel aberto'
    $aberto = [IO.File]::Open($livro, 'Open', 'ReadWrite', 'None')
    try { $r = Atualizar @() } finally { $aberto.Dispose() }
    Confirmar ($r.Codigo -eq 1 -and $r.Texto -match 'feche o Excel') 'pede para fechar o Excel'

    Write-Host '6. Atualização em curso'
    $lock = [IO.File]::Open((Join-Path $dados '.atualizacao.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
    try { $r = Atualizar @() } finally { $lock.Dispose() }
    Confirmar ($r.Codigo -eq 1 -and $r.Texto -match 'em curso') 'recusa uma segunda atualização simultânea'

    Write-Host '7. Recalcular sem descarregar'
    $exportAntes = Hash (Join-Path $dados 'exportacao-zepp\zepp-completo.json')
    $r = Atualizar @('-SoExcel')
    Confirmar ($r.Codigo -eq 0) 'termina com sucesso'
    Confirmar ((Hash (Join-Path $dados 'exportacao-zepp\zepp-completo.json')) -eq $exportAntes) 'não altera a exportação'

    Write-Host '8. Falha do gerador repõe tudo'
    $antes = Hash $livro
    Copy-Item $livro (Join-Path $t 'bom.xlsx')
    [IO.File]::WriteAllText($livro, 'corrompido')
    $corrompido = Hash $livro
    $r = Atualizar @('-SoExcel')
    Confirmar ($r.Codigo -eq 1 -and $r.Texto -match 'preservar as medidas manuais') 'recusa um Excel ilegível em vez de perder medidas'
    Confirmar ((Hash $livro) -eq $corrompido) 'não substitui o ficheiro'
    Confirmar (@(Get-ChildItem (Join-Path $dados 'temporario') -Directory).Count -ge 1) 'guarda a tentativa falhada com o registo'
    Copy-Item (Join-Path $t 'bom.xlsx') $livro -Force

    Write-Host '9. Esquecer conta'
    $r = Atualizar @('-EsquecerConta')
    Confirmar ($r.Codigo -eq 0 -and -not (Test-Path (Join-Path $privado 'sessao.clixml'))) 'remove a sessão guardada'
} finally {
    if (Test-Path -LiteralPath $t) { Remove-Item -LiteralPath $t -Recurse -Force }
}
if ($falhas) { Write-Host "$falhas verificação(ões) falharam." -ForegroundColor Red; exit 1 }
Write-Host 'Todos os testes do atualizador passaram.' -ForegroundColor Green
