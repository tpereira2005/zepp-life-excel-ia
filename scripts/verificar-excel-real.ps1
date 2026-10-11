<#
  Abre uma CÓPIA do livro no Microsoft Excel instalado (só leitura), recalcula tudo e grava
  os valores calculados em JSON. Opcionalmente exporta os gráficos do Painel em PNG.
  Não altera o livro original. Código de saída 3 = Excel não disponível.
#>
param(
    [Parameter(Mandatory)][string]$Livro,
    [Parameter(Mandatory)][string]$Saida,
    [string]$Graficos
)
$ErrorActionPreference = 'Stop'
# O Excel rejeita chamadas COM enquanto calcula (RPC_E_CALL_REJECTED); tentar de novo.
# Os blocos atribuem a variáveis $script: porque o pipeline enumeraria Ranges e matrizes.
function Repetir([scriptblock]$acao) {
    for ($tentativa = 1; ; $tentativa++) {
        try { & $acao; return }
        catch {
            if ($tentativa -ge 120 -or $_.Exception.Message -notmatch '0x8001010A|0x80010001|RPC_E') { throw }
            Start-Sleep -Milliseconds 500
        }
    }
}
$copia = Join-Path ([IO.Path]::GetTempPath()) ('zepp-verificacao-' + [guid]::NewGuid().ToString('N') + '.xlsx')
Copy-Item -LiteralPath $Livro -Destination $copia
$excel = $null
try {
    try { $excel = New-Object -ComObject Excel.Application }
    catch { Write-Host 'Excel nao disponivel para verificacao.'; exit 3 }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    Repetir { $script:wb = $excel.Workbooks.Open($copia, 0, $true) }
    Repetir { $excel.CalculateFull() }
    Repetir { [void]$excel.Ready }
    $resultado = [ordered]@{ versao_excel = [string]$excel.Version; folhas = [ordered]@{}; erros = @(); graficos = @() }
    foreach ($ws in $wb.Worksheets) {
        Repetir { $script:usado = $ws.UsedRange }
        # Atribuição direta: devolver a matriz 2D por pipeline achata-a.
        Repetir { $script:valores = $usado.Value2 }
        $linhas = New-Object System.Collections.Generic.List[object]
        $primeiraLinha = $usado.Row; $primeiraColuna = $usado.Column
        # Linhas e colunas vazias antes da área usada para manter endereços Excel.
        for ($r = 1; $r -lt $primeiraLinha; $r++) { $linhas.Add(@()) }
        if ($valores -is [Array]) {
            $nl = $valores.GetLength(0); $nc = $valores.GetLength(1)
            for ($r = 1; $r -le $nl; $r++) {
                $linha = New-Object object[] ($nc + $primeiraColuna - 1)
                for ($c = 1; $c -le $nc; $c++) {
                    $v = $valores[$r, $c]
                    if ($v -is [int]) {
                        $endereco = $ws.Cells.Item($r + $primeiraLinha - 1, $c + $primeiraColuna - 1).Address($false, $false)
                        $resultado.erros += ($ws.Name + '!' + $endereco + ' ' + $ws.Range($endereco).Text)
                        $v = '#ERRO'
                    }
                    $linha[$c + $primeiraColuna - 2] = $v
                }
                $linhas.Add($linha)
            }
        }
        $resultado.folhas[$ws.Name] = $linhas
    }
    if ($Graficos) {
        [void](New-Item -ItemType Directory -Force -Path $Graficos)
        $painel = $wb.Worksheets.Item('Painel')
        $painel.Activate()
        $i = 0
        foreach ($objeto in $painel.ChartObjects()) {
            $i++
            $objeto.Activate()
            $destino = Join-Path $Graficos ('grafico-' + $i + '.png')
            Repetir { [void]$objeto.Chart.Export($destino, 'PNG') }
            $resultado.graficos += [ordered]@{ ficheiro = $destino; titulo = $objeto.Chart.ChartTitle.Text;
                series = $objeto.Chart.SeriesCollection().Count; vazios = $objeto.Chart.DisplayBlanksAs }
        }
    }
    Repetir { $wb.Close($false) }
    $json = $resultado | ConvertTo-Json -Depth 6 -Compress
    [IO.File]::WriteAllText($Saida, $json, [Text.UTF8Encoding]::new($false))
} finally {
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    if (Test-Path -LiteralPath $copia) { [IO.File]::Delete($copia) }
}
