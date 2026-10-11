<#
  Abre uma CÓPIA do livro no Microsoft Excel instalado (só leitura), recalcula tudo e grava
  os valores calculados em JSON. Opcionalmente exporta os gráficos do Painel em PNG.
  Não altera o livro original. Código de saída 3 = Excel não disponível.
  -FolhasEsperadas 'A|B|C': a leitura falha (em vez de devolver dados parciais) se faltar alguma destas folhas.
#>
param(
    [Parameter(Mandatory)][string]$Livro,
    [Parameter(Mandatory)][string]$Saida,
    [string]$Graficos,
    [string]$FolhasEsperadas = ''
)
$ErrorActionPreference = 'Stop'
$esperadas = @($FolhasEsperadas -split '\|' | Where-Object { $_ })
. (Join-Path $PSScriptRoot 'excel-com.ps1')

$copia = Join-Path ([IO.Path]::GetTempPath()) ('zepp-verificacao-' + [guid]::NewGuid().ToString('N') + '.xlsx')
Copy-Item -LiteralPath $Livro -Destination $copia
$excel = $null
Iniciar-FiltroExcel
try {
    try { $excel = New-Object -ComObject Excel.Application }
    catch { Write-Host 'Excel nao disponivel para verificacao.'; exit 3 }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $wb = Abrir-LivroExcel $excel $copia
    $excel.CalculateFull()
    Esperar-Excel $excel

    # Confirmar que o livro está completo antes de ler (folhas esperadas presentes).
    $limite = (Get-Date).AddSeconds(60)
    do {
        $nomes = @(foreach ($folha in $wb.Worksheets) { $folha.Name })
        $faltam = @($esperadas | Where-Object { $nomes -notcontains $_ })
        if ($faltam.Count -eq 0 -and $nomes.Count -gt 0) { break }
        if ((Get-Date) -gt $limite) { throw ('Folhas em falta no Excel: ' + ($faltam -join ', ')) }
        Start-Sleep -Milliseconds 500
    } while ($true)

    $resultado = [ordered]@{ versao_excel = [string]$excel.Version; folhas = [ordered]@{}; erros = @(); graficos = @() }
    foreach ($nome in $nomes) {
        $ws = $wb.Worksheets.Item($nome)
        $usado = $ws.UsedRange
        $valores = $usado.Value2      # atribuição direta: o pipeline achataria a matriz 2D
        $linhas = New-Object System.Collections.Generic.List[object]
        $primeiraLinha = $usado.Row; $primeiraColuna = $usado.Column
        # Linhas vazias antes da área usada, para manter os endereços Excel.
        for ($r = 1; $r -lt $primeiraLinha; $r++) { $linhas.Add(@()) }
        if ($valores -is [Array]) {
            $nl = $valores.GetLength(0); $nc = $valores.GetLength(1)
            for ($r = 1; $r -le $nl; $r++) {
                $linha = New-Object object[] ($nc + $primeiraColuna - 1)
                for ($c = 1; $c -le $nc; $c++) {
                    $v = $valores[$r, $c]
                    if ($v -is [int]) {
                        # Value2 devolve os erros de fórmula como inteiros (CVErr).
                        $endereco = $ws.Cells.Item($r + $primeiraLinha - 1, $c + $primeiraColuna - 1).Address($false, $false)
                        $resultado.erros += ($ws.Name + '!' + $endereco + ' ' + $ws.Range($endereco).Text)
                        $v = '#ERRO'
                    }
                    $linha[$c + $primeiraColuna - 2] = $v
                }
                $linhas.Add($linha)
            }
        }
        $resultado.folhas[$nome] = $linhas
    }
    if ($resultado.folhas.Count -ne $nomes.Count) { throw "Leitura incompleta: $($resultado.folhas.Count) de $($nomes.Count) folhas." }

    if ($Graficos) {
        [void](New-Item -ItemType Directory -Force -Path $Graficos)
        $painel = $wb.Worksheets.Item('Painel')
        $painel.Activate()
        $i = 0
        foreach ($objeto in $painel.ChartObjects()) {
            $i++
            $objeto.Activate()
            $destino = Join-Path $Graficos ('grafico-' + $i + '.png')
            [void]$objeto.Chart.Export($destino, 'PNG')
            $resultado.graficos += [ordered]@{ ficheiro = $destino; titulo = $objeto.Chart.ChartTitle.Text;
                series = $objeto.Chart.SeriesCollection().Count; vazios = $objeto.Chart.DisplayBlanksAs }
        }
    }
    $wb.Close($false)
    $json = $resultado | ConvertTo-Json -Depth 6 -Compress
    [IO.File]::WriteAllText($Saida, $json, [Text.UTF8Encoding]::new($false))
} finally {
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    Parar-FiltroExcel
    if (Test-Path -LiteralPath $copia) { [IO.File]::Delete($copia) }
}
