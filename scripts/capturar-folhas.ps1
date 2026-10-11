<#
  Captura imagens PNG de intervalos de folhas tal como o Microsoft Excel as desenha.
  Usado para rever o design e criar as imagens de demonstração (com dados fictícios).
  Exemplo: -Livro demo.xlsx -Pasta docs\imagens -Intervalos 'Painel!A1:O40','Resumo Semanal!A1:V20'
#>
param(
    [Parameter(Mandatory)][string]$Livro,
    [Parameter(Mandatory)][string]$Pasta,
    [Parameter(Mandatory)][string[]]$Intervalos
)
$ErrorActionPreference = 'Stop'
function Repetir([scriptblock]$acao) {
    for ($tentativa = 1; ; $tentativa++) {
        try { & $acao; return }
        catch { if ($tentativa -ge 120 -or $_.Exception.Message -notmatch '0x8001010A|0x80010001|0x800A03EC|RPC_E|CopyPicture') { throw }; Start-Sleep -Milliseconds 500 }
    }
}
[void](New-Item -ItemType Directory -Force -Path $Pasta)
$copia = Join-Path ([IO.Path]::GetTempPath()) ('zepp-captura-' + [guid]::NewGuid().ToString('N') + '.xlsx')
Copy-Item -LiteralPath $Livro -Destination $copia
$excel = New-Object -ComObject Excel.Application
try {
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($copia, 0, $true)
    Repetir { $excel.CalculateFull() }
    foreach ($item in $Intervalos) {
        $folha, $endereco = $item -split '!', 2
        $ws = $wb.Worksheets.Item($folha)
        $ws.Activate()
        $intervalo = $ws.Range($endereco)
        Repetir { [void]$intervalo.CopyPicture(1, 2) }   # xlScreen, xlBitmap
        $grafico = $ws.ChartObjects().Add(0, 0, $intervalo.Width, $intervalo.Height)
        $grafico.Activate()
        $grafico.Chart.ChartArea.Format.Line.Visible = 0
        Repetir { [void]$grafico.Chart.Paste() }
        $nome = ($folha.Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', '' -replace '[^A-Za-z0-9]+', '-').ToLowerInvariant().Trim('-')
        $destino = Join-Path $Pasta ($nome + '.png')
        for ($n = 2; Test-Path -LiteralPath $destino; $n++) { $destino = Join-Path $Pasta ($nome + '-' + $n + '.png') }
        Repetir { [void]$grafico.Chart.Export($destino, 'PNG') }
        $grafico.Delete()
        Write-Host $destino
    }
    $wb.Close($false)
} finally {
    $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $copia) { [IO.File]::Delete($copia) }
}
