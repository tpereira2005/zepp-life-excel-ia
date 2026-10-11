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
. (Join-Path $PSScriptRoot 'excel-com.ps1')
[void](New-Item -ItemType Directory -Force -Path $Pasta)
$copia = Join-Path ([IO.Path]::GetTempPath()) ('zepp-captura-' + [guid]::NewGuid().ToString('N') + '.xlsx')
Copy-Item -LiteralPath $Livro -Destination $copia
Iniciar-FiltroExcel
$excel = New-Object -ComObject Excel.Application
try {
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $wb = Abrir-LivroExcel $excel $copia
    $excel.CalculateFull()
    Esperar-Excel $excel
    foreach ($item in $Intervalos) {
        $folha, $endereco = $item -split '!', 2
        $ws = $wb.Worksheets.Item($folha)
        $ws.Activate()
        $intervalo = $ws.Range($endereco)
        Repetir-Excel { [void]$intervalo.CopyPicture(1, 2) }   # xlScreen, xlBitmap
        $grafico = $ws.ChartObjects().Add(0, 0, $intervalo.Width, $intervalo.Height)
        $grafico.Activate()
        $grafico.Chart.ChartArea.Format.Line.Visible = 0
        Repetir-Excel { [void]$grafico.Chart.Paste() }
        $nome = ($folha.Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', '' -replace '[^A-Za-z0-9]+', '-').ToLowerInvariant().Trim('-')
        $destino = Join-Path $Pasta ($nome + '.png')
        for ($n = 2; Test-Path -LiteralPath $destino; $n++) { $destino = Join-Path $Pasta ($nome + '-' + $n + '.png') }
        Repetir-Excel { [void]$grafico.Chart.Export($destino, 'PNG') }
        $grafico.Delete()
        Write-Host $destino
    }
    $wb.Close($false)
} finally {
    $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Parar-FiltroExcel
    if (Test-Path -LiteralPath $copia) { [IO.File]::Delete($copia) }
}
