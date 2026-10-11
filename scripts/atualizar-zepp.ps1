<#
  Atualiza o Excel e o JSON para IA.
  - Sem parâmetros: descarrega as pesagens da Zepp Life (sessão/conta guardadas com DPAPI) e regenera.
  - -SoExcel: regenera a partir da exportação local, sem contactar a Xiaomi/Zepp.
  - -EsquecerConta: apaga a conta e a sessão guardadas neste computador.
  Os resultados são preparados numa pasta temporária, validados e só depois substituem os anteriores;
  as versões anteriores ficam em privado\historico\. Em caso de falha, nada é substituído.
  Pastas: 'Excel e JSON' (resultados que se abrem) e 'privado' (exportação, histórico, temporários).
#>
param(
    [switch]$SoExcel,
    [switch]$EsquecerConta,
    [switch]$SemPausa,
    [switch]$SemExcelReal,
    [switch]$PermitirRemocoes,
    [string]$PastaDados,
    [string]$PastaResultados,
    [string]$PastaPrivada = (Join-Path $env:LOCALAPPDATA 'ZeppLifeExcel'),
    [string]$Leitor,
    [string]$Config,
    [string]$Python,
    [string]$Hoje
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$raiz = Split-Path $PSScriptRoot -Parent
if (-not $PastaDados) { $PastaDados = Join-Path $raiz 'privado' }
if (-not $PastaResultados) { $PastaResultados = Join-Path $raiz 'Excel e JSON' }
if (-not $Leitor) { $Leitor = Join-Path $raiz 'bin\zepp-reader.exe' }
if (-not $Config) { $Config = Join-Path $raiz 'config.json' }
if (-not $Python) { $Python = Join-Path $raiz '.venv\Scripts\python.exe' }

$contaPath = Join-Path $PastaPrivada 'conta.clixml'
$sessaoPath = Join-Path $PastaPrivada 'sessao.clixml'
$exportJson = Join-Path $PastaDados 'exportacao-zepp\zepp-completo.json'
$exportCsv = Join-Path $PastaDados 'exportacao-zepp\zepp-completo.csv'
$livroPath = Join-Path $PastaResultados 'analise-historico-zepp-life.xlsx'
$iaPath = Join-Path $PastaResultados 'dados-zepp-life-ia.json'
$pastaTemporaria = Join-Path $PastaDados 'temporario'
$stage = $null
$lockStream = $null
$commitados = @()
$credencial = $null
$resposta = $null
$resultado = $null
$falhou = $false
$sucesso = $false
$logPath = $null

function Ler-Segredo([Security.SecureString]$segredo) {
    $ponteiro = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($segredo)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ponteiro) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ponteiro) }
}
function Guardar-Privado($valor, [string]$destino) {
    $temp = $destino + '.new'
    $valor | Export-Clixml -LiteralPath $temp -Encoding UTF8
    Move-Item -LiteralPath $temp -Destination $destino -Force
}
function Invoke-Leitor($pedido) {
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $Leitor
    $info.WorkingDirectory = $stage
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardInput = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $info.StandardOutputEncoding = [Text.UTF8Encoding]::new($false)
    $processo = New-Object Diagnostics.Process
    $processo.StartInfo = $info
    try {
        [void]$processo.Start()
        $saida = $processo.StandardOutput.ReadToEndAsync()
        $erro = $processo.StandardError.ReadToEndAsync()
        # Os segredos seguem por stdin; nunca por argumentos, ficheiros ou registos.
        $processo.StandardInput.WriteLine(($pedido | ConvertTo-Json -Compress))
        $processo.StandardInput.Close()
        if (-not $processo.WaitForExit(300000)) { $processo.Kill(); throw 'A Zepp demorou demasiado a responder. Tente mais tarde.' }
        return @{ Codigo = $processo.ExitCode; Saida = $saida.Result; Erro = $erro.Result }
    } finally { $processo.Dispose() }
}
function Concluir-Verificacao($res) {
    $m = [regex]::Match($res.Erro, 'XIAOMI_VERIFICATION_URL=(https://\S+)')
    if (-not $m.Success) { return $res }
    $uri = [uri]$m.Groups[1].Value
    if ($uri.Host -ne 'account.xiaomi.com') { throw 'Endereço de verificação inesperado.' }
    $html = Join-Path $stage 'verificacao.html'
    $codificado = [Net.WebUtility]::HtmlEncode($uri.AbsoluteUri)
    [IO.File]::WriteAllText($html, '<meta charset="utf-8"><a id="go" href="' + $codificado + '">Continuar na Xiaomi</a><script>document.getElementById("go").click()</script>', [Text.UTF8Encoding]::new($false))
    $chrome = Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'
    if (Test-Path -LiteralPath $chrome) { Start-Process -FilePath $chrome -ArgumentList @('--incognito', '--new-window', ('"' + $html + '"')) }
    else { Start-Process -FilePath $html }
    Write-Host 'Conclua a verificação Xiaomi no navegador.' -ForegroundColor Yellow
    Write-Host 'Na página final (mesmo com erro 404), copie o endereço com Ctrl+L e Ctrl+C.'
    [void](Read-Host 'Regresse aqui e prima Enter')
    $regresso = [uri](Get-Clipboard -Raw).Trim()
    if ($regresso.Scheme -ne 'https' -or $regresso.Host -ne 'api-mifit-cn.huami.com') { throw 'O endereço copiado não corresponde ao regresso da Xiaomi.' }
    $codigo = [regex]::Match($regresso.Query, '(?:^|[?&])code=([^&]+)')
    if (-not $codigo.Success) { throw 'Falta o código de autorização no endereço copiado.' }
    return Invoke-Leitor @{ Code = [Uri]::UnescapeDataString($codigo.Groups[1].Value) }
}
function Instante([string]$data) { return ([DateTimeOffset]::Parse($data, [Globalization.CultureInfo]::InvariantCulture)).ToUnixTimeSeconds() }
function Propriedade($objeto, [string]$nome) {
    $p = $objeto.PSObject.Properties[$nome]
    if ($null -eq $p) { return $null } else { return $p.Value }
}

try {
    if ($EsquecerConta) {
        foreach ($alvo in @($contaPath, $sessaoPath)) { if (Test-Path -LiteralPath $alvo) { Remove-Item -LiteralPath $alvo -Force } }
        Write-Host 'Conta e sessão guardadas removidas deste computador. Os dados e o Excel foram mantidos.'
        $sucesso = $true
        return
    }
    [void](New-Item -ItemType Directory -Force -Path $PastaDados)
    [void](New-Item -ItemType Directory -Force -Path $PastaResultados)
    try { $lockStream = [IO.File]::Open((Join-Path $PastaDados '.atualizacao.lock'), 'OpenOrCreate', 'ReadWrite', 'None') }
    catch { throw 'Já existe uma atualização em curso. Feche a outra janela de atualização e tente novamente.' }
    if (Test-Path -LiteralPath $livroPath) {
        try { $teste = [IO.File]::Open($livroPath, 'Open', 'ReadWrite', 'None'); $teste.Dispose() }
        catch { throw 'Guarde e feche o Excel antes de atualizar.' }
    }
    if (-not (Test-Path -LiteralPath $Python)) { throw 'O ambiente Python não está instalado. Execute INSTALAR.cmd.' }
    if (-not (Test-Path -LiteralPath $Config)) { throw 'Falta config.json. Copie config.example.json para config.json e ajuste os valores.' }
    $stage = Join-Path $pastaTemporaria ((Get-Date -Format 'yyyy-MM-dd_HH-mm-ss') + '_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    [void](New-Item -ItemType Directory -Path $stage -Force)
    $logPath = Join-Path $stage 'gerador.log'
    $novoJson = Join-Path $stage 'zepp-completo.json'
    $novoCsv = Join-Path $stage 'zepp-completo.csv'
    $novoLivro = Join-Path $stage 'analise-historico-zepp-life.xlsx'
    $novoIa = Join-Path $stage 'dados-zepp-life-ia.json'

    if (-not $SoExcel) {
        if (-not (Test-Path -LiteralPath $Leitor)) { throw "Leitor Zepp não encontrado em $Leitor. Consulte docs/INSTALACAO.md." }
        [void](New-Item -ItemType Directory -Path $PastaPrivada -Force)
        # Só o utilizador atual acede à pasta privada (apenas DACL; sem pedir privilégios de auditoria).
        $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User
        $acl = New-Object Security.AccessControl.DirectorySecurity
        $acl.SetAccessRuleProtection($true, $false)
        $acl.SetAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($sid, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')))
        [IO.Directory]::SetAccessControl($PastaPrivada, $acl)
        if (Test-Path -LiteralPath $sessaoPath) {
            Write-Host 'A reutilizar a sessão guardada...' -ForegroundColor Cyan
            try { $resultado = Invoke-Leitor @{ Token = (Ler-Segredo (Import-Clixml -LiteralPath $sessaoPath)) } }
            catch { Write-Host 'Não foi possível reutilizar a sessão.' }
        }
        if ($null -eq $resultado -or $resultado.Codigo -ne 0) {
            if (Test-Path -LiteralPath $contaPath) {
                $guardada = (Ler-Segredo (Import-Clixml -LiteralPath $contaPath)) | ConvertFrom-Json
                $credencial = New-Object Management.Automation.PSCredential($guardada.Account, (ConvertTo-SecureString $guardada.Password -AsPlainText -Force))
                $guardada = $null
            } else {
                Write-Host 'Primeira utilização: a conta será guardada cifrada para o seu utilizador Windows.' -ForegroundColor Cyan
                $conta = Read-Host 'Email ou identificador Xiaomi'
                if ([string]::IsNullOrWhiteSpace($conta)) { throw 'Falta a conta Xiaomi.' }
                $senha = Read-Host 'Palavra-passe Xiaomi' -AsSecureString
                $credencial = New-Object Management.Automation.PSCredential($conta.Trim(), $senha)
            }
            Write-Host 'A autenticar na Xiaomi...'
            $resultado = Invoke-Leitor @{ Account = $credencial.UserName; Password = $credencial.GetNetworkCredential().Password }
            $resultado = Concluir-Verificacao $resultado
        }
        if ($resultado.Codigo -ne 0) {
            $detalhe = ($resultado.Erro -split "`n" | Where-Object { $_ -and $_ -notmatch 'XIAOMI_VERIFICATION_URL' } | Select-Object -First 1)
            throw ('A autenticação ou a leitura Zepp falhou' + $(if ($detalhe) { ': ' + $detalhe.Trim() } else { '.' }) +
                ' Os dados anteriores foram preservados. Se mudou a palavra-passe, execute ESQUECER-CONTA.cmd.')
        }
        $resposta = $resultado.Saida | ConvertFrom-Json
        $dados = @($resposta.weights)
        if ($dados.Count -eq 0) { throw 'A exportação está vazia.' }
        $novos = @{}
        foreach ($linha in $dados) {
            $data = Propriedade $linha 'Date'; $peso = Propriedade $linha 'Weight'
            if ($null -eq $data -or $null -eq $peso -or [double]$peso -le 0) { throw 'A exportação contém pesagens inválidas.' }
            $chave = Instante ([string]$data)
            if ($novos.ContainsKey($chave)) { throw 'A exportação contém pesagens duplicadas no mesmo instante.' }
            $novos[$chave] = $true
        }
        # Uma exportação completa tem de manter todas as pesagens anteriores (comparação por instante UTC).
        if (Test-Path -LiteralPath $exportJson) {
            # No PowerShell 5.1, ConvertFrom-Json devolve a matriz como um único objeto: enumerar explicitamente.
            $antigos = @((Get-Content -LiteralPath $exportJson -Raw -Encoding UTF8 | ConvertFrom-Json) | ForEach-Object { $_ })
            $perdidas = @($antigos | Where-Object { $null -ne (Propriedade $_ 'Date') -and -not $novos.ContainsKey((Instante ([string]$_.Date))) })
            if ($perdidas.Count -gt 0) {
                $exemplos = ($perdidas | Select-Object -First 5 | ForEach-Object { $_.Date }) -join ', '
                if (-not $PermitirRemocoes) {
                    throw "A nova exportação perdeu $($perdidas.Count) pesagem(ns) anterior(es) ($exemplos). Atualização interrompida. Se as apagou de propósito na Zepp Life, execute o atualizador com -PermitirRemocoes."
                }
                Write-Host "Aviso: $($perdidas.Count) pesagem(ns) removida(s) na Zepp foram aceites (-PermitirRemocoes)." -ForegroundColor Yellow
            }
        }
        [IO.File]::WriteAllText($novoJson, (ConvertTo-Json -InputObject $dados -Depth 10), [Text.UTF8Encoding]::new($false))
        $dados | Export-Csv -LiteralPath $novoCsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        Guardar-Privado (ConvertTo-SecureString -String $resposta.token -AsPlainText -Force) $sessaoPath
        if ($null -ne $credencial) {
            $pacote = @{ Account = $credencial.UserName; Password = $credencial.GetNetworkCredential().Password } | ConvertTo-Json -Compress
            Guardar-Privado (ConvertTo-SecureString $pacote -AsPlainText -Force) $contaPath
            $pacote = $null
        }
        Write-Host "Exportação recebida: $($dados.Count) pesagens." -ForegroundColor Cyan
    } else {
        if (-not (Test-Path -LiteralPath $exportJson)) { throw 'Ainda não existe exportação local. Execute ATUALIZAR-EXCEL.cmd.' }
        Copy-Item -LiteralPath $exportJson -Destination $novoJson
    }

    Write-Host 'A gerar o Excel e o JSON e a preservar as medidas manuais...' -ForegroundColor Cyan
    $argumentos = @('-m', 'zepp_transformacao.gerar', '--origem', $novoJson, '--existente', $livroPath,
        '--excel', $novoLivro, '--json', $novoIa, '--config', $Config)
    if ($SemExcelReal) { $argumentos += '--sem-excel-real' }
    if ($Hoje) { $argumentos += @('--hoje', $Hoje) }
    # O Windows PowerShell trata stderr nativo como erro: capturar tudo e decidir pelo código de saída.
    $preferencia = $ErrorActionPreference
    $env:PYTHONIOENCODING = 'utf-8'
    # O PowerShell 5.1 descodifica a saída do Python com a página de código da consola (cp850),
    # o que estraga os acentos; usar UTF-8 durante a chamada e repor a codificação no fim.
    $codificacaoConsola = $null
    try { $codificacaoConsola = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false) } catch { }
    Push-Location $raiz
    try {
        $ErrorActionPreference = 'Continue'
        & $Python @argumentos *> $logPath
        $codigoGerador = $LASTEXITCODE
    } finally {
        Pop-Location; $ErrorActionPreference = $preferencia
        if ($null -ne $codificacaoConsola) { try { [Console]::OutputEncoding = $codificacaoConsola } catch { } }
    }
    Get-Content -LiteralPath $logPath | Where-Object { $_ -match '^(Excel real|Aviso|LEITURAS=)' } | ForEach-Object { Write-Host $_ }
    if ($codigoGerador -ne 0 -or -not (Test-Path -LiteralPath $novoLivro) -or -not (Test-Path -LiteralPath $novoIa)) {
        $motivo = (Get-Content -LiteralPath $logPath | Where-Object { $_ -match '^ERRO:' } | Select-Object -First 1)
        throw ('Falha ao gerar o Excel/JSON. ' + $(if ($motivo) { $motivo + ' ' } else { '' }) + 'Registo completo: ' + $logPath)
    }
    $ia = Get-Content -LiteralPath $novoIa -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($ia.metadados.sha256_excel -ne (Get-FileHash -LiteralPath $novoLivro -Algorithm SHA256).Hash.ToLowerInvariant()) { throw 'O Excel e o JSON não correspondem (hash diferente).' }

    $pares = @(@($novoLivro, $livroPath), @($novoIa, $iaPath))
    if (-not $SoExcel) { $pares = @(@($novoJson, $exportJson), @($novoCsv, $exportCsv)) + $pares }
    $copias = Join-Path $PastaDados ('historico\' + (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss-fff'))
    [void](New-Item -ItemType Directory -Path $copias -Force)
    [void](New-Item -ItemType Directory -Path (Split-Path $exportJson) -Force)
    foreach ($par in $pares) {
        $copia = Join-Path $copias ([IO.Path]::GetFileName($par[1]))
        if (Test-Path -LiteralPath $par[1]) { [IO.File]::Replace($par[0], $par[1], $copia); $commitados += , @($par[1], $copia) }
        else { [IO.File]::Move($par[0], $par[1]); $commitados += , @($par[1], $null) }
    }
    if (@(Get-ChildItem -LiteralPath $copias).Count -eq 0) { [IO.Directory]::Delete($copias) }
    $sucesso = $true
    Write-Host 'Atualização concluída. Excel e JSON para IA validados; versões anteriores guardadas em privado\historico.' -ForegroundColor Green
    Write-Host $livroPath
    Write-Host $iaPath
} catch {
    $falhou = $true
    $mensagem = $_.Exception.Message
    foreach ($par in $commitados) {
        try {
            if ($null -ne $par[1]) { Copy-Item -LiteralPath $par[1] -Destination $par[0] -Force }
            elseif (Test-Path -LiteralPath $par[0]) { Remove-Item -LiteralPath $par[0] -Force }
        } catch { $mensagem += "`nATENÇÃO: não foi possível repor $($par[0]); a cópia anterior está em $($par[1])." }
    }
    if ($commitados.Count) { $mensagem += "`nOs ficheiros anteriores foram repostos." }
} finally {
    foreach ($variavel in @('PYTHONIOENCODING')) { [Environment]::SetEnvironmentVariable($variavel, $null) }
    # Libertar o bloqueio ANTES de qualquer pausa: uma janela de erro aberta não pode bloquear a próxima atualização.
    if ($null -ne $lockStream) { $lockStream.Dispose() }
    if ($null -ne $stage -and (Test-Path -LiteralPath $stage)) {
        $html = Join-Path $stage 'verificacao.html'
        if (Test-Path -LiteralPath $html) { Remove-Item -LiteralPath $html -Force }
        if ($sucesso) { Remove-Item -LiteralPath $stage -Recurse -Force }
    }
    # Guardar só as 5 tentativas falhadas mais recentes (contêm dados pessoais).
    if (Test-Path -LiteralPath $pastaTemporaria) {
        Get-ChildItem -LiteralPath $pastaTemporaria -Directory | Sort-Object Name -Descending | Select-Object -Skip 5 |
            ForEach-Object { Remove-Item -LiteralPath $_.FullName -Recurse -Force }
    }
    $credencial = $null; $resposta = $null; $resultado = $null
}
if ($falhou) {
    Write-Host $mensagem -ForegroundColor Red
    if ($null -ne $logPath -and (Test-Path -LiteralPath $logPath)) { Write-Host "Detalhes: $logPath" }
    if (-not $SemPausa) { [void](Read-Host 'Prima Enter para fechar') }
    exit 1
}
if (-not $SemPausa) { [void](Read-Host 'Prima Enter para fechar') }
exit 0
