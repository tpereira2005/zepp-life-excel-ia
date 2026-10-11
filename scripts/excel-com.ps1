<#
  Apoio comum à automação do Microsoft Excel por COM (dot-source: . .\excel-com.ps1).
  Regista um filtro de mensagens OLE: quando o Excel está ocupado (a calcular ou a abrir),
  rejeita chamadas com RPC_E_CALL_REJECTED; o filtro faz o COM repetir a chamada sozinho,
  em vez de a deixar falhar a meio. Requer um thread STA (o Windows PowerShell 5.1 usa STA).
#>
if (-not ('ZeppExcel.FiltroMensagens' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace ZeppExcel {
    [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleMessageFilter {
        [PreserveSig] int HandleInComingCall(int tipo, IntPtr tarefa, int tempo, IntPtr info);
        [PreserveSig] int RetryRejectedCall(IntPtr tarefa, int tempo, int tipoRejeicao);
        [PreserveSig] int MessagePending(IntPtr tarefa, int tempo, int tipoPendente);
    }
    public class FiltroMensagens : IOleMessageFilter {
        [DllImport("Ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter novo, out IOleMessageFilter antigo);
        public static void Registar() { IOleMessageFilter antigo; CoRegisterMessageFilter(new FiltroMensagens(), out antigo); }
        public static void Revogar() { IOleMessageFilter antigo; CoRegisterMessageFilter(null, out antigo); }
        public int HandleInComingCall(int tipo, IntPtr tarefa, int tempo, IntPtr info) { return 0; }
        // SERVERCALL_RETRYLATER (2): voltar a tentar daqui a 250 ms, durante até 3 minutos.
        public int RetryRejectedCall(IntPtr tarefa, int tempo, int tipoRejeicao) {
            return (tipoRejeicao == 2 && tempo < 180000) ? 250 : -1;
        }
        public int MessagePending(IntPtr tarefa, int tempo, int tipoPendente) { return 2; }
    }
}
'@
}

function Iniciar-FiltroExcel {
    if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') { throw 'A automação do Excel exige um thread STA.' }
    [ZeppExcel.FiltroMensagens]::Registar()
}

function Parar-FiltroExcel { [ZeppExcel.FiltroMensagens]::Revogar() }

# Espera até o Excel estar pronto (Ready) — por exemplo, depois de abrir ou recalcular o livro.
function Esperar-Excel($excel, [int]$segundos = 120) {
    $limite = (Get-Date).AddSeconds($segundos)
    while (-not $excel.Ready) {
        if ((Get-Date) -gt $limite) { throw "O Excel não ficou pronto em $segundos segundos." }
        Start-Sleep -Milliseconds 200
    }
}

# Abre o livro só de leitura e devolve-o. Com outro Excel aberto, Open pode devolver null
# apesar de abrir o livro: nesse caso obtém-no da coleção pelo nome.
function Abrir-LivroExcel($excel, [string]$caminho) {
    $livro = $excel.Workbooks.Open($caminho, 0, $true)
    if ($null -eq $livro) {
        $nome = [IO.Path]::GetFileName($caminho)
        foreach ($aberto in $excel.Workbooks) { if ($aberto.Name -eq $nome) { $livro = $aberto } }
    }
    if ($null -eq $livro) { throw 'O Excel não abriu o livro.' }
    Esperar-Excel $excel
    return , $livro
}
