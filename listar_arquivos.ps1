<#
.SYNOPSIS
    Lista todos os arquivos e pastas de um ou mais diretórios (incluindo subpastas) e grava
    tudo em um único arquivo de texto CSV, SEM baixar os arquivos "somente online" do OneDrive.

.DESCRIPTION
    Lê apenas os metadados do sistema de arquivos (as mesmas informações que o Explorer
    mostra na listagem: nome, tamanho, datas, atributos). Nenhum arquivo é aberto ou lido,
    então os arquivos que estão só na nuvem continuam só na nuvem.

    Cada execução gera um CSV separado por ';' (abre no Bloco de Notas e no Excel), com uma
    linha por item de todas as pastas selecionadas; a coluna PastaOrigem indica de qual pasta
    selecionada o item veio. Pastas que não puderam ser lidas aparecem com Tipo = ERRO, para
    que uma listagem incompleta não passe despercebida.

.PARAMETER Pasta
    Uma ou mais pastas a serem analisadas. Se omitido, abre uma janela onde é possível
    selecionar várias pastas de uma vez (Ctrl+clique ou Shift+clique); depois o script
    pergunta se você quer adicionar pastas de outro local.

.PARAMETER Destino
    Pasta onde o inventário será gravado (criada se não existir). Padrão: a subpasta
    "inventarios" ao lado do script.
    Nome do arquivo: inventario_<pasta>_<data_hora>.csv quando uma pasta é analisada, ou
    inventario_<n>_pastas_<data_hora>.csv quando são várias.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\listar_arquivos.ps1 "C:\Users\fulano\OneDrive\Documentos" "E:\Backup\Documentos"

.EXAMPLE
    Selecionar uma ou mais pastas no Explorer e arrastá-las para cima do listar_arquivos.bat.
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Pasta,
    [string]$Destino
)

$ErrorActionPreference = 'Stop'

# Atributos de arquivos na nuvem que o enum FileAttributes do .NET não nomeia
$ATTR_OFFLINE               = 0x00001000
$ATTR_RECALL_ON_OPEN        = 0x00040000
$ATTR_PINNED                = 0x00080000
$ATTR_UNPINNED              = 0x00100000
$ATTR_RECALL_ON_DATA_ACCESS = 0x00400000
$ATTR_NOMEADOS              = 0x0002FFFF  # bits que o FileAttributes sabe converter em texto

# Janela "Selecionar pasta" do Windows com seleção múltipla. A FolderBrowserDialog do .NET
# só permite escolher uma pasta, por isso a interface COM da janela é usada diretamente.
$codigoSeletor = @'
using System;
using System.Runtime.InteropServices;

public static class SeletorDePastas
{
    [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
    private class FileOpenDialog { }

    [ComImport, Guid("d57c7288-d4ad-4768-be02-9d969532d960"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileOpenDialog
    {
        [PreserveSig] int Show(IntPtr hwndOwner);
        void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
        void SetFileTypeIndex(uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise(IntPtr pfde, out uint pdwCookie);
        void Unadvise(uint dwCookie);
        void SetOptions(uint fos);
        void GetOptions(out uint pfos);
        void SetDefaultFolder(IShellItem psi);
        void SetFolder(IShellItem psi);
        void GetFolder(out IShellItem ppsi);
        void GetCurrentSelection(out IShellItem ppsi);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetResult(out IShellItem ppsi);
        void AddPlace(IShellItem psi, int fdap);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close(int hr);
        void SetClientGuid(ref Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr pFilter);
        void GetResults(out IShellItemArray ppenum);
        void GetSelectedItems(out IShellItemArray ppsai);
    }

    [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem
    {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    [ComImport, Guid("b63ea76d-1f85-456f-a19c-48159efa858b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItemArray
    {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppvOut);
        void GetPropertyStore(int flags, ref Guid riid, out IntPtr ppv);
        void GetPropertyDescriptionList(IntPtr keyType, ref Guid riid, out IntPtr ppv);
        void GetAttributes(int attribFlags, uint sfgaoMask, out uint psfgaoAttribs);
        void GetCount(out uint pdwNumItems);
        void GetItemAt(uint dwIndex, out IShellItem ppsi);
        void EnumItems(out IntPtr ppenumShellItems);
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern IShellItem SHCreateItemFromParsingName(string pszPath, IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid riid);

    private const uint FOS_PICKFOLDERS = 0x20, FOS_FORCEFILESYSTEM = 0x40, FOS_ALLOWMULTISELECT = 0x200, FOS_PATHMUSTEXIST = 0x800;
    private const uint SIGDN_FILESYSPATH = 0x80058000;

    public static string[] Escolher(string titulo, string pastaInicial)
    {
        IFileOpenDialog dlg = (IFileOpenDialog)new FileOpenDialog();
        try
        {
            uint opcoes;
            dlg.GetOptions(out opcoes);
            dlg.SetOptions(opcoes | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_ALLOWMULTISELECT | FOS_PATHMUSTEXIST);
            dlg.SetTitle(titulo);
            if (!string.IsNullOrEmpty(pastaInicial))
            {
                try { dlg.SetFolder(SHCreateItemFromParsingName(pastaInicial, IntPtr.Zero, typeof(IShellItem).GUID)); }
                catch (COMException) { }  // pasta inicial inválida: abre no último local usado
            }
            if (dlg.Show(IntPtr.Zero) != 0) return new string[0];  // cancelado

            IShellItemArray itens;
            dlg.GetResults(out itens);
            uint total;
            itens.GetCount(out total);
            string[] caminhos = new string[total];
            for (uint i = 0; i < total; i++)
            {
                IShellItem item;
                itens.GetItemAt(i, out item);
                item.GetDisplayName(SIGDN_FILESYSPATH, out caminhos[i]);
            }
            return caminhos;
        }
        finally { Marshal.ReleaseComObject(dlg); }
    }
}
'@

function Select-Pastas {
    Add-Type -TypeDefinition $codigoSeletor
    Add-Type -AssemblyName System.Windows.Forms
    $inicial = if ($env:OneDrive) { $env:OneDrive } else { $env:USERPROFILE }
    $escolhidas = New-Object System.Collections.Generic.List[string]

    while ($true) {
        $novas = [SeletorDePastas]::Escolher('Selecione as pastas a listar (Ctrl+clique para escolher várias)', $inicial)
        if (-not $novas.Count) { break }
        foreach ($p in $novas) { $escolhidas.Add($p) }

        $mensagem = "Pastas selecionadas ($($escolhidas.Count)):`n`n" + ($escolhidas -join "`n") +
                    "`n`nDeseja adicionar pastas de outro local?"
        $resposta = [System.Windows.Forms.MessageBox]::Show($mensagem, 'Listar arquivos', 'YesNo', 'Question')
        if ($resposta -ne [System.Windows.Forms.DialogResult]::Yes) { break }
    }
    , $escolhidas.ToArray()
}

function Get-StatusNuvem([int]$attr) {
    if ($attr -band ($ATTR_RECALL_ON_DATA_ACCESS -bor $ATTR_RECALL_ON_OPEN -bor $ATTR_OFFLINE)) { 'Somente online' }
    elseif (($attr -band $ATTR_PINNED) -and -not ($attr -band $ATTR_UNPINNED)) { 'Sempre manter neste dispositivo' }
    else { 'Disponivel localmente' }
}

function Format-Tamanho([long]$bytes) {
    if ($bytes -ge 1GB) { '{0:N2} GB' -f ($bytes / 1GB) }
    elseif ($bytes -ge 1MB) { '{0:N2} MB' -f ($bytes / 1MB) }
    elseif ($bytes -ge 1KB) { '{0:N1} KB' -f ($bytes / 1KB) }
    else { "$bytes B" }
}

# "OneDrive - Fotos de Férias" -> "onedrive_fotos_de_ferias"
function ConvertTo-SnakeCase([string]$texto) {
    $semAcento = -join ($texto.Normalize([System.Text.NormalizationForm]::FormD).ToCharArray() |
        Where-Object { [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne 'NonSpacingMark' })
    ($semAcento.ToLowerInvariant() -replace '[^a-z0-9]+', '_').Trim('_')
}

# O prefixo \\?\ permite ler caminhos com mais de 260 caracteres no PowerShell 5.1
function ConvertTo-CaminhoLongo([string]$caminho) {
    if ($caminho.StartsWith('\\?\')) { $caminho }
    elseif ($caminho.StartsWith('\\')) { '\\?\UNC\' + $caminho.Substring(2) }
    else { '\\?\' + $caminho }
}

function New-Linha([System.IO.FileSystemInfo]$item) {
    $rel = $item.FullName.Substring($inicioRelativo)
    $barra = $rel.LastIndexOf('\')
    $attr = [int]$item.Attributes
    $status = Get-StatusNuvem $attr

    if ($item -is [System.IO.DirectoryInfo]) {
        $stats.Pastas++
        $tipo = 'Pasta'; $bytes = ''; $tamanho = ''; $ext = ''
    }
    else {
        $stats.Arquivos++
        $stats.Bytes += $item.Length
        if ($status -eq 'Somente online') { $stats.SomenteOnline++ }
        $tipo = 'Arquivo'; $bytes = $item.Length; $tamanho = Format-Tamanho $item.Length; $ext = $item.Extension
    }

    if ((($stats.Arquivos + $stats.Pastas) % 1000) -eq 0) {
        Write-Progress -Activity "Listando $raiz" -Status ('{0} arquivos, {1} pastas' -f $stats.Arquivos, $stats.Pastas)
    }

    [pscustomobject]@{
        PastaOrigem     = $raiz
        Tipo            = $tipo
        CaminhoRelativo = $rel
        Nome            = $item.Name
        PastaRelativa   = if ($barra -ge 0) { $rel.Substring(0, $barra) } else { '' }
        Extensao        = $ext
        Nivel           = $rel.Split('\').Count
        TamanhoBytes    = $bytes
        Tamanho         = $tamanho
        Criado          = $item.CreationTime.ToString('yyyy-MM-dd HH:mm:ss')
        Modificado      = $item.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
        Atributos       = ([System.IO.FileAttributes]($attr -band $ATTR_NOMEADOS)).ToString()
        StatusOneDrive  = $status
        Observacao      = ''
    }
}

# --- Pastas de origem ------------------------------------------------------------------
if (-not $Pasta) {
    $Pasta = Select-Pastas
    if (-not $Pasta) { Write-Host 'Cancelado.'; exit 1 }
}
$naoEncontradas = @($Pasta | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Container) })
if ($naoEncontradas.Count) { throw "Pasta(s) não encontrada(s):`n  " + ($naoEncontradas -join "`n  ") }

$jaIncluidas = @{}  # ignora a mesma pasta escolhida duas vezes (não diferencia maiúsculas)
$raizes = @(foreach ($p in $Pasta) {
    $r = (Resolve-Path -LiteralPath $p).ProviderPath.TrimEnd('\')
    if (-not $jaIncluidas.ContainsKey($r)) { $jaIncluidas[$r] = $true; $r }
})

# --- Pasta de destino ------------------------------------------------------------------
if (-not $Destino) { $Destino = Join-Path $(if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }) 'inventarios' }
if (-not (Test-Path -LiteralPath $Destino -PathType Container)) { New-Item -ItemType Directory -Path $Destino | Out-Null }
$Destino = (Resolve-Path -LiteralPath $Destino).ProviderPath

# --- Arquivo de saída (um só para todas as pastas) -------------------------------------
if ($raizes.Count -eq 1) {
    $nome = ConvertTo-SnakeCase $raizes[0].Substring($raizes[0].LastIndexOf('\') + 1)  # Split-Path trataria "D:" como a pasta atual de D:
    if (-not $nome) { $nome = 'pasta' }
}
else { $nome = '{0}_pastas' -f $raizes.Count }
$baseNome = 'inventario_{0}_{1:yyyyMMdd_HHmmss}' -f $nome, (Get-Date)
$saida = Join-Path $Destino "$baseNome.csv"
for ($n = 2; Test-Path -LiteralPath $saida; $n++) { $saida = Join-Path $Destino "${baseNome}_$n.csv" }  # não sobrescreve
$saidaLonga = ConvertTo-CaminhoLongo $saida

# --- Varredura de cada pasta -----------------------------------------------------------
$resumo = New-Object System.Collections.Generic.List[object]

& {
    $numero = 0
    foreach ($raiz in $raizes) {
        $numero++
        $raizLonga = ConvertTo-CaminhoLongo $raiz
        $inicioRelativo = $raizLonga.Length + 1
        $stats = @{ Arquivos = 0; Pastas = 0; Erros = 0; Bytes = [long]0; SomenteOnline = 0 }
        $pilha = New-Object 'System.Collections.Generic.Stack[System.IO.DirectoryInfo]'
        $pilha.Push((New-Object System.IO.DirectoryInfo ($raizLonga + '\')))
        $inicio = Get-Date

        Write-Host ''
        Write-Host ('[{0}/{1}] Analisando: {2}' -f $numero, $raizes.Count, $raiz)

        # Percorre em ordem de árvore: a pasta, seus arquivos e depois cada subpasta (em ordem alfabética)
        while ($pilha.Count -gt 0) {
            $dir = $pilha.Pop()
            if ($dir.FullName.Length -gt $inicioRelativo) { New-Linha $dir }  # a raiz em si não é listada

            try {
                # GetFileSystemInfos lê só a entrada do diretório (nome, tamanho, datas, atributos);
                # não abre os arquivos, então não dispara download do OneDrive.
                $itens = $dir.GetFileSystemInfos() | Sort-Object Name
            }
            catch {
                $stats.Erros++
                $rel = if ($dir.FullName.Length -gt $inicioRelativo) { $dir.FullName.Substring($inicioRelativo) } else { '.' }
                Write-Warning "Não foi possível ler '$rel': $($_.Exception.Message)"
                [pscustomobject]@{
                    PastaOrigem = $raiz; Tipo = 'ERRO'; CaminhoRelativo = $rel; Nome = $dir.Name; PastaRelativa = ''; Extensao = ''
                    Nivel = ''; TamanhoBytes = ''; Tamanho = ''; Criado = ''; Modificado = ''
                    Atributos = ''; StatusOneDrive = ''; Observacao = $_.Exception.Message
                }
                continue
            }

            $subpastas = New-Object System.Collections.Generic.List[System.IO.DirectoryInfo]
            foreach ($item in $itens) {
                if ($item -is [System.IO.DirectoryInfo]) { $subpastas.Add($item) }
                elseif ($item.FullName -ne $saidaLonga) { New-Linha $item }  # não lista o próprio arquivo de saída
            }
            for ($i = $subpastas.Count - 1; $i -ge 0; $i--) { $pilha.Push($subpastas[$i]) }
        }

        Write-Progress -Activity "Listando $raiz" -Completed
        Write-Host ('  Arquivos        : {0} ({1})' -f $stats.Arquivos, (Format-Tamanho $stats.Bytes))
        Write-Host ('    somente online: {0} (não foram baixados)' -f $stats.SomenteOnline)
        Write-Host ('  Pastas          : {0}' -f $stats.Pastas)
        Write-Host ('  Tempo           : {0:mm\:ss}' -f ((Get-Date) - $inicio))
        if ($stats.Erros) {
            Write-Host ('  Erros de leitura: {0} (linhas com Tipo = ERRO no arquivo; a listagem está incompleta)' -f $stats.Erros) -ForegroundColor Yellow
        }
        $resumo.Add([pscustomobject]@{ Pasta = $raiz; Arquivos = $stats.Arquivos; Pastas = $stats.Pastas; Erros = $stats.Erros })
    }
} | Export-Csv -LiteralPath $saida -Delimiter ';' -NoTypeInformation -Encoding UTF8

# Sem nenhum item (só pastas vazias) o Export-Csv não grava nem o cabeçalho
if (-not ($resumo | Where-Object { $_.Arquivos -or $_.Pastas -or $_.Erros })) {
    $colunas = 'PastaOrigem', 'Tipo', 'CaminhoRelativo', 'Nome', 'PastaRelativa', 'Extensao', 'Nivel',  # mesma ordem de New-Linha
               'TamanhoBytes', 'Tamanho', 'Criado', 'Modificado', 'Atributos', 'StatusOneDrive', 'Observacao'
    Set-Content -LiteralPath $saida -Value ('"' + ($colunas -join '";"') + '"') -Encoding UTF8
}

Write-Host ''
Write-Host ('Concluído: {0} pasta(s) analisada(s), {1} arquivo(s) e {2} subpasta(s) no total.' -f $resumo.Count,
    ($resumo | Measure-Object Arquivos -Sum).Sum, ($resumo | Measure-Object Pastas -Sum).Sum) -ForegroundColor Green
$errosTotal = ($resumo | Measure-Object Erros -Sum).Sum
if ($errosTotal) {
    Write-Host ('  {0} pasta(s) não puderam ser lidas (linhas com Tipo = ERRO); o inventário está incompleto.' -f $errosTotal) -ForegroundColor Yellow
}
Write-Host ('Arquivo gerado: {0}' -f $saida)
