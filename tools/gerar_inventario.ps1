<#
  Gera o INVENTARIO de uma pasta de databook: um .csv com o caminho relativo, o tamanho e
  a data de cada arquivo. So le nomes e tamanhos: nao abre nem altera nada.

  Feito para rodar onde a pasta esta (ex.: sessao remota do cliente no Windows App) e
  depois comparar no seu PC com tools\comparar_pastas.ps1 (-PastaB inventario.csv).

  Uso:  powershell -ExecutionPolicy Bypass -File gerar_inventario.ps1 -Pasta "T:\...\08. Databook"
  Sem -Pasta, usa a pasta atual do PowerShell.
#>
param(
    [string]$Pasta = (Get-Location).ProviderPath,
    [string]$Saida = ""
)
$ErrorActionPreference = "Stop"
$raiz = (Resolve-Path -LiteralPath $Pasta).ProviderPath.TrimEnd('\', '/')
if (-not $Saida) {
    $nome = ([System.IO.Path]::GetFileName($raiz) -replace '[^\w\-. ]', '_')
    $Saida = Join-Path ([Environment]::GetFolderPath("Desktop")) ("inventario_{0}_{1}.csv" -f $nome, (Get-Date -Format "yyyy-MM-dd_HHmm"))
}
$windows = ($null -eq $IsWindows -or $IsWindows)
$base = if (-not $windows) { $raiz } elseif ($raiz.StartsWith('\\')) { '\\?\UNC\' + $raiz.Substring(2) } else { '\\?\' + $raiz }
$corte = $base.Length + 1
$pilha = New-Object System.Collections.Stack
$pilha.Push($base)
$lista = New-Object System.Collections.ArrayList
$falhas = New-Object System.Collections.ArrayList
Write-Host "Listando $raiz ..."
while ($pilha.Count) {
    $d = $pilha.Pop()
    try {
        foreach ($i in (New-Object System.IO.DirectoryInfo($d)).EnumerateFileSystemInfos()) {
            if ($i -is [System.IO.DirectoryInfo]) { $pilha.Push($i.FullName); continue }
            [void]$lista.Add([pscustomobject]@{
                CaminhoRelativo = $i.FullName.Substring($corte)
                Tamanho         = $i.Length
                Modificado      = $i.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") })
            if ($lista.Count % 5000 -eq 0) { Write-Host "  $($lista.Count) arquivos..." }
        }
    } catch { [void]$falhas.Add($d.Replace($base, $raiz) + "  ->  " + $_.Exception.Message) }
}
$lista | Export-Csv -LiteralPath $Saida -Delimiter ';' -NoTypeInformation -Encoding UTF8
if ($falhas.Count) { $falhas | Set-Content -LiteralPath ($Saida -replace '\.csv$', '_pastas_sem_acesso.txt') -Encoding UTF8 }
$gb = [math]::Round((($lista | Measure-Object Tamanho -Sum).Sum) / 1GB, 1)
Write-Host ("Pronto: {0} arquivos ({1} GB), {2} pasta(s) sem acesso." -f $lista.Count, $gb, $falhas.Count) -ForegroundColor Green
Write-Host "Inventário: $Saida" -ForegroundColor Green
