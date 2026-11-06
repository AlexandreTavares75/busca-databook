<#
  Compara duas pastas de databooks (ex.: OneDrive x rede do cliente) e gera planilhas
  com os arquivos que faltam de cada lado e os que tem tamanho diferente.

  - So LISTA os arquivos: nao abre, nao copia, nao altera nada.
  - Nao baixa arquivos do OneDrive (le apenas nome e tamanho).
  - Funciona com caminhos acima de 260 caracteres.

  Cada lado pode ser uma PASTA ou um INVENTARIO (.csv gerado por gerar_inventario.ps1,
  p.ex. na sessao remota de um cliente acessado pelo Windows App).

  Uso (PowerShell):
    powershell -ExecutionPolicy Bypass -File .\tools\comparar_pastas.ps1
  ou informando as pastas:
    powershell -ExecutionPolicy Bypass -File .\tools\comparar_pastas.ps1 -PastaA "C:\...\Databooks - Irapuru" -PastaB "Y:\...\4_Databooks"
#>
param(
    [string]$PastaA = "$env:USERPROFILE\OneDrive\Databooks - Irapuru",
    [string]$PastaB = "Y:\00-Obras\2024\24002 Eng. Proprietário UFV Irapuru MG  - Elera\4_Databooks",
    [string]$NomeA  = "OneDrive",
    [string]$NomeB  = "Rede",
    [string]$Saida  = "$env:USERPROFILE\Documents\Comparativo_Databooks_$(Get-Date -Format 'yyyy-MM-dd_HHmm')"
)

$ErrorActionPreference = "Stop"

function Get-CaminhoLongo([string]$p) {
    if ($null -ne $IsWindows -and -not $IsWindows) { return $p }   # (testes fora do Windows)
    if ($p.StartsWith("\\?\")) { return $p }
    $p = [System.IO.Path]::GetFullPath($p)
    if ($p.StartsWith("\\"))   { return "\\?\UNC\" + $p.Substring(2) }
    return "\\?\" + $p
}

function Get-Inventario([string]$csv, [string]$rotulo) {
    # Inventario: colunas CaminhoRelativo;Tamanho (Modificado e opcional)
    $tab = @{}
    foreach ($l in (Import-Csv -LiteralPath $csv -Delimiter ';' -Encoding UTF8)) {
        if (-not $l.CaminhoRelativo) { continue }
        $sep = [System.IO.Path]::DirectorySeparatorChar
        $rel = $l.CaminhoRelativo.Replace('/', $sep).Replace('\', $sep).TrimStart($sep)
        if ([System.IO.Path]::GetFileName($rel).StartsWith('~$')) { continue }
        $tab[$rel.ToLowerInvariant()] = [pscustomobject]@{ Rel = $rel; Tamanho = [int64]$l.Tamanho }
    }
    Write-Host "  $rotulo : $($tab.Count) arquivos no inventário $csv"
    return @{ Arquivos = $tab; Falhas = (New-Object System.Collections.Generic.List[string]) }
}

function Get-Arquivos([string]$raiz, [string]$rotulo) {
    if ($raiz -like "*.csv" -and (Test-Path -LiteralPath $raiz -PathType Leaf)) { return Get-Inventario $raiz $rotulo }
    # Lista recursiva tolerante a erros: uma pasta sem acesso nao interrompe a varredura.
    $raizLonga = (Get-CaminhoLongo $raiz).TrimEnd([System.IO.Path]::DirectorySeparatorChar)
    $tab = @{}
    $pilha = New-Object System.Collections.Generic.Stack[string]
    $pilha.Push($raizLonga)
    $falhas = New-Object System.Collections.Generic.List[string]
    $n = 0
    while ($pilha.Count -gt 0) {
        $dir = $pilha.Pop()
        try {
            $info = New-Object System.IO.DirectoryInfo($dir)
            foreach ($item in $info.EnumerateFileSystemInfos()) {
                if ($item -is [System.IO.DirectoryInfo]) {
                    $pilha.Push($item.FullName)
                } else {
                    if ($item.Name.StartsWith('~$')) { continue }      # temporario do Office
                    $rel = $item.FullName.Substring($raizLonga.Length + 1)
                    $tab[$rel.ToLowerInvariant()] = [pscustomobject]@{ Rel = $rel; Tamanho = $item.Length }
                    $n++
                    if ($n % 2000 -eq 0) { Write-Host "  $rotulo : $n arquivos..." }
                }
            }
        } catch {
            $falhas.Add($dir.Replace($raizLonga, $raiz) + "  ->  " + $_.Exception.Message)
        }
    }
    Write-Host "  $rotulo : $n arquivos listados ($($falhas.Count) pasta(s) sem acesso)"
    return @{ Arquivos = $tab; Falhas = $falhas }
}

function Get-Empresa([string]$rel) {
    $sep = [System.IO.Path]::DirectorySeparatorChar
    if ($rel.Contains($sep)) { return $rel.Split($sep)[0] } else { return "(raiz)" }
}

foreach ($p in @($PastaA, $PastaB)) {
    if (-not (Test-Path -LiteralPath $p)) { throw "Pasta ou inventário não encontrado: $p" }
}
New-Item -ItemType Directory -Force -Path $Saida | Out-Null

Write-Host "Listando $NomeA ..."; $ListaA = Get-Arquivos $PastaA $NomeA
Write-Host "Listando $NomeB ..."; $ListaB = Get-Arquivos $PastaB $NomeB

# indice por nome de arquivo + tamanho, para sugerir onde um arquivo "faltante" esta do outro lado
function New-IndiceNome($tab) {
    $idx = @{}
    foreach ($v in $tab.Values) {
        $k = ([System.IO.Path]::GetFileName($v.Rel)).ToLowerInvariant() + "|" + $v.Tamanho
        if (-not $idx.ContainsKey($k)) { $idx[$k] = $v.Rel }
    }
    return $idx
}
$idxA = New-IndiceNome $ListaA.Arquivos
$idxB = New-IndiceNome $ListaB.Arquivos

$detalhe = New-Object System.Collections.Generic.List[object]
foreach ($k in $ListaA.Arquivos.Keys) {
    $a = $ListaA.Arquivos[$k]
    if ($ListaB.Arquivos.ContainsKey($k)) {
        $b = $ListaB.Arquivos[$k]
        if ($a.Tamanho -ne $b.Tamanho) {
            $detalhe.Add([pscustomobject]@{ Situacao = "Tamanho diferente"; Empresa = Get-Empresa $a.Rel
                Arquivo = [System.IO.Path]::GetFileName($a.Rel); CaminhoRelativo = $a.Rel
                "Tamanho_$NomeA (KB)" = [math]::Round($a.Tamanho / 1KB, 1); "Tamanho_$NomeB (KB)" = [math]::Round($b.Tamanho / 1KB, 1)
                "Mesmo arquivo em outro lugar" = "" })
        }
    } else {
        $chave = ([System.IO.Path]::GetFileName($a.Rel)).ToLowerInvariant() + "|" + $a.Tamanho
        $detalhe.Add([pscustomobject]@{ Situacao = "Falta em $NomeB"; Empresa = Get-Empresa $a.Rel
            Arquivo = [System.IO.Path]::GetFileName($a.Rel); CaminhoRelativo = $a.Rel
            "Tamanho_$NomeA (KB)" = [math]::Round($a.Tamanho / 1KB, 1); "Tamanho_$NomeB (KB)" = ""
            "Mesmo arquivo em outro lugar" = $(if ($idxB.ContainsKey($chave)) { $idxB[$chave] } else { "" }) })
    }
}
foreach ($k in $ListaB.Arquivos.Keys) {
    if (-not $ListaA.Arquivos.ContainsKey($k)) {
        $b = $ListaB.Arquivos[$k]
        $chave = ([System.IO.Path]::GetFileName($b.Rel)).ToLowerInvariant() + "|" + $b.Tamanho
        $detalhe.Add([pscustomobject]@{ Situacao = "Falta em $NomeA"; Empresa = Get-Empresa $b.Rel
            Arquivo = [System.IO.Path]::GetFileName($b.Rel); CaminhoRelativo = $b.Rel
            "Tamanho_$NomeA (KB)" = ""; "Tamanho_$NomeB (KB)" = [math]::Round($b.Tamanho / 1KB, 1)
            "Mesmo arquivo em outro lugar" = $(if ($idxA.ContainsKey($chave)) { $idxA[$chave] } else { "" }) })
    }
}

# resumo por empresa (primeira pasta: DB_...)
$empresas = @{}
foreach ($t in @(@($ListaA.Arquivos, "A"), @($ListaB.Arquivos, "B"))) {
    foreach ($v in $t[0].Values) {
        $e = Get-Empresa $v.Rel
        if (-not $empresas.ContainsKey($e)) { $empresas[$e] = @{ A = 0; B = 0; FaltaA = 0; FaltaB = 0; Dif = 0; Movidos = 0 } }
        $empresas[$e][$t[1]]++
    }
}
foreach ($d in $detalhe) {
    $r = $empresas[$d.Empresa]
    switch -Wildcard ($d.Situacao) {
        "Falta em $NomeB" { $r.FaltaB++ }
        "Falta em $NomeA" { $r.FaltaA++ }
        "Tamanho*"        { $r.Dif++ }
    }
    if ($d."Mesmo arquivo em outro lugar") { $r.Movidos++ }
}
$resumo = foreach ($e in ($empresas.Keys | Sort-Object)) {
    $r = $empresas[$e]
    [pscustomobject]@{ Empresa = $e; "Arquivos $NomeA" = $r.A; "Arquivos $NomeB" = $r.B
        "Faltam em $NomeB" = $r.FaltaB; "Faltam em $NomeA" = $r.FaltaA; "Tamanho diferente" = $r.Dif
        "Provavelmente movidos (mesmo nome e tamanho em outra pasta)" = $r.Movidos }
}

# CSV com ";" e UTF-8 com BOM: abre direto no Excel em portugues, com acentos corretos
$resumo  | Export-Csv -LiteralPath (Join-Path $Saida "1_Resumo_por_empresa.csv") -Delimiter ';' -NoTypeInformation -Encoding UTF8
$detalhe | Sort-Object Situacao, Empresa, CaminhoRelativo |
    Export-Csv -LiteralPath (Join-Path $Saida "2_Diferencas_detalhadas.csv") -Delimiter ';' -NoTypeInformation -Encoding UTF8
$falhas = @($ListaA.Falhas | ForEach-Object { "$NomeA`t$_" }) + @($ListaB.Falhas | ForEach-Object { "$NomeB`t$_" })
if ($falhas.Count) { $falhas | Set-Content -LiteralPath (Join-Path $Saida "3_Pastas_sem_acesso.txt") -Encoding UTF8 }

Write-Host ""
Write-Host "Resumo por empresa:" -ForegroundColor Cyan
$resumo | Format-Table -AutoSize | Out-String -Width 300 | Write-Host
Write-Host ("Total: {0} faltam em {1}, {2} faltam em {3}, {4} com tamanho diferente." -f `
    ($detalhe | Where-Object Situacao -eq "Falta em $NomeB").Count, $NomeB, `
    ($detalhe | Where-Object Situacao -eq "Falta em $NomeA").Count, $NomeA, `
    ($detalhe | Where-Object Situacao -eq "Tamanho diferente").Count)
Write-Host "Planilhas salvas em: $Saida" -ForegroundColor Green
if ($null -eq $IsWindows -or $IsWindows) { Invoke-Item -LiteralPath $Saida }
