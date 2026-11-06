<#
  Deixa a pasta do OneDrive IGUAL a pasta oficial (rede), em lotes, sem estourar o disco C:.

  Para cada lote (uma pasta de ate -LimiteGB, ou so os arquivos soltos de uma pasta maior):
    1. espera haver espaco livre suficiente (o lote + -FolgaGB);
    2. espelha com robocopy /MIR (copia o que falta, troca o que mudou e remove do OneDrive
       o que nao existe na rede - p.ex. pastas com nome abreviado);
    3. marca os arquivos do lote como "Liberar espaco": o OneDrive os envia para a nuvem
       e devolve o espaco do disco.
  Removidos do OneDrive vao para a Lixeira do OneDrive na web (recuperaveis).
  ANTES de rodar, salve os arquivos que so existem no OneDrive (sincronizar_pastas.ps1).

  Uso:
    powershell -ExecutionPolicy Bypass -File .\tools\espelhar_em_lotes.ps1              # so mostra o plano
    powershell -ExecutionPolicy Bypass -File .\tools\espelhar_em_lotes.ps1 -Executar    # executa
  Interrompido? Rode de novo: o que ja esta igual e pulado rapidamente.
#>
param(
    [string]$Origem  = "Y:\00-Obras\2024\24002 Eng. Proprietário UFV Irapuru MG  - Elera\4_Databooks",
    [string]$Destino = "$env:USERPROFILE\OneDrive\Databooks - Irapuru",
    [double]$LimiteGB = 8,
    [double]$FolgaGB = 6,
    [int]$MaxEsperaMin = 480,
    [switch]$Executar
)

$ErrorActionPreference = "Stop"
$GB = 1GB
$log = Join-Path "$env:USERPROFILE\Documents" ("espelho_databooks_{0}.log" -f (Get-Date -Format "yyyy-MM-dd_HHmm"))
$EXCLUIR = @("desktop.ini", "Thumbs.db", "~`$*")

function Get-CaminhoLongo([string]$p) {
    if ($null -ne $IsWindows -and -not $IsWindows) { return $p }
    if ($p.StartsWith("\\?\")) { return $p }
    $p = [System.IO.Path]::GetFullPath($p)
    if ($p.StartsWith("\\")) { return "\\?\UNC\" + $p.Substring(2) }
    return "\\?\" + $p
}
function Join-Rel([string]$base, [string]$rel) { if ($rel) { [System.IO.Path]::Combine($base, $rel) } else { $base } }
function Get-LivreGB {
    $raiz = [System.IO.Path]::GetPathRoot([System.IO.Path]::GetFullPath($Destino))
    return [math]::Round(([System.IO.DriveInfo]::new($raiz)).AvailableFreeSpace / $GB, 1)
}
function Test-Excluido([string]$nome) {
    foreach ($m in $EXCLUIR) { if ($nome -like $m) { return $true } }
    return $false
}

# ---- "Liberar espaco" (mesmo efeito do menu do Explorador; aceita caminhos longos) ----------
if ($null -eq $IsWindows -or $IsWindows) {
    Add-Type -Namespace Databook -Name Win32 -MemberDefinition @"
[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern uint GetFileAttributesW(string lpFileName);
[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern bool SetFileAttributesW(string lpFileName, uint dwFileAttributes);
"@
}
function Set-LiberarEspaco([string]$pasta, [bool]$recursivo) {
    $n = 0
    if ($null -ne $IsWindows -and -not $IsWindows) { return 0 }
    $pilha = New-Object System.Collections.Generic.Stack[string]
    $pilha.Push((Get-CaminhoLongo $pasta))
    while ($pilha.Count) {
        $d = $pilha.Pop()
        try { $itens = ([System.IO.DirectoryInfo]::new($d)).EnumerateFileSystemInfos() } catch { continue }
        foreach ($i in $itens) {
            if ($i -is [System.IO.DirectoryInfo]) { if ($recursivo) { $pilha.Push($i.FullName) }; continue }
            $a = [Databook.Win32]::GetFileAttributesW($i.FullName)
            if ($a -eq [uint32]::MaxValue) { continue }
            # mantem somente-leitura/oculto/sistema/arquivo; tira "manter neste dispositivo"; poe "liberar espaco"
            $novo = ($a -band 0x2127) -bor 0x100000
            if ([Databook.Win32]::SetFileAttributesW($i.FullName, [uint32]$novo)) { $n++ }
        }
    }
    return $n
}

# ---- 1) mapa da pasta oficial: tamanho por pasta --------------------------------------------
foreach ($p in @($Origem, $Destino)) { if (-not (Test-Path -LiteralPath $p)) { throw "Pasta não encontrada: $p" } }
Write-Host "Lendo a estrutura da pasta oficial (pode levar alguns minutos)..."
$raizO = (Get-CaminhoLongo $Origem).TrimEnd([System.IO.Path]::DirectorySeparatorChar)
$arqs = @{}; $filhos = @{}               # rel da pasta -> lista de (nome, tamanho) / subpastas
$pilha = New-Object System.Collections.Generic.Stack[string]
$pilha.Push("")
$total = 0
while ($pilha.Count) {
    $rel = $pilha.Pop()
    $arqs[$rel] = New-Object System.Collections.Generic.List[object]
    $filhos[$rel] = New-Object System.Collections.Generic.List[string]
    foreach ($i in ([System.IO.DirectoryInfo]::new((Join-Rel $raizO $rel))).EnumerateFileSystemInfos()) {
        $r = if ($rel) { [System.IO.Path]::Combine($rel, $i.Name) } else { $i.Name }
        if ($i -is [System.IO.DirectoryInfo]) { $filhos[$rel].Add($r); $pilha.Push($r) }
        elseif (-not (Test-Excluido $i.Name)) { $arqs[$rel].Add([pscustomobject]@{ Nome = $i.Name; Tam = $i.Length }); $total++ }
    }
}
$tamTotal = @{}
function Get-Tam([string]$rel) {
    if ($tamTotal.ContainsKey($rel)) { return $tamTotal[$rel] }
    $s = [int64]0
    foreach ($a in $arqs[$rel]) { $s += $a.Tam }
    foreach ($f in $filhos[$rel]) { $s += Get-Tam $f }
    $tamTotal[$rel] = $s
    return $s
}
[void](Get-Tam "")

# ---- 2) plano de lotes ---------------------------------------------------------------------
$lotes = New-Object System.Collections.Generic.List[object]
function Add-Lotes([string]$rel) {
    if ((Get-Tam $rel) -le $LimiteGB * $GB) {
        $lotes.Add([pscustomobject]@{ Rel = $rel; Tipo = "Pasta"; Bytes = (Get-Tam $rel) })
        return
    }
    # pasta grande: primeiro os arquivos soltos (e a limpeza de subpastas extras), depois cada subpasta
    $soltos = [int64]0; foreach ($a in $arqs[$rel]) { $soltos += $a.Tam }
    $lotes.Add([pscustomobject]@{ Rel = $rel; Tipo = "SoArquivos"; Bytes = $soltos })
    foreach ($f in ($filhos[$rel] | Sort-Object)) { Add-Lotes $f }
}
Add-Lotes ""
$maior = ($lotes | Measure-Object Bytes -Maximum).Maximum
Write-Host ("Pasta oficial: {0} arquivos, {1:N1} GB  |  {2} lotes, maior lote {3:N1} GB  |  livre no disco: {4} GB" -f `
    $total, ((Get-Tam "") / $GB), $lotes.Count, ($maior / $GB), (Get-LivreGB))
if ($maior -gt $LimiteGB * $GB) {
    Write-Host ("Atenção: há um único arquivo/pasta de {0:N1} GB acima do limite; ele será copiado sozinho." -f ($maior / $GB)) -ForegroundColor Yellow
}
if (-not $Executar) {
    $lotes | Select-Object @{n = "Lote"; e = { if ($_.Rel) { $_.Rel } else { "(raiz)" } } }, Tipo, @{n = "GB"; e = { [math]::Round($_.Bytes / $GB, 2) } } |
        Format-Table -AutoSize | Out-String -Width 300 | Write-Host
    Write-Host "Plano apenas (nada foi alterado). Para executar: acrescente -Executar" -ForegroundColor Yellow
    exit
}

# ---- 3) execucao ---------------------------------------------------------------------------
function Get-AFaltar($lote) {
    # bytes que ainda precisam vir da rede (arquivo ausente ou com tamanho diferente no OneDrive)
    $rels = New-Object System.Collections.Generic.List[string]
    if ($lote.Tipo -eq "SoArquivos") { $rels.Add($lote.Rel) }
    else {
        $pend = New-Object System.Collections.Generic.Stack[string]
        $pend.Push($lote.Rel)
        while ($pend.Count -gt 0) {
            $x = $pend.Pop()
            $rels.Add($x)
            if ($filhos.ContainsKey($x)) { foreach ($f in $filhos[$x]) { if ($null -ne $f) { $pend.Push($f) } } }
        }
    }
    $s = [int64]0
    foreach ($r in $rels) {
        if (-not $arqs.ContainsKey($r)) { continue }
        foreach ($a in $arqs[$r]) {
            $dst = Get-CaminhoLongo ([System.IO.Path]::Combine((Join-Rel $Destino $r), $a.Nome))
            $fi = [System.IO.FileInfo]::new($dst)
            if (-not $fi.Exists -or $fi.Length -ne $a.Tam) { $s += $a.Tam }
        }
    }
    return $s
}

$resp = Read-Host "Espelhar $($lotes.Count) lote(s) da rede para o OneDrive? Arquivos que não existem na rede serão REMOVIDOS do OneDrive (vão para a lixeira do OneDrive). Já salvou os exclusivos? Digite SIM"
if ($resp -ne "SIM") { Write-Host "Cancelado. Nada foi alterado."; exit }
"Início: $(Get-Date)  Origem: $Origem  Destino: $Destino" | Out-File -LiteralPath $log -Encoding UTF8

$i = 0
foreach ($lote in $lotes) {
    $i++
    $nome = if ($lote.Rel) { $lote.Rel } else { "(raiz)" }
    $falta = Get-AFaltar $lote
    $precisa = [math]::Round($falta / $GB + $FolgaGB, 1)
    $inicio = Get-Date
    while ((Get-LivreGB) -lt $precisa) {
        if (((Get-Date) - $inicio).TotalMinutes -gt $MaxEsperaMin) {
            Write-Host "Esperei $MaxEsperaMin min e o OneDrive não liberou espaço. Verifique o OneDrive e rode de novo." -ForegroundColor Red
            exit 1
        }
        Write-Host ("  aguardando o OneDrive liberar espaço: livre {0} GB, necessário {1} GB..." -f (Get-LivreGB), $precisa)
        Start-Sleep -Seconds 60
    }
    Write-Host ("[{0}/{1}] {2}  ({3}; {4:N2} GB a copiar; livre {5} GB)" -f $i, $lotes.Count, $nome, $lote.Tipo, ($falta / $GB), (Get-LivreGB)) -ForegroundColor Cyan
    $src = Join-Rel $Origem $lote.Rel
    $dst = Join-Rel $Destino $lote.Rel
    $rcArgs = @($src, $dst, "/MIR", "/COPY:DAT", "/DCOPY:T", "/XJ", "/R:1", "/W:2", "/MT:8", "/NP", "/NFL", "/NDL", "/LOG+:$log", "/XF") + $EXCLUIR
    if ($lote.Tipo -eq "SoArquivos") { $rcArgs += "/LEV:1" }
    & robocopy @rcArgs | Out-Null
    if ($LASTEXITCODE -ge 8) {
        Write-Host "robocopy falhou (código $LASTEXITCODE) no lote $nome. Veja o log: $log" -ForegroundColor Red
        exit 1
    }
    $lib = Set-LiberarEspaco $dst ($lote.Tipo -eq "Pasta")
    "$(Get-Date)  lote ok: $nome  (liberar espaço em $lib arquivos)" | Out-File -LiteralPath $log -Append -Encoding UTF8
}
Write-Host "Concluído. Log: $log" -ForegroundColor Green
Write-Host "O OneDrive ainda pode estar enviando os últimos lotes; deixe-o terminar antes de desligar o PC."
