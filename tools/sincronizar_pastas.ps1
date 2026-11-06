<#
  Sincroniza duas pastas de databooks copiando SO o que falta de cada lado, a partir do
  comparativo gerado por comparar_pastas.ps1 (arquivo 2_Diferencas_detalhadas.csv).

  Seguranca:
  - Nunca sobrescreve e nunca apaga: se o arquivo ja existir no destino, ele e pulado.
  - "Tamanho diferente" NAO e copiado (exige decidir qual versao vale): fica no relatorio.
  - Arquivos que existem do outro lado em OUTRA pasta (provavelmente movidos) NAO sao
    copiados, para nao duplicar. Use -IncluirMovidos para copiar mesmo assim.
  - Sem -Executar, apenas SIMULA e gera o plano (nada e copiado).

  Uso:
    powershell -ExecutionPolicy Bypass -File .\tools\sincronizar_pastas.ps1              # simula
    powershell -ExecutionPolicy Bypass -File .\tools\sincronizar_pastas.ps1 -Executar    # copia
  Opcional: -Sentido "AparaB" | "BparaA" | "Ambos" (padrao Ambos)
#>
param(
    [string]$PastaA = "$env:USERPROFILE\OneDrive\Databooks - Irapuru",
    [string]$PastaB = "Y:\00-Obras\2024\24002 Eng. Proprietário UFV Irapuru MG  - Elera\4_Databooks",
    [string]$NomeA  = "OneDrive",
    [string]$NomeB  = "Rede",
    [string]$Comparativo = "",
    [ValidateSet("Ambos", "AparaB", "BparaA")][string]$Sentido = "Ambos",
    [switch]$IncluirMovidos,
    [switch]$Executar
)

$ErrorActionPreference = "Stop"

function Get-CaminhoLongo([string]$p) {
    if ($null -ne $IsWindows -and -not $IsWindows) { return $p }   # (testes fora do Windows)
    if ($p.StartsWith("\\?\")) { return $p }
    $p = [System.IO.Path]::GetFullPath($p)
    if ($p.StartsWith("\\")) { return "\\?\UNC\" + $p.Substring(2) }
    return "\\?\" + $p
}

# comparativo: o informado ou o mais recente em Documentos
if (-not $Comparativo) {
    $ultimo = Get-ChildItem -LiteralPath "$env:USERPROFILE\Documents" -Directory -Filter "Comparativo_Databooks_*" |
              Sort-Object Name -Descending | Select-Object -First 1
    if (-not $ultimo) { throw "Nenhum comparativo encontrado. Rode antes: tools\comparar_pastas.ps1" }
    $Comparativo = $ultimo.FullName
}
$csv = Join-Path $Comparativo "2_Diferencas_detalhadas.csv"
if (-not (Test-Path -LiteralPath $csv)) { throw "Arquivo não encontrado: $csv" }
foreach ($p in @($PastaA, $PastaB)) {
    if (-not (Test-Path -LiteralPath $p)) { throw "Pasta não encontrada: $p" }
}

$linhas = Import-Csv -LiteralPath $csv -Delimiter ';' -Encoding UTF8
$colMovido = "Mesmo arquivo em outro lugar"
Write-Host "Comparativo: $Comparativo"
Write-Host "Modo: $(if ($Executar) { 'EXECUTAR (copia de verdade)' } else { 'SIMULAÇÃO (nada será copiado)' })" -ForegroundColor Yellow

$plano = New-Object System.Collections.Generic.List[object]
foreach ($l in $linhas) {
    $origem = $null
    if ($l.Situacao -eq "Falta em $NomeB" -and $Sentido -ne "BparaA") {
        $origem = $PastaA; $destino = $PastaB; $direcao = "$NomeA -> $NomeB"
    } elseif ($l.Situacao -eq "Falta em $NomeA" -and $Sentido -ne "AparaB") {
        $origem = $PastaB; $destino = $PastaA; $direcao = "$NomeB -> $NomeA"
    } else { continue }                     # "Tamanho diferente" ou sentido nao escolhido
    if ($l.$colMovido -and -not $IncluirMovidos) {
        $plano.Add([pscustomobject]@{ Direcao = $direcao; Resultado = "Pulado: existe em outra pasta"
            CaminhoRelativo = $l.CaminhoRelativo; Detalhe = $l.$colMovido })
        continue
    }
    $plano.Add([pscustomobject]@{ Direcao = $direcao; Resultado = "A copiar"
        CaminhoRelativo = $l.CaminhoRelativo; Detalhe = ""; Origem = $origem; Destino = $destino })
}

$aCopiar = @($plano | Where-Object Resultado -eq "A copiar")
Write-Host ("Arquivos a copiar: {0}  |  pulados por existirem em outra pasta: {1}" -f `
    $aCopiar.Count, @($plano | Where-Object Resultado -like "Pulado*").Count)
foreach ($g in ($aCopiar | Group-Object Direcao)) { Write-Host ("  {0}: {1} arquivo(s)" -f $g.Name, $g.Count) }

if ($Executar -and $aCopiar.Count) {
    $resp = Read-Host "Confirma a cópia de $($aCopiar.Count) arquivo(s)? Digite SIM para continuar"
    if ($resp -ne "SIM") { Write-Host "Cancelado. Nada foi copiado."; exit }
    $i = 0
    foreach ($item in $aCopiar) {
        $i++
        $src = Get-CaminhoLongo (Join-Path $item.Origem  $item.CaminhoRelativo)
        $dst = Get-CaminhoLongo (Join-Path $item.Destino $item.CaminhoRelativo)
        try {
            if ([System.IO.File]::Exists($dst)) {
                $item.Resultado = "Pulado: já existe no destino"
            } else {
                [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($dst))
                [System.IO.File]::Copy($src, $dst, $false)        # $false = nunca sobrescreve
                $item.Resultado = "Copiado"
            }
        } catch {
            $item.Resultado = "Erro"
            $item.Detalhe = $_.Exception.Message
        }
        if ($i % 50 -eq 0) { Write-Host "  $i / $($aCopiar.Count) ..." }
    }
}

$nome = if ($Executar) { "5_Sincronizacao_resultado.csv" } else { "4_Sincronizacao_plano.csv" }
$plano | Select-Object Direcao, Resultado, CaminhoRelativo, Detalhe |
    Export-Csv -LiteralPath (Join-Path $Comparativo $nome) -Delimiter ';' -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "Resumo:" -ForegroundColor Cyan
foreach ($g in ($plano | Group-Object { "$($_.Direcao)  |  $($_.Resultado)" } | Sort-Object Name)) {
    Write-Host ("  {0,6}  {1}" -f $g.Count, $g.Name)
}
Write-Host "Relatório: $(Join-Path $Comparativo $nome)" -ForegroundColor Green
if (-not $Executar) {
    Write-Host "Confira o plano. Para copiar de verdade, rode de novo com -Executar." -ForegroundColor Yellow
}
