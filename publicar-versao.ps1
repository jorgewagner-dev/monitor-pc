<#
    Publica uma versão nova do Monitor PC para a atualização automática.
    Uso:  .\publicar-versao.ps1 -Versao 2.6.0 -Notas "o que mudou"
    Faz: troca o número da versão no MonitorPC.ps1, confere se o arquivo não tem erro, calcula a impressão
    digital (SHA-256), grava o versao.json, refaz o MonitorPC.zip e envia tudo para o GitHub com a tag vX.Y.Z.
#>
param(
    [Parameter(Mandatory)][string]$Versao,
    [string]$Notas = ''
)
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
$Repositorio = 'jorgewagner-dev/monitor-pc'

$nova = [version]$Versao
$arq = Join-Path $PSScriptRoot 'MonitorPC.ps1'
$texto = [IO.File]::ReadAllText($arq, [Text.Encoding]::UTF8)
$atual = [version]([regex]::Match($texto, "(?m)^\`$VersaoApp = '([^']*)'").Groups[1].Value)
if ($nova -le $atual) { throw "A versão nova ($nova) precisa ser maior que a atual ($atual)." }

$texto = [regex]::Replace($texto, "(?m)^\`$VersaoApp = '[^']*'", "`$VersaoApp = '$Versao'")
[IO.File]::WriteAllText($arq, $texto, (New-Object Text.UTF8Encoding $true))
$erros = $null
[void][Management.Automation.Language.Parser]::ParseFile($arq, [ref]$null, [ref]$erros)
if ($erros.Count) { throw "O MonitorPC.ps1 tem erro de sintaxe na linha $($erros[0].Extent.StartLineNumber): $($erros[0].Message)" }

$hash = (Get-FileHash $arq -Algorithm SHA256).Hash
$info = [ordered]@{
    versao  = $Versao
    data    = (Get-Date -Format 'yyyy-MM-dd')
    notas   = $Notas
    # Aponta para a tag da versão: o arquivo dessa URL nunca muda, então a impressão digital sempre confere
    arquivo = "https://raw.githubusercontent.com/$Repositorio/v$Versao/MonitorPC.ps1"
    sha256  = $hash
}
[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'versao.json'), ($info | ConvertTo-Json), (New-Object Text.UTF8Encoding $false))

Compress-Archive -Path $arq, (Join-Path $PSScriptRoot 'Iniciar Monitor PC.bat'), (Join-Path $PSScriptRoot 'LEIA-ME.txt') `
    -DestinationPath (Join-Path $PSScriptRoot 'MonitorPC.zip') -Force

git add MonitorPC.ps1 versao.json LEIA-ME.txt 'Iniciar Monitor PC.bat' README.md publicar-versao.ps1 .gitignore .gitattributes
git commit -m "Versão $Versao$(if ($Notas) { ": $Notas" })"
git tag "v$Versao"
git push origin main --tags
Write-Host "Versão $Versao publicada. Os Monitores dos seus amigos vão oferecer a atualização em até 24 h (ou na hora, em Aparência > Verificar atualização)." -ForegroundColor Green
