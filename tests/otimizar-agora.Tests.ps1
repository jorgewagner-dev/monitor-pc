# Executar com Windows PowerShell 5.1. Nao fecha processos nem altera energia.
$ErrorActionPreference = 'Stop'
$source = Join-Path (Split-Path $PSScriptRoot -Parent) 'MonitorPC.ps1'
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$null, [ref]$errors)
if ($errors.Count) { throw 'Erro de sintaxe no aplicativo.' }
foreach ($name in 'Test-AgoraProcesso','Restore-AgoraSessao') {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    . ([scriptblock]::Create($node.Extent.Text))
}
function Assert($value, $message) { if (-not $value) { throw $message } }
$AgoraPermitidos = @('Spotify')
$script:JogoAtual = ''
$stamp = (Get-Date).ToUniversalTime()
$fake = [pscustomobject]@{ Id=123456; ProcessName='Spotify'; StartTime=$stamp; Path='C:\Apps\Spotify.exe'; MainWindowHandle=[IntPtr]1 }
$item = [pscustomobject]@{ Id=123456; Nome='Spotify'; Inicio=$stamp.Ticks; Caminho=$fake.Path }
Assert (Test-AgoraProcesso $fake $item) 'Aplicativo conhecido deveria ser elegivel.'
$fake.StartTime = $stamp.AddSeconds(1)
Assert (-not (Test-AgoraProcesso $fake $item)) 'PID reutilizado deve ser recusado.'
$fake.StartTime = $stamp; $fake.Path = Join-Path $env:windir 'Spotify.exe'
Assert (-not (Test-AgoraProcesso $fake $item)) 'Executavel no Windows deve ser recusado.'
$fake.Path = $item.Caminho; $script:JogoAtual = 'Spotify'
Assert (-not (Test-AgoraProcesso $fake $item)) 'Jogo atual deve ser recusado.'
$script:JogoAtual = ''; $fake.MainWindowHandle = [IntPtr]::Zero
Assert (-not (Test-AgoraProcesso $fake $item)) 'Processo sem janela deve ser recusado.'

# Dubles: restauracao nao executa comandos reais.
$script:changes = @(); $script:opened = @(); $script:removed = $false
$script:active = '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'
function Get-AgoraPlano { $script:active }
function powercfg { $script:changes += $args[1]; $global:LASTEXITCODE=0 }
function Get-Process { param($Name,$ErrorAction) $null }
function Test-Path { param($LiteralPath,$PathType) $true }
function Start-Process { param($FilePath,$WindowStyle,$ErrorAction) $script:opened += $FilePath }
function Remove-Item { param($LiteralPath,$ErrorAction) $script:removed=$true }
$ui = @{ btnAgoraRestaurar=[pscustomobject]@{IsEnabled=$true}; txtAgoraStatus=[pscustomobject]@{Text=''} }
$AgoraPath = 'unused-test-state.json'
function New-TestSession {
    [pscustomobject]@{ PlanoAnterior='381b4222-f694-41f0-9685-ff5bb260df2e'; PlanoAplicado='8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'; Aplicativos=@($item) }
}
$script:AgoraSessao = New-TestSession
Restore-AgoraSessao
Assert ($script:changes.Count -eq 1) 'Plano anterior deveria ser restaurado.'
Assert ($script:opened.Count -eq 1) 'Aplicativo fechado deveria reabrir.'
Assert ($script:removed -and -not $script:AgoraSessao) 'Estado recuperado deveria ser limpo.'
$script:active='a1841308-3541-4fab-bc81-f71556f20b4a'
$script:changes=@(); $script:AgoraSessao=New-TestSession
Restore-AgoraSessao
Assert ($script:changes.Count -eq 0) 'Uma troca externa de plano deve ser respeitada.'

function Start-Process { param($FilePath,$WindowStyle,$ErrorAction) throw 'falha simulada' }
$script:removed=$false; $script:AgoraSessao=New-TestSession
Restore-AgoraSessao
Assert ($script:AgoraSessao -and -not $script:removed) 'Recuperacao parcial deve manter o registro para tentar novamente.'
'PASS: identidade, processos protegidos, restauracao, plano externo e recuperacao parcial.'
