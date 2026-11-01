#Requires -Version 5.1
<#
    Monitor PC - gerenciador de tarefas com alertas, vida útil das peças e otimização.
    Execute pelo arquivo "Iniciar Monitor PC.bat".
#>
param([switch]$Minimizado, [switch]$Reiniciando, [switch]$Manutencao)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing, Microsoft.VisualBasic

#region Uma cópia só
# Se o Monitor já estiver aberto, pede para ele aparecer e sai (evita vários ícones na bandeja e alertas repetidos).
# -Reiniciando: a cópia anterior está fechando (troca de tema / abrir como admin), então espera ela liberar.
$JaAberto = $false
if ($Manutencao) { } else {   # a manutenção agendada roda à parte, mesmo com o Monitor aberto
try {
    $Trava = New-Object System.Threading.Mutex($false, 'Local\MonitorPC_Instancia')
    try { $JaAberto = -not $Trava.WaitOne($(if ($Reiniciando) { 15000 } else { 0 })) }
    catch [System.Threading.AbandonedMutexException] { $JaAberto = $false }   # a anterior fechou sem liberar
} catch [System.UnauthorizedAccessException] {
    $JaAberto = $true   # aberto como administrador e esta cópia não é admin
}
if ($JaAberto) {
    $avisou = $false
    try { [void][System.Threading.EventWaitHandle]::OpenExisting('Local\MonitorPC_Mostrar').Set(); $avisou = $true } catch {}
    if (-not $avisou) {
        [void][Windows.MessageBox]::Show('O Monitor PC já está aberto (como administrador). Procure o ícone dele na bandeja, perto do relógio.', 'Monitor PC')
    }
    exit
}
$EventoMostrar = New-Object System.Threading.EventWaitHandle($false, 'AutoReset', 'Local\MonitorPC_Mostrar')
}
#endregion

$AppNome   = 'Monitor PC'
$VersaoApp = '2.5.0'
# Endereço do arquivo versao.json publicado (vazio = atualização automática desligada)
$UrlAtualizacaoPadrao = 'https://raw.githubusercontent.com/jorgewagner-dev/monitor-pc/main/versao.json'
$Intervalo = 1      # segundos entre leituras
$HistMax   = 120    # pontos nos gráficos (120 x 1s = 2 minutos)
$ConfigDir  = Join-Path $env:APPDATA 'MonitorPC'
$ConfigPath = Join-Path $ConfigDir 'config.json'
$LogPath    = Join-Path $ConfigDir 'alertas.log'
if (-not (Test-Path $ConfigDir)) { New-Item -ItemType Directory -Path $ConfigDir | Out-Null }

$EhAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)

# Paleta
$Cor = @{
    Lima       = '#A6F23A'
    LimaMedia  = '#7BC62D'
    LimaEscura = '#4E9A1E'
    Trilha     = '#1C2418'
    Ambar      = '#FFB020'
    Vermelho   = '#FF5A4D'
    Estimado   = '#5ED6B5'
    Suave      = '#8E8E96'
}

#region Configuração
function Get-ConfigPadrao {
    [pscustomobject]@{
        AlertasAtivos      = $true
        Som                = $true
        CpuLimite          = 90
        RamLimite          = 85
        DiscoLimite        = 95
        EspacoMinimo       = 10
        GpuTempLimite      = 83
        CorDestaque        = '#A6F23A'
        Brilho             = $true
        IconeCor           = ''
        MedirFps           = $true
        ModoJogo           = $false
        ModoJogoPlano      = $true
        ModoJogoIndexacao  = $true
        ModoJogoReabrir    = $true
        ModoJogoFechar     = @()
        PingAlvo           = '1.1.1.1'
        ManutAtiva         = $false
        ManutDia           = 'Sunday'
        ManutHora          = '12:00'
        ManutTemp          = $true
        ManutLixeira       = $false
        ManutSistema       = $false
        Overlay            = $false
        OverlayCanto       = 'SuperiorEsquerdo'
        UrlAtualizacao     = ''
        Sustentar          = 10
        IntervaloAlertaMin = 5
        AnoCompra          = ''
    }
}

function Read-Config {
    $cfg = Get-ConfigPadrao
    if (Test-Path $ConfigPath) {
        try {
            $salvo = Get-Content $ConfigPath -Raw | ConvertFrom-Json
            foreach ($p in $salvo.PSObject.Properties) {
                if ($cfg.PSObject.Properties[$p.Name]) { $cfg.($p.Name) = $p.Value }
            }
        } catch {}
    }
    $cfg
}

function Save-Config { $script:Config | ConvertTo-Json | Set-Content -Path $ConfigPath -Encoding UTF8 }

$Config = Read-Config
#endregion

#region Manutenção (o que a tarefa agendada executa)
$ManutLogPath = Join-Path $ConfigDir 'manutencao.log'

# Apaga arquivos com mais de N dias numa pasta; devolve os bytes liberados (arquivos em uso são pulados)
function Clear-Antigos([string]$pasta, [int]$dias) {
    if (-not (Test-Path -LiteralPath $pasta)) { return 0 }
    $limite = (Get-Date).AddDays(-$dias)
    $liberado = 0
    foreach ($f in Get-ChildItem -LiteralPath $pasta -Recurse -Force -File -ErrorAction SilentlyContinue) {
        if ($f.LastWriteTime -gt $limite) { continue }
        try { $tam = $f.Length; Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop; $liberado += $tam } catch {}
    }
    # Pastas que ficaram vazias
    Get-ChildItem -LiteralPath $pasta -Recurse -Force -Directory -ErrorAction SilentlyContinue | Sort-Object { $_.FullName.Length } -Descending |
        Where-Object { $_.LastWriteTime -lt $limite -and -not (Get-ChildItem -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue) } |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
    $liberado
}

# Lixeira: só o que foi apagado há mais de N dias
function Clear-LixeiraAntiga([int]$dias) {
    $liberado = 0
    try {
        $limite = (Get-Date).AddDays(-$dias)
        $lixeira = (New-Object -ComObject Shell.Application).Namespace(10)
        foreach ($item in @($lixeira.Items())) {
            $quando = $item.ExtendedProperty('System.Recycle.DateDeleted')
            if (-not $quando -or [datetime]$quando -gt $limite) { continue }
            $r = $item.Path
            $i = Join-Path (Split-Path $r) ((Split-Path $r -Leaf) -replace '^\$R', '$I')
            try {
                $tam = if (Test-Path -LiteralPath $r -PathType Leaf) { (Get-Item -LiteralPath $r -Force).Length } else { (Get-ChildItem -LiteralPath $r -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum }
                Remove-Item -LiteralPath $r -Recurse -Force -ErrorAction Stop
                Remove-Item -LiteralPath $i -Force -ErrorAction SilentlyContinue
                $liberado += $tam
            } catch {}
        }
    } catch {}
    $liberado
}

if ($Manutencao) {
    $partes = @(); $total = 0
    if ($Config.ManutTemp) { $b = Clear-Antigos $env:TEMP 2; $total += $b; $partes += 'temporários: {0:N0} MB' -f ($b / 1MB) }
    if ($Config.ManutLixeira) { $b = Clear-LixeiraAntiga 30; $total += $b; $partes += 'Lixeira antiga: {0:N0} MB' -f ($b / 1MB) }
    if ($Config.ManutSistema) {
        $b = (Clear-Antigos "$env:windir\Temp" 2) + (Clear-Antigos "$env:windir\SoftwareDistribution\Download" 7)
        $total += $b; $partes += 'Windows: {0:N0} MB' -f ($b / 1MB)
    }
    $resumo = '{0:N0} MB liberados ({1})' -f ($total / 1MB), ($partes -join ', ')
    try { Add-Content -Path $ManutLogPath -Value ('{0:dd/MM/yyyy HH:mm}  {1}' -f (Get-Date), $resumo) -Encoding UTF8 } catch {}
    # Aviso rápido perto do relógio
    try {
        $aviso = New-Object System.Windows.Forms.NotifyIcon
        $icoAviso = Join-Path $PSScriptRoot 'MonitorPC.ico'
        $aviso.Icon = if (Test-Path $icoAviso) { New-Object System.Drawing.Icon($icoAviso, 16, 16) } else { [System.Drawing.SystemIcons]::Information }
        $aviso.Visible = $true
        $aviso.ShowBalloonTip(8000, 'Monitor PC · manutenção semanal', $resumo, 'Info')
        Start-Sleep -Seconds 9
        $aviso.Visible = $false; $aviso.Dispose()
    } catch {}
    exit
}
#endregion

#region Tema
# Cores prontas do seletor de aparência
$Temas = [ordered]@{
    'Verde'    = '#A6F23A'
    'Vermelho' = '#FF4766'
    'Azul'     = '#3D9BFF'
    'Roxo'     = '#A970FF'
    'Branco'   = '#E6E6EB'
    'Laranja'  = '#FF9A3D'
    'Ciano'    = '#2EE6D6'
    'Rosa'     = '#FF5FC8'
}

# A interface foi desenhada em verde; cada tom verde do XAML vira um tom derivado da cor escolhida
function Get-Paleta([string]$hex) {
    $c = [Windows.Media.ColorConverter]::ConvertFromString($hex)
    $fundo = @(11, 11, 13)
    function Mix($de, $para, [double]$t) {
        '{0:X2}{1:X2}{2:X2}' -f [int]($de[0] + ($para[0] - $de[0]) * $t), [int]($de[1] + ($para[1] - $de[1]) * $t), [int]($de[2] + ($para[2] - $de[2]) * $t)
    }
    $p = @($c.R, $c.G, $c.B)
    $preto = @(0, 0, 0); $branco = @(255, 255, 255)
    [ordered]@{
        'A6F23A' = Mix $p $p 0          # cor principal
        'B8FF4A' = Mix $p $branco 0.18  # topo das barras
        '7BC62D' = Mix $p $preto 0.22
        '5DB82A' = Mix $p $preto 0.32
        '4E9A1E' = Mix $p $preto 0.45
        '3E6B12' = Mix $p $preto 0.60   # base das barras
        '2F5A14' = Mix $p $preto 0.66   # parte livre da rosca
        '1C2418' = Mix $fundo $p 0.10   # fundo de ícones e trilhas
        '2C3A22' = Mix $fundo $p 0.20
        '1A2614' = Mix $fundo $p 0.12   # item ativo do menu
        '2F4A16' = Mix $fundo $p 0.30
        '1F2E12' = Mix $fundo $p 0.14   # card de dica
        '1E2A14' = Mix $fundo $p 0.15
    }
}

if ("$($Config.CorDestaque)" -notmatch '^#[0-9A-Fa-f]{6}$') { $Config.CorDestaque = '#A6F23A' }
$Paleta = Get-Paleta $Config.CorDestaque
$Cor.Lima       = '#' + $Paleta['A6F23A']
$Cor.LimaMedia  = '#' + $Paleta['7BC62D']
$Cor.LimaEscura = '#' + $Paleta['4E9A1E']
$Cor.Trilha     = '#' + $Paleta['1C2418']
#endregion

#region Ícone do aplicativo
# Desenha o ícone (duas setas, igual ao logo do menu) na cor do tema
function New-ImagemIcone([int]$s, [string]$hex) {
    $cor = [Windows.Media.ColorConverter]::ConvertFromString($hex)
    $dv = New-Object Windows.Media.DrawingVisual
    $dc = $dv.RenderOpen()
    $raio = [Math]::Max(3, $s * 0.22)
    $fundo = New-Object Windows.Media.LinearGradientBrush([Windows.Media.Color]::FromRgb(30, 30, 35), [Windows.Media.Color]::FromRgb(12, 12, 14), 90)
    $borda = New-Object Windows.Media.Pen((New-Object Windows.Media.SolidColorBrush([Windows.Media.Color]::FromArgb(110, $cor.R, $cor.G, $cor.B))), [Math]::Max(1, $s / 56))
    $m = $borda.Thickness / 2
    $dc.DrawRoundedRectangle($fundo, $borda, (New-Object Windows.Rect($m, $m, ($s - 2 * $m), ($s - 2 * $m))), $raio, $raio)
    $esp = [Math]::Max(1.6, $s * 0.105)
    $l = $s * 0.27; $r = $s * 0.73; $cx = $s / 2; $h = $s * 0.20
    $escura = [Windows.Media.Color]::FromRgb([int]($cor.R * 0.62), [int]($cor.G * 0.62), [int]($cor.B * 0.62))
    foreach ($par in @(@(0.50, $cor), @(0.72, $escura))) {
        $caneta = New-Object Windows.Media.Pen((New-Object Windows.Media.SolidColorBrush($par[1])), $esp)
        $caneta.StartLineCap = 'Round'; $caneta.EndLineCap = 'Round'; $caneta.LineJoin = 'Round'
        $yb = $s * $par[0]
        $fig = New-Object Windows.Media.PathFigure
        $fig.StartPoint = New-Object Windows.Point($l, $yb)
        $fig.Segments.Add((New-Object Windows.Media.LineSegment((New-Object Windows.Point($cx, ($yb - $h))), $true)))
        $fig.Segments.Add((New-Object Windows.Media.LineSegment((New-Object Windows.Point($r, $yb)), $true)))
        $geo = New-Object Windows.Media.PathGeometry
        $geo.Figures.Add($fig)
        $dc.DrawGeometry($null, $caneta, $geo)
    }
    $dc.Close()
    $rtb = New-Object Windows.Media.Imaging.RenderTargetBitmap($s, $s, 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($dv)
    $rtb
}

# Grava um .ico com vários tamanhos. Tamanhos pequenos em BMP (compatível com bandeja/atalhos), 256 em PNG.
function Save-IconeIco([string]$hex, [string]$caminho) {
    $tamanhos = 16, 20, 24, 32, 40, 48, 64, 256
    $dados = foreach ($s in $tamanhos) {
        $img = New-ImagemIcone $s $hex
        if ($s -eq 256) {
            $enc = New-Object Windows.Media.Imaging.PngBitmapEncoder
            $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($img))
            $ms = New-Object IO.MemoryStream
            $enc.Save($ms)
            , $ms.ToArray()
        } else {
            $bgra = New-Object Windows.Media.Imaging.FormatConvertedBitmap($img, [Windows.Media.PixelFormats]::Bgra32, $null, 0)
            $px = New-Object byte[] ($s * $s * 4)
            $bgra.CopyPixels($px, $s * 4, 0)
            $mascara = [int]([Math]::Ceiling($s / 32.0) * 4) * $s
            $ms = New-Object IO.MemoryStream
            $bw = New-Object IO.BinaryWriter($ms)
            $bw.Write([int]40); $bw.Write([int]$s); $bw.Write([int]($s * 2)); $bw.Write([int16]1); $bw.Write([int16]32)
            $bw.Write([int]0); $bw.Write([int]($px.Length + $mascara)); $bw.Write([int]0); $bw.Write([int]0); $bw.Write([int]0); $bw.Write([int]0)
            for ($y = $s - 1; $y -ge 0; $y--) { $bw.Write($px, $y * $s * 4, $s * 4) }   # BMP guarda de baixo para cima
            $bw.Write((New-Object byte[] $mascara))
            $bw.Flush()
            , $ms.ToArray()
        }
    }
    $fs = [IO.File]::Create($caminho)
    $bw = New-Object IO.BinaryWriter($fs)
    $bw.Write([int16]0); $bw.Write([int16]1); $bw.Write([int16]$tamanhos.Count)
    $offset = 6 + 16 * $tamanhos.Count
    for ($i = 0; $i -lt $tamanhos.Count; $i++) {
        $s = $tamanhos[$i]; $lado = if ($s -ge 256) { 0 } else { $s }
        $bw.Write([byte]$lado); $bw.Write([byte]$lado); $bw.Write([byte]0); $bw.Write([byte]0)
        $bw.Write([int16]1); $bw.Write([int16]32); $bw.Write([int]$dados[$i].Length); $bw.Write([int]$offset)
        $offset += $dados[$i].Length
    }
    foreach ($d in $dados) { $bw.Write($d) }
    $bw.Close()
}

# O ícone fica ao lado do programa (atalhos leem dali); se a pasta não permitir gravar, usa a de configurações.
# Só é redesenhado quando a cor muda.
$IconePath = Join-Path $PSScriptRoot 'MonitorPC.ico'
if (-not (Test-Path $IconePath) -or $Config.IconeCor -ne $Config.CorDestaque) {
    try { Save-IconeIco $Cor.Lima $IconePath } catch {
        $IconePath = Join-Path $ConfigDir 'MonitorPC.ico'
        try { Save-IconeIco $Cor.Lima $IconePath } catch {}
    }
    $Config.IconeCor = $Config.CorDestaque
    Save-Config
}

# Lançador "Monitor PC.exe": é o que aparece em Configurações > Inicialização e nos atalhos,
# com nome e ícone próprios (em vez de "Windows PowerShell"). Abre o Monitor sem piscar janela preta.
$LancadorPath = Join-Path $PSScriptRoot 'Monitor PC.exe'
$CodigoLancador = @'
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
[assembly: AssemblyTitle("Monitor PC")]
[assembly: AssemblyDescription("Monitor PC - desempenho, alertas e segurança")]
[assembly: AssemblyCompany("Monitor PC")]
[assembly: AssemblyProduct("Monitor PC")]
[assembly: AssemblyCopyright("Monitor PC")]
[assembly: AssemblyVersion("2.0.0.0")]
[assembly: AssemblyFileVersion("2.0.0.0")]
public static class Lancador {
    [STAThread]
    public static void Main(string[] args) {
        string pasta = AppDomain.CurrentDomain.BaseDirectory;
        string script = Path.Combine(pasta, "MonitorPC.ps1");
        string ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"WindowsPowerShell\v1.0\powershell.exe");
        ProcessStartInfo psi = new ProcessStartInfo(ps, "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" + script + "\" " + string.Join(" ", args));
        psi.UseShellExecute = false;
        psi.CreateNoWindow = true;
        psi.WorkingDirectory = pasta;
        Process.Start(psi);
    }
}
'@

function Update-Lancador {
    if ((Test-Path $LancadorPath) -and (Test-Path $IconePath) -and
        (Get-Item $LancadorPath).LastWriteTime -ge (Get-Item $IconePath).LastWriteTime) { return $true }
    try {
        $prov = New-Object Microsoft.CSharp.CSharpCodeProvider
        $par = New-Object System.CodeDom.Compiler.CompilerParameters
        $par.GenerateExecutable = $true
        $par.OutputAssembly = $LancadorPath
        $par.CompilerOptions = "/target:winexe /optimize" + $(if (Test-Path $IconePath) { " /win32icon:`"$IconePath`"" } else { '' })
        [void]$par.ReferencedAssemblies.Add('System.dll')
        $res = $prov.CompileAssemblyFromSource($par, $CodigoLancador)
        -not $res.Errors.HasErrors
    } catch { $false }
}
$TemLancador = Update-Lancador
# Atalhos antigos (que abriam o PowerShell direto) passam a usar o lançador — fica feito em segundo plano depois de abrir
$script:AtualizarAtalhosAntigos = $TemLancador

# Identidade própria na barra de tarefas (senão o Windows agrupa com o PowerShell)
try {
    Add-Type -Namespace MonitorPC -Name Shell -MemberDefinition '[DllImport("shell32.dll")] public static extern int SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string id);'
    [void][MonitorPC.Shell]::SetCurrentProcessExplicitAppUserModelID('MonitorPC.App')
} catch {}
#endregion

#region Interface (XAML)
$XamlTexto = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Monitor PC" Width="1340" Height="860" MinWidth="1120" MinHeight="680"
        WindowStyle="None" ResizeMode="CanResize" Background="#0B0B0D"
        Foreground="#F2F2F5" FontFamily="Segoe UI" FontSize="13"
        TextOptions.TextFormattingMode="Display" UseLayoutRounding="True" SnapsToDevicePixels="True"
        WindowStartupLocation="CenterScreen">
  <WindowChrome.WindowChrome>
    <WindowChrome CaptionHeight="64" ResizeBorderThickness="6" GlassFrameThickness="0" CornerRadius="0" UseAeroCaptionButtons="False"/>
  </WindowChrome.WindowChrome>

  <Window.Resources>
    <FontFamily x:Key="Icones">Segoe Fluent Icons, Segoe MDL2 Assets</FontFamily>

    <SolidColorBrush x:Key="FundoCard" Color="#121214"/>
    <LinearGradientBrush x:Key="BarraLima" StartPoint="0,0" EndPoint="0,1">
      <GradientStop Color="#B8FF4A" Offset="0"/>
      <GradientStop Color="#3E6B12" Offset="1"/>
    </LinearGradientBrush>

    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource FundoCard}"/>
      <Setter Property="BorderBrush" Value="#1F1F23"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="18"/>
      <Setter Property="Padding" Value="20,18"/>
      <Setter Property="Margin" Value="7"/>
    </Style>
    <Style x:Key="TituloCard" TargetType="TextBlock">
      <Setter Property="FontSize" Value="16"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style x:Key="Suave" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#8E8E96"/>
      <Setter Property="FontSize" Value="12"/>
    </Style>
    <Style x:Key="Icone" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="{StaticResource Icones}"/>
      <Setter Property="Foreground" Value="#A6F23A"/>
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="HorizontalAlignment" Value="Center"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <!-- Etiqueta arredondada no canto dos cards -->
    <Style x:Key="Chip" TargetType="Border">
      <Setter Property="Background" Value="#18181B"/>
      <Setter Property="BorderBrush" Value="#2A2A2F"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="10"/>
      <Setter Property="Padding" Value="10,4"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <!-- Círculo de ícone -->
    <Style x:Key="Bolha" TargetType="Border">
      <Setter Property="Width" Value="34"/>
      <Setter Property="Height" Value="34"/>
      <Setter Property="CornerRadius" Value="17"/>
      <Setter Property="Background" Value="#1C2418"/>
      <Setter Property="BorderBrush" Value="#2C3A22"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>

    <Style TargetType="TabControl">
      <Setter Property="Foreground" Value="#F2F2F5"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="0"/>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Visibility" Value="Collapsed"/>
      <Setter Property="Foreground" Value="#F2F2F5"/>
    </Style>

    <!-- Itens do menu lateral -->
    <Style x:Key="Menu" TargetType="RadioButton">
      <Setter Property="Foreground" Value="#9A9AA2"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Margin" Value="0,2"/>
      <Setter Property="GroupName" Value="menu"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="Bd" CornerRadius="12" Padding="14,11" Background="Transparent" BorderThickness="1" BorderBrush="Transparent">
              <StackPanel Orientation="Horizontal">
                <TextBlock x:Name="Ic" Text="{TemplateBinding Tag}" FontFamily="{StaticResource Icones}" FontSize="16" Foreground="{TemplateBinding Foreground}" VerticalAlignment="Center"/>
                <ContentPresenter Margin="14,0,0,0" VerticalAlignment="Center"/>
              </StackPanel>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#161618"/>
                <Setter Property="Foreground" Value="#F2F2F5"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#1A2614"/>
                <Setter TargetName="Bd" Property="BorderBrush" Value="#2F4A16"/>
                <Setter Property="Foreground" Value="#A6F23A"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="Button">
      <Setter Property="Background" Value="#A6F23A"/>
      <Setter Property="Foreground" Value="#0B0B0D"/>
      <Setter Property="BorderBrush" Value="Transparent"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="18,8"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1"
                    CornerRadius="11" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Opacity" Value="0.85"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Bd" Property="Opacity" Value="0.65"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Bd" Property="Opacity" Value="0.35"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BtnSecundario" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="#18181B"/>
      <Setter Property="Foreground" Value="#F2F2F5"/>
      <Setter Property="BorderBrush" Value="#2A2A2F"/>
    </Style>
    <Style x:Key="BtnIcone" TargetType="Button">
      <Setter Property="Width" Value="40"/>
      <Setter Property="Height" Value="40"/>
      <Setter Property="Foreground" Value="#C8C8CE"/>
      <Setter Property="FontFamily" Value="{StaticResource Icones}"/>
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="WindowChrome.IsHitTestVisibleInChrome" Value="True"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="#161618" BorderBrush="#1F1F23" BorderThickness="1" CornerRadius="20">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#1C2418"/>
                <Setter Property="Foreground" Value="#A6F23A"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BtnJanela" TargetType="Button">
      <Setter Property="Width" Value="44"/>
      <Setter Property="Height" Value="32"/>
      <Setter Property="Foreground" Value="#8E8E96"/>
      <Setter Property="FontFamily" Value="{StaticResource Icones}"/>
      <Setter Property="FontSize" Value="10"/>
      <Setter Property="WindowChrome.IsHitTestVisibleInChrome" Value="True"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="Transparent" CornerRadius="8">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#1C1C20"/>
                <Setter Property="Foreground" Value="#FFFFFF"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Link" TargetType="Button">
      <Setter Property="Foreground" Value="#8E8E96"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <TextBlock x:Name="T" Text="{TemplateBinding Content}" Foreground="{TemplateBinding Foreground}" FontSize="12"/>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="T" Property="Foreground" Value="#A6F23A"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Etiqueta que liga/desliga (programas do Modo Jogo) -->
    <Style x:Key="ChipSel" TargetType="ToggleButton">
      <Setter Property="Foreground" Value="#C8C8CE"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Margin" Value="0,0,8,8"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ToggleButton">
            <Border x:Name="Bd" CornerRadius="10" Padding="10,6" Background="#141416" BorderBrush="#2A2A2F" BorderThickness="1">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#1A2614"/>
                <Setter TargetName="Bd" Property="BorderBrush" Value="#A6F23A"/>
                <Setter Property="Foreground" Value="#A6F23A"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Seletor segmentado -->
    <Style x:Key="Seg" TargetType="RadioButton">
      <Setter Property="Foreground" Value="#8E8E96"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="Bd" Padding="14,6" CornerRadius="9" Background="Transparent">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#A6F23A"/>
                <Setter Property="Foreground" Value="#0B0B0D"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="ProgressBar">
      <Setter Property="Height" Value="8"/>
      <Setter Property="Maximum" Value="100"/>
      <Setter Property="Background" Value="#1C2418"/>
      <Setter Property="Foreground" Value="#A6F23A"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ProgressBar">
            <Grid>
              <Border x:Name="PART_Track" Background="{TemplateBinding Background}" CornerRadius="4"/>
              <Border x:Name="PART_Indicator" Background="{TemplateBinding Foreground}" CornerRadius="4" HorizontalAlignment="Left"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#F2F2F5"/>
      <Setter Property="Margin" Value="0,6"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <StackPanel Orientation="Horizontal" Background="Transparent">
              <Border x:Name="Fundo" Width="38" Height="22" CornerRadius="11" Background="#1C2418" BorderBrush="#2C3A22" BorderThickness="1">
                <Ellipse x:Name="Bola" Width="14" Height="14" Fill="#5E5E66" HorizontalAlignment="Left" Margin="4,0"/>
              </Border>
              <ContentPresenter Margin="12,0,0,0" VerticalAlignment="Center"/>
            </StackPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="Fundo" Property="Background" Value="#A6F23A"/>
                <Setter TargetName="Bola" Property="Fill" Value="#0B0B0D"/>
                <Setter TargetName="Bola" Property="HorizontalAlignment" Value="Right"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="TextBox">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="#F2F2F5"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="CaretBrush" Value="#A6F23A"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>

    <!-- Slider -->
    <Style x:Key="TrilhaCheia" TargetType="RepeatButton">
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RepeatButton"><Border Height="6" CornerRadius="3" Background="#A6F23A"/></ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="TrilhaVazia" TargetType="RepeatButton">
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RepeatButton"><Border Height="6" CornerRadius="3" Background="#1C2418"/></ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="Slider">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Slider">
            <Track x:Name="PART_Track" Height="24">
              <Track.DecreaseRepeatButton><RepeatButton Command="Slider.DecreaseLarge" Style="{StaticResource TrilhaCheia}"/></Track.DecreaseRepeatButton>
              <Track.IncreaseRepeatButton><RepeatButton Command="Slider.IncreaseLarge" Style="{StaticResource TrilhaVazia}"/></Track.IncreaseRepeatButton>
              <Track.Thumb>
                <Thumb Cursor="Hand">
                  <Thumb.Template>
                    <ControlTemplate TargetType="Thumb">
                      <Ellipse Width="18" Height="18" Fill="#F2F2F5" Stroke="#A6F23A" StrokeThickness="3"/>
                    </ControlTemplate>
                  </Thumb.Template>
                </Thumb>
              </Track.Thumb>
            </Track>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="ScrollBar">
      <Setter Property="Width" Value="6"/>
      <Setter Property="MinWidth" Value="6"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Track x:Name="PART_Track" IsDirectionReversed="True">
              <Track.Thumb>
                <Thumb><Thumb.Template><ControlTemplate TargetType="Thumb"><Border Background="#2A2A2F" CornerRadius="3"/></ControlTemplate></Thumb.Template></Thumb>
              </Track.Thumb>
            </Track>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="Orientation" Value="Horizontal">
          <Setter Property="Width" Value="Auto"/>
          <Setter Property="MinWidth" Value="0"/>
          <Setter Property="Height" Value="6"/>
          <Setter Property="MinHeight" Value="6"/>
          <Setter Property="Template">
            <Setter.Value>
              <ControlTemplate TargetType="ScrollBar">
                <Track x:Name="PART_Track">
                  <Track.Thumb>
                    <Thumb><Thumb.Template><ControlTemplate TargetType="Thumb"><Border Background="#2A2A2F" CornerRadius="3"/></ControlTemplate></Thumb.Template></Thumb>
                  </Track.Thumb>
                </Track>
              </ControlTemplate>
            </Setter.Value>
          </Setter>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style TargetType="DataGrid">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="#F2F2F5"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="RowBackground" Value="Transparent"/>
      <Setter Property="AlternatingRowBackground" Value="Transparent"/>
      <Setter Property="GridLinesVisibility" Value="None"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="AutoGenerateColumns" Value="False"/>
      <Setter Property="IsReadOnly" Value="True"/>
      <Setter Property="SelectionMode" Value="Single"/>
      <Setter Property="CanUserAddRows" Value="False"/>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="#8E8E96"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Padding" Value="10,8"/>
      <Setter Property="BorderBrush" Value="#1F1F23"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
    </Style>
    <Style TargetType="DataGridRow">
      <Setter Property="Margin" Value="0,2"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#161618"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="DataGridCell">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridCell">
            <Border Background="{TemplateBinding Background}" Padding="10,0">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#1A2614"/>
          <Setter Property="Foreground" Value="#A6F23A"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style TargetType="ListBoxItem">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListBoxItem">
            <Border Background="#141416" CornerRadius="10" Padding="12,9" Margin="0,3">
              <ContentPresenter/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Storyboard x:Key="Pulsar" RepeatBehavior="Forever" AutoReverse="True">
      <DoubleAnimation Storyboard.TargetProperty="Opacity" From="1" To="0.25" Duration="0:0:1.1"/>
    </Storyboard>
  </Window.Resources>

  <Border x:Name="raiz">
    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="236"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <!-- ================= MENU LATERAL ================= -->
      <Border Background="#0B0B0D" BorderBrush="#18181B" BorderThickness="0,0,1,0">
        <DockPanel Margin="16,18,16,16">
          <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="6,2,0,28">
            <Path Data="M0,14 L10,4 L20,14 M0,22 L10,12 L20,22" Stroke="#A6F23A" StrokeThickness="3" StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round" VerticalAlignment="Center"/>
            <TextBlock Text="MONITOR PC" FontSize="17" FontWeight="Bold" Margin="12,0,0,0" VerticalAlignment="Center"/>
          </StackPanel>

          <StackPanel DockPanel.Dock="Bottom">
            <Border Background="#161618" BorderBrush="#1F1F23" BorderThickness="1" CornerRadius="16" Padding="16,14" Margin="0,0,0,12">
              <StackPanel>
                <TextBlock Text="Saúde do sistema" FontWeight="SemiBold"/>
                <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                  <Border Style="{StaticResource Bolha}"><TextBlock Text="&#xE95E;" Style="{StaticResource Icone}"/></Border>
                  <TextBlock x:Name="txtSaudeNum" Text="--" FontSize="30" FontWeight="SemiBold" Margin="10,0,6,0" VerticalAlignment="Center"/>
                  <TextBlock x:Name="txtSaudeRot" Style="{StaticResource Suave}" VerticalAlignment="Center" Margin="0,8,0,0"/>
                </StackPanel>
                <UniformGrid x:Name="ugPontos" Columns="10" Margin="0,12,0,0"/>
                <TextBlock x:Name="txtUptime" Style="{StaticResource Suave}" Margin="0,10,0,0" FontSize="11"/>
              </StackPanel>
            </Border>
            <TextBlock x:Name="txtStatus" Style="{StaticResource Suave}" FontSize="10" TextWrapping="Wrap" Margin="6,0,6,8"/>
            <Button x:Name="btnSair" Style="{StaticResource Link}" Content="⏻   Sair do Monitor" Margin="8,4,0,0" HorizontalAlignment="Left"/>
          </StackPanel>

          <StackPanel>
            <RadioButton x:Name="mnPainel" Style="{StaticResource Menu}" Tag="&#xECA5;" Content="Painel" IsChecked="True"/>
            <RadioButton x:Name="mnJogos" Style="{StaticResource Menu}" Tag="&#xE7FC;" Content="Jogos e FPS"/>
            <RadioButton x:Name="mnProc" Style="{StaticResource Menu}" Tag="&#xE9F5;" Content="Processos"/>
            <RadioButton x:Name="mnApps" Style="{StaticResource Menu}" Tag="&#xE71D;" Content="Aplicativos"/>
            <RadioButton x:Name="mnAlertas" Style="{StaticResource Menu}" Tag="&#xEA8F;" Content="Alertas"/>
            <RadioButton x:Name="mnVida" Style="{StaticResource Menu}" Tag="&#xE95E;" Content="Vida útil"/>
            <RadioButton x:Name="mnSeg" Style="{StaticResource Menu}" Tag="&#xEA18;" Content="Segurança"/>
            <RadioButton x:Name="mnOtim" Style="{StaticResource Menu}" Tag="&#xEC4A;" Content="Otimizar"/>
          </StackPanel>
        </DockPanel>
      </Border>

      <!-- ================= ÁREA PRINCIPAL ================= -->
      <Grid Grid.Column="1">
        <Grid.RowDefinitions>
          <RowDefinition Height="64"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- Barra superior -->
        <DockPanel Margin="22,0,8,0">
          <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
            <Border x:Name="bdAdmin" Style="{StaticResource Chip}" Margin="0,0,10,0">
              <TextBlock x:Name="txtAdmin" FontSize="11"/>
            </Border>
            <Button x:Name="btnAdmin" Style="{StaticResource BtnSecundario}" Content="Abrir como admin" FontSize="12" Padding="12,6" Margin="0,0,10,0"
                    WindowChrome.IsHitTestVisibleInChrome="True"/>
            <TextBlock x:Name="txtRelogio" FontSize="15" FontWeight="SemiBold" VerticalAlignment="Center" Margin="4,0,14,0"/>
            <Button x:Name="btnNovaVersao" Content="Nova versão" FontSize="12" Padding="12,6" Margin="0,0,10,0" Visibility="Collapsed"
                    WindowChrome.IsHitTestVisibleInChrome="True"/>
            <Button x:Name="btnTema" Style="{StaticResource BtnIcone}" Content="&#xE790;" Margin="0,0,10,0" ToolTip="Aparência e cores"/>
            <Grid Margin="0,0,10,0">
              <Button x:Name="btnSino" Style="{StaticResource BtnIcone}" Content="&#xEA8F;"/>
              <Border x:Name="bdBadge" Background="#A6F23A" CornerRadius="8" MinWidth="16" Height="16" HorizontalAlignment="Right" VerticalAlignment="Top"
                      Margin="0,-2,-2,0" Visibility="Collapsed" IsHitTestVisible="False">
                <TextBlock x:Name="txtBadge" FontSize="10" FontWeight="Bold" Foreground="#0B0B0D" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="4,0"/>
              </Border>
            </Grid>
            <Border Width="40" Height="40" CornerRadius="20" Background="#1A2614" BorderBrush="#2F4A16" BorderThickness="1" Margin="0,0,14,0">
              <TextBlock x:Name="txtAvatar" FontWeight="Bold" Foreground="#A6F23A" FontSize="15" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <Button x:Name="btnMin" Style="{StaticResource BtnJanela}" Content="&#xE921;"/>
            <Button x:Name="btnMax" Style="{StaticResource BtnJanela}" Content="&#xE922;"/>
            <Button x:Name="btnFechar" Style="{StaticResource BtnJanela}" Content="&#xE8BB;"/>
          </StackPanel>
          <Border Width="320" Height="40" HorizontalAlignment="Left" CornerRadius="20" Background="#141416" BorderBrush="#1F1F23" BorderThickness="1"
                  WindowChrome.IsHitTestVisibleInChrome="True">
            <Grid Margin="16,0">
              <TextBlock Text="&#xE721;" Style="{StaticResource Icone}" Foreground="#8E8E96" HorizontalAlignment="Left" FontSize="13"/>
              <TextBox x:Name="txtBusca" Margin="26,0,0,0" WindowChrome.IsHitTestVisibleInChrome="True"/>
              <TextBlock x:Name="phBusca" Text="Buscar processo..." Foreground="#5E5E66" Margin="28,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
            </Grid>
          </Border>
        </DockPanel>

        <TabControl x:Name="tabs" Grid.Row="1" TextElement.Foreground="#F2F2F5">

          <!-- ===================== PAINEL ===================== -->
          <TabItem>
            <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="14,0,14,14">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="1.55*"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="330"/>
                  <RowDefinition Height="290"/>
                  <RowDefinition Height="220"/>
                  <RowDefinition Height="260"/>
                </Grid.RowDefinitions>

                <!-- Visão geral -->
                <Border Style="{StaticResource Card}" Padding="26,22">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition/>
                      <ColumnDefinition Width="170"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                      <RowDefinition/>
                      <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <!-- Chip decorativo -->
                    <Viewbox Grid.Column="1" Margin="0,-8,-8,4" VerticalAlignment="Center">
                      <Canvas Width="200" Height="200">
                        <Ellipse x:Name="brilhoChip" Width="200" Height="200">
                          <Ellipse.Fill>
                            <RadialGradientBrush>
                              <GradientStop Color="#55A6F23A" Offset="0"/>
                              <GradientStop Color="#00A6F23A" Offset="1"/>
                            </RadialGradientBrush>
                          </Ellipse.Fill>
                        </Ellipse>
                        <Path Stroke="#4E9A1E" StrokeThickness="1.5" StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                              Data="M78,34 L78,14 L40,14 M122,34 L122,22 L170,22 M166,92 L188,92 M166,120 L180,120 L180,170 M34,106 L12,106 M34,78 L22,78 L22,40 M92,166 L92,190 M120,166 L120,180 L60,180"/>
                        <Ellipse Canvas.Left="36" Canvas.Top="11" Width="6" Height="6" Fill="#A6F23A"/>
                        <Ellipse Canvas.Left="167" Canvas.Top="19" Width="6" Height="6" Fill="#A6F23A"/>
                        <Ellipse Canvas.Left="185" Canvas.Top="89" Width="6" Height="6" Fill="#A6F23A"/>
                        <Ellipse Canvas.Left="177" Canvas.Top="167" Width="6" Height="6" Fill="#A6F23A"/>
                        <Ellipse Canvas.Left="9" Canvas.Top="103" Width="6" Height="6" Fill="#A6F23A"/>
                        <Ellipse Canvas.Left="19" Canvas.Top="37" Width="6" Height="6" Fill="#A6F23A"/>
                        <Ellipse Canvas.Left="89" Canvas.Top="187" Width="6" Height="6" Fill="#A6F23A"/>
                        <Ellipse Canvas.Left="57" Canvas.Top="177" Width="6" Height="6" Fill="#A6F23A"/>
                        <Path Stroke="#A6F23A" StrokeThickness="3" StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                              Data="M64,34 L64,50 M78,34 L78,50 M92,34 L92,50 M108,34 L108,50 M122,34 L122,50 M136,34 L136,50 M64,150 L64,166 M78,150 L78,166 M92,150 L92,166 M108,150 L108,166 M122,150 L122,166 M136,150 L136,166 M34,64 L50,64 M34,78 L50,78 M34,92 L50,92 M34,106 L50,106 M34,120 L50,120 M34,136 L50,136 M150,64 L166,64 M150,78 L166,78 M150,92 L166,92 M150,106 L166,106 M150,120 L166,120 M150,136 L166,136"/>
                        <Rectangle Canvas.Left="48" Canvas.Top="48" Width="104" Height="104" RadiusX="16" RadiusY="16" Stroke="#A6F23A" StrokeThickness="2.5">
                          <Rectangle.Fill>
                            <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                              <GradientStop Color="#1E2A14" Offset="0"/>
                              <GradientStop Color="#0E0E10" Offset="1"/>
                            </LinearGradientBrush>
                          </Rectangle.Fill>
                          <Rectangle.Effect><DropShadowEffect Color="#A6F23A" BlurRadius="30" ShadowDepth="0" Opacity="0.8"/></Rectangle.Effect>
                        </Rectangle>
                        <Rectangle Canvas.Left="66" Canvas.Top="66" Width="68" Height="68" RadiusX="8" RadiusY="8" Stroke="#4E9A1E" StrokeThickness="1.5" StrokeDashArray="3 2"/>
                        <TextBlock Canvas.Left="66" Canvas.Top="76" Width="68" Text="CPU" TextAlignment="Center" FontSize="12" FontWeight="Bold" Foreground="#7BC62D"/>
                        <TextBlock x:Name="txtChip" Canvas.Left="56" Canvas.Top="92" Width="88" Text="--" TextAlignment="Center" FontSize="24" FontWeight="Bold" Foreground="#A6F23A"/>
                      </Canvas>
                    </Viewbox>

                      <StackPanel VerticalAlignment="Center">
                        <TextBlock x:Name="txtSaudacao" Style="{StaticResource Suave}" FontSize="13"/>
                        <TextBlock Text="Visão Geral do Sistema" FontSize="28" FontWeight="SemiBold" Margin="0,4,0,0" TextWrapping="Wrap"/>
                        <TextBlock x:Name="txtResumo" Style="{StaticResource Suave}" FontSize="13" TextWrapping="Wrap" Margin="0,8,10,0"/>
                      </StackPanel>
                      <Border Grid.Row="1" Grid.ColumnSpan="2" Background="#0E0E10" BorderBrush="#1F1F23" BorderThickness="1" CornerRadius="16" Padding="6" Margin="0,12,0,0">
                        <UniformGrid Columns="3">
                          <StackPanel Orientation="Horizontal" Margin="10,8">
                            <Border Style="{StaticResource Bolha}" Width="40" Height="40" CornerRadius="20" VerticalAlignment="Center"><TextBlock Text="&#xE950;" Style="{StaticResource Icone}"/></Border>
                            <StackPanel Margin="12,0,0,0" VerticalAlignment="Center">
                              <TextBlock Text="CPU" Style="{StaticResource Suave}"/>
                              <TextBlock x:Name="txtKCpu" Text="--" FontSize="22" FontWeight="SemiBold" Margin="0,-2,0,-2"/>
                              <TextBlock x:Name="txtKCpuSt" FontSize="11"/>
                            </StackPanel>
                          </StackPanel>
                          <StackPanel Orientation="Horizontal" Margin="10,8">
                            <Border Style="{StaticResource Bolha}" Width="40" Height="40" CornerRadius="20" VerticalAlignment="Center"><TextBlock Text="&#xE964;" Style="{StaticResource Icone}"/></Border>
                            <StackPanel Margin="12,0,0,0" VerticalAlignment="Center">
                              <TextBlock Text="Memória" Style="{StaticResource Suave}"/>
                              <TextBlock x:Name="txtKRam" Text="--" FontSize="22" FontWeight="SemiBold" Margin="0,-2,0,-2"/>
                              <TextBlock x:Name="txtKRamSt" FontSize="11"/>
                            </StackPanel>
                          </StackPanel>
                          <StackPanel Orientation="Horizontal" Margin="10,8">
                            <Border Style="{StaticResource Bolha}" Width="40" Height="40" CornerRadius="20" VerticalAlignment="Center"><TextBlock Text="&#xEDA2;" Style="{StaticResource Icone}"/></Border>
                            <StackPanel Margin="12,0,0,0" VerticalAlignment="Center">
                              <TextBlock Text="Disco" Style="{StaticResource Suave}"/>
                              <TextBlock x:Name="txtKDsc" Text="--" FontSize="22" FontWeight="SemiBold" Margin="0,-2,0,-2"/>
                              <TextBlock x:Name="txtKDscSt" FontSize="11"/>
                            </StackPanel>
                          </StackPanel>
                        </UniformGrid>
                      </Border>
                  </Grid>
                </Border>

                <!-- Anéis de uso -->
                <Border Grid.Column="1" Style="{StaticResource Card}">
                  <DockPanel>
                    <DockPanel DockPanel.Dock="Top">
                      <Border DockPanel.Dock="Right" Style="{StaticResource Chip}">
                        <StackPanel Orientation="Horizontal">
                          <Ellipse x:Name="pontoVivo" Width="7" Height="7" Fill="#A6F23A" VerticalAlignment="Center"/>
                          <TextBlock Text="Ao vivo" FontSize="11" Margin="6,0,0,0"/>
                        </StackPanel>
                      </Border>
                      <TextBlock Text="Uso em Tempo Real" Style="{StaticResource TituloCard}"/>
                    </DockPanel>
                    <UniformGrid DockPanel.Dock="Bottom" Columns="3" Margin="0,8,0,0">
                      <StackPanel HorizontalAlignment="Center">
                        <Border Style="{StaticResource Bolha}" Width="28" Height="28"><TextBlock Text="&#xE950;" Style="{StaticResource Icone}" FontSize="12"/></Border>
                        <TextBlock Text="CPU" Style="{StaticResource Suave}" HorizontalAlignment="Center" Margin="0,4,0,0"/>
                        <TextBlock x:Name="txtRCpu" FontSize="17" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                      </StackPanel>
                      <StackPanel HorizontalAlignment="Center">
                        <Border Style="{StaticResource Bolha}" Width="28" Height="28"><TextBlock Text="&#xE964;" Style="{StaticResource Icone}" FontSize="12" Foreground="#7BC62D"/></Border>
                        <TextBlock Text="RAM" Style="{StaticResource Suave}" HorizontalAlignment="Center" Margin="0,4,0,0"/>
                        <TextBlock x:Name="txtRRam" FontSize="17" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                      </StackPanel>
                      <StackPanel HorizontalAlignment="Center">
                        <Border Style="{StaticResource Bolha}" Width="28" Height="28"><TextBlock Text="&#xEDA2;" Style="{StaticResource Icone}" FontSize="12" Foreground="#5DB82A"/></Border>
                        <TextBlock Text="Disco" Style="{StaticResource Suave}" HorizontalAlignment="Center" Margin="0,4,0,0"/>
                        <TextBlock x:Name="txtRDsc" FontSize="17" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                      </StackPanel>
                    </UniformGrid>
                    <Viewbox Margin="0,10,0,0">
                      <Grid Width="180" Height="180">
                        <Ellipse Width="156" Height="156" Stroke="#1C2418" StrokeThickness="13"/>
                        <Ellipse Width="120" Height="120" Stroke="#1C2418" StrokeThickness="13"/>
                        <Ellipse Width="84" Height="84" Stroke="#1C2418" StrokeThickness="13"/>
                        <Path x:Name="anCpu" Stroke="#A6F23A" StrokeThickness="13" StrokeStartLineCap="Round" StrokeEndLineCap="Round">
                          <Path.Effect><DropShadowEffect Color="#A6F23A" BlurRadius="14" ShadowDepth="0" Opacity="0.6"/></Path.Effect>
                        </Path>
                        <Path x:Name="anRam" Stroke="#7BC62D" StrokeThickness="13" StrokeStartLineCap="Round" StrokeEndLineCap="Round">
                          <Path.Effect><DropShadowEffect Color="#7BC62D" BlurRadius="14" ShadowDepth="0" Opacity="0.5"/></Path.Effect>
                        </Path>
                        <Path x:Name="anDsc" Stroke="#5DB82A" StrokeThickness="13" StrokeStartLineCap="Round" StrokeEndLineCap="Round">
                          <Path.Effect><DropShadowEffect Color="#5DB82A" BlurRadius="14" ShadowDepth="0" Opacity="0.5"/></Path.Effect>
                        </Path>
                        <TextBlock Text="&#xE945;" Style="{StaticResource Icone}" FontSize="20"/>
                      </Grid>
                    </Viewbox>
                  </DockPanel>
                </Border>

                <!-- Armazenamento -->
                <Border Grid.Column="2" Style="{StaticResource Card}">
                  <DockPanel>
                    <DockPanel DockPanel.Dock="Top">
                      <Border DockPanel.Dock="Right" Style="{StaticResource Chip}">
                        <StackPanel Orientation="Horizontal">
                          <TextBlock Text="&#xEDA2;" Style="{StaticResource Icone}" FontSize="11"/>
                          <TextBlock x:Name="txtEspDrive" FontSize="11" Margin="6,0,0,0"/>
                        </StackPanel>
                      </Border>
                      <TextBlock Text="Armazenamento" Style="{StaticResource TituloCard}"/>
                    </DockPanel>
                    <UniformGrid DockPanel.Dock="Bottom" Columns="3" Margin="0,8,0,0">
                      <StackPanel HorizontalAlignment="Center">
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                          <Ellipse Width="8" Height="8" Fill="#A6F23A" VerticalAlignment="Center"/>
                          <TextBlock x:Name="txtEspUsadoPct" FontSize="17" FontWeight="SemiBold" Margin="6,0,0,0"/>
                        </StackPanel>
                        <TextBlock Text="Usado" Style="{StaticResource Suave}" HorizontalAlignment="Center"/>
                      </StackPanel>
                      <StackPanel HorizontalAlignment="Center">
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                          <Ellipse Width="8" Height="8" Fill="#2F5A14" VerticalAlignment="Center"/>
                          <TextBlock x:Name="txtEspLivrePct" FontSize="17" FontWeight="SemiBold" Margin="6,0,0,0"/>
                        </StackPanel>
                        <TextBlock Text="Livre" Style="{StaticResource Suave}" HorizontalAlignment="Center"/>
                      </StackPanel>
                      <StackPanel HorizontalAlignment="Center">
                        <TextBlock x:Name="txtEspTotal" FontSize="17" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                        <TextBlock Text="Total" Style="{StaticResource Suave}" HorizontalAlignment="Center"/>
                      </StackPanel>
                    </UniformGrid>
                    <Viewbox Margin="0,10,0,0">
                      <Grid Width="180" Height="180">
                        <Ellipse Width="136" Height="136" Stroke="#2F5A14" StrokeThickness="26"/>
                        <Path x:Name="anEsp" Stroke="#A6F23A" StrokeThickness="26">
                          <Path.Effect><DropShadowEffect Color="#A6F23A" BlurRadius="16" ShadowDepth="0" Opacity="0.5"/></Path.Effect>
                        </Path>
                        <Ellipse Width="96" Height="96" Fill="#0E0E10" Stroke="#1F1F23" StrokeThickness="1"/>
                        <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                          <TextBlock x:Name="txtEspLivre" Text="--" FontSize="20" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                          <TextBlock Text="livres" Style="{StaticResource Suave}" FontSize="11" HorizontalAlignment="Center"/>
                        </StackPanel>
                      </Grid>
                    </Viewbox>
                  </DockPanel>
                </Border>

                <!-- Histórico de CPU em barras -->
                <Border Grid.Row="1" Style="{StaticResource Card}">
                  <DockPanel>
                    <DockPanel DockPanel.Dock="Top">
                      <Border DockPanel.Dock="Right" Style="{StaticResource Chip}"><TextBlock Text="Últimos 2 min" FontSize="11"/></Border>
                      <StackPanel Orientation="Horizontal">
                        <Border Style="{StaticResource Bolha}" Width="30" Height="30"><TextBlock Text="&#xE950;" Style="{StaticResource Icone}" FontSize="13"/></Border>
                        <TextBlock Text="Histórico do Processador" Style="{StaticResource TituloCard}" Margin="10,0,0,0"/>
                      </StackPanel>
                    </DockPanel>
                    <StackPanel DockPanel.Dock="Left" Width="120" Margin="0,18,10,0">
                      <StackPanel Orientation="Horizontal">
                        <TextBlock x:Name="txtCpuMedia" Text="--" FontSize="34" FontWeight="SemiBold"/>
                        <TextBlock Text="%" FontSize="16" Foreground="#8E8E96" VerticalAlignment="Bottom" Margin="3,0,0,7"/>
                      </StackPanel>
                      <TextBlock Text="Média" Style="{StaticResource Suave}"/>
                      <StackPanel Orientation="Horizontal" Margin="0,18,0,0">
                        <TextBlock x:Name="txtCpuPico" Text="--" FontSize="24" FontWeight="SemiBold"/>
                        <TextBlock Text="%" FontSize="13" Foreground="#8E8E96" VerticalAlignment="Bottom" Margin="3,0,0,5"/>
                      </StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <Ellipse Width="7" Height="7" Fill="#A6F23A" VerticalAlignment="Center"/>
                        <TextBlock Text="Pico" Style="{StaticResource Suave}" Margin="6,0,0,0"/>
                      </StackPanel>
                      <TextBlock x:Name="txtCpuThreads" Style="{StaticResource Suave}" Margin="0,18,0,0" TextWrapping="Wrap"/>
                    </StackPanel>
                    <Grid Margin="0,12,0,0">
                      <Grid.RowDefinitions>
                        <RowDefinition/>
                        <RowDefinition Height="Auto"/>
                      </Grid.RowDefinitions>
                      <Grid>
                        <Line X1="0" X2="1" Stretch="Fill" VerticalAlignment="Top" Stroke="#1C1C20" StrokeDashArray="4 4"/>
                        <Line X1="0" X2="1" Stretch="Fill" VerticalAlignment="Center" Stroke="#1C1C20" StrokeDashArray="4 4"/>
                        <UniformGrid x:Name="ugBarras" Rows="1"/>
                      </Grid>
                      <UniformGrid x:Name="ugRotulos" Grid.Row="1" Rows="1" Margin="0,8,0,0"/>
                    </Grid>
                  </DockPanel>
                </Border>

                <!-- Rede -->
                <Border Grid.Row="1" Grid.Column="1" Style="{StaticResource Card}">
                  <DockPanel>
                    <DockPanel DockPanel.Dock="Top">
                      <Border DockPanel.Dock="Right" Style="{StaticResource Chip}"><TextBlock Text="Últimos 60 s" FontSize="11"/></Border>
                      <TextBlock Text="Rede" Style="{StaticResource TituloCard}"/>
                    </DockPanel>
                    <Grid DockPanel.Dock="Top" Margin="0,12,0,0">
                      <StackPanel>
                        <TextBlock x:Name="txtRx" Text="--" FontSize="26" FontWeight="SemiBold"/>
                        <StackPanel Orientation="Horizontal">
                          <TextBlock Text="&#xE896;" Style="{StaticResource Icone}" FontSize="11"/>
                          <TextBlock Text="Download" Style="{StaticResource Suave}" Margin="6,0,0,0"/>
                        </StackPanel>
                      </StackPanel>
                      <StackPanel HorizontalAlignment="Right">
                        <TextBlock x:Name="txtTx" Text="--" FontSize="15" FontWeight="SemiBold" HorizontalAlignment="Right" Foreground="#7BC62D"/>
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                          <TextBlock Text="&#xE898;" Style="{StaticResource Icone}" FontSize="11" Foreground="#7BC62D"/>
                          <TextBlock Text="Upload" Style="{StaticResource Suave}" Margin="6,0,0,0"/>
                        </StackPanel>
                      </StackPanel>
                    </Grid>
                    <TextBlock x:Name="txtRedeEscala" DockPanel.Dock="Bottom" Style="{StaticResource Suave}" FontSize="11" Margin="0,6,0,0"/>
                    <Grid Margin="0,12,0,0" ClipToBounds="True">
                      <Line X1="0" X2="1" Stretch="Fill" VerticalAlignment="Top" Stroke="#1C1C20" StrokeDashArray="4 4"/>
                      <Line X1="0" X2="1" Stretch="Fill" VerticalAlignment="Center" Stroke="#1C1C20" StrokeDashArray="4 4"/>
                      <Canvas x:Name="cvRede">
                        <Polygon x:Name="agRx">
                          <Polygon.Fill><LinearGradientBrush StartPoint="0,0" EndPoint="0,1"><GradientStop Color="#44A6F23A" Offset="0"/><GradientStop Color="#00A6F23A" Offset="1"/></LinearGradientBrush></Polygon.Fill>
                        </Polygon>
                        <Polygon x:Name="agTx" Fill="Transparent"/>
                        <Polyline x:Name="lnTx" Stroke="#4E9A1E" StrokeThickness="1.5" StrokeDashArray="3 2"/>
                        <Polyline x:Name="lnRx" Stroke="#A6F23A" StrokeThickness="2.2" StrokeLineJoin="Round">
                          <Polyline.Effect><DropShadowEffect Color="#A6F23A" BlurRadius="10" ShadowDepth="0" Opacity="0.7"/></Polyline.Effect>
                        </Polyline>
                      </Canvas>
                      <Canvas x:Name="cvPontos" IsHitTestVisible="False"/>
                    </Grid>
                  </DockPanel>
                </Border>

                <!-- Top processos -->
                <Border Grid.Row="1" Grid.Column="2" Grid.RowSpan="2" Style="{StaticResource Card}">
                  <DockPanel>
                    <DockPanel DockPanel.Dock="Top" Margin="0,0,0,8">
                      <Button x:Name="btnVerTodos" DockPanel.Dock="Right" Style="{StaticResource Link}" Content="Ver todos" VerticalAlignment="Center"/>
                      <TextBlock Text="Top Processos" Style="{StaticResource TituloCard}"/>
                    </DockPanel>
                    <ItemsControl x:Name="icTop">
                      <ItemsControl.ItemTemplate>
                        <DataTemplate>
                          <Border Background="#141416" CornerRadius="12" Padding="10,8" Margin="0,4">
                            <Grid>
                              <Grid.ColumnDefinitions><ColumnDefinition Width="44"/><ColumnDefinition/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                              <Border Width="36" Height="36" CornerRadius="10" Background="#1C2418" HorizontalAlignment="Left">
                                <Grid>
                                  <TextBlock Text="&#xE9F5;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" Foreground="#A6F23A" HorizontalAlignment="Center" VerticalAlignment="Center">
                                    <TextBlock.Style>
                                      <Style TargetType="TextBlock">
                                        <Setter Property="Visibility" Value="Collapsed"/>
                                        <Style.Triggers>
                                          <DataTrigger Binding="{Binding Icone}" Value="{x:Null}"><Setter Property="Visibility" Value="Visible"/></DataTrigger>
                                        </Style.Triggers>
                                      </Style>
                                    </TextBlock.Style>
                                  </TextBlock>
                                  <Image Source="{Binding Icone}" Width="22" Height="22"/>
                                </Grid>
                              </Border>
                              <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                <TextBlock Text="{Binding Nome}" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
                                <TextBlock Text="{Binding Sub}" Foreground="#8E8E96" FontSize="11" TextTrimming="CharacterEllipsis"/>
                              </StackPanel>
                              <StackPanel Grid.Column="2" VerticalAlignment="Center" Margin="8,0,2,0">
                                <TextBlock Text="{Binding Texto}" FontWeight="SemiBold" HorizontalAlignment="Right" Foreground="#A6F23A"/>
                                <TextBlock Text="RAM" Foreground="#8E8E96" FontSize="11" HorizontalAlignment="Right"/>
                              </StackPanel>
                            </Grid>
                          </Border>
                        </DataTemplate>
                      </ItemsControl.ItemTemplate>
                    </ItemsControl>
                  </DockPanel>
                </Border>

                <!-- Memória -->
                <Border Grid.Row="2" Style="{StaticResource Card}">
                  <DockPanel>
                    <DockPanel DockPanel.Dock="Top">
                      <Border DockPanel.Dock="Right" Style="{StaticResource Chip}"><TextBlock Text="Últimos 2 min" FontSize="11"/></Border>
                      <StackPanel Orientation="Horizontal">
                        <Border Style="{StaticResource Bolha}" Width="30" Height="30"><TextBlock Text="&#xE964;" Style="{StaticResource Icone}" FontSize="13"/></Border>
                        <TextBlock Text="Memória RAM" Style="{StaticResource TituloCard}" Margin="10,0,0,0"/>
                        <TextBlock x:Name="txtRamGB" Style="{StaticResource Suave}" FontSize="13" VerticalAlignment="Center" Margin="14,0,0,0"/>
                      </StackPanel>
                    </DockPanel>
                    <Grid Margin="0,12,0,0" ClipToBounds="True">
                      <Line X1="0" X2="1" Stretch="Fill" VerticalAlignment="Center" Stroke="#1C1C20" StrokeDashArray="4 4"/>
                      <Canvas x:Name="cvRam">
                        <Polygon x:Name="agRam">
                          <Polygon.Fill><LinearGradientBrush StartPoint="0,0" EndPoint="0,1"><GradientStop Color="#557BC62D" Offset="0"/><GradientStop Color="#007BC62D" Offset="1"/></LinearGradientBrush></Polygon.Fill>
                        </Polygon>
                        <Polyline x:Name="lnRam" Stroke="#A6F23A" StrokeThickness="2.2" StrokeLineJoin="Round">
                          <Polyline.Effect><DropShadowEffect Color="#A6F23A" BlurRadius="10" ShadowDepth="0" Opacity="0.6"/></Polyline.Effect>
                        </Polyline>
                      </Canvas>
                    </Grid>
                  </DockPanel>
                </Border>

                <!-- Dica -->
                <Border Grid.Row="2" Grid.Column="1" Style="{StaticResource Card}">
                  <Border.Background>
                    <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                      <GradientStop Color="#1F2E12" Offset="0"/>
                      <GradientStop Color="#0F0F11" Offset="1"/>
                    </LinearGradientBrush>
                  </Border.Background>
                  <DockPanel>
                    <TextBlock DockPanel.Dock="Top" Text="DICA DO MONITOR" Foreground="#A6F23A" FontSize="11" FontWeight="Bold"/>
                    <Button x:Name="btnDica" DockPanel.Dock="Bottom" Content="Ir para Otimizar  →" HorizontalAlignment="Left" Padding="14,6" FontSize="12"/>
                    <StackPanel Margin="0,8,0,10">
                      <TextBlock x:Name="txtDicaTitulo" FontSize="19" FontWeight="SemiBold" TextWrapping="Wrap"/>
                      <TextBlock x:Name="txtDica" Style="{StaticResource Suave}" TextWrapping="Wrap" Margin="0,6,0,0"/>
                    </StackPanel>
                  </DockPanel>
                </Border>

                <!-- Placa de vídeo -->
                <Border Grid.Row="3" Grid.ColumnSpan="2" Style="{StaticResource Card}">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="190"/>
                      <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Viewbox Margin="0,0,18,0">
                      <Grid Width="180" Height="180">
                        <Ellipse Width="140" Height="140" Stroke="#1C2418" StrokeThickness="16"/>
                        <Path x:Name="anGpu" Stroke="#A6F23A" StrokeThickness="16" StrokeStartLineCap="Round" StrokeEndLineCap="Round">
                          <Path.Effect><DropShadowEffect Color="#A6F23A" BlurRadius="16" ShadowDepth="0" Opacity="0.6"/></Path.Effect>
                        </Path>
                        <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                          <TextBlock x:Name="txtGpuUso" Text="--" FontSize="32" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                          <TextBlock Text="uso da GPU" Style="{StaticResource Suave}" FontSize="11" HorizontalAlignment="Center"/>
                        </StackPanel>
                      </Grid>
                    </Viewbox>
                    <DockPanel Grid.Column="1">
                      <DockPanel DockPanel.Dock="Top">
                        <Border DockPanel.Dock="Right" Style="{StaticResource Chip}"><TextBlock x:Name="txtGpuNome" FontSize="11" Text="Detectando..."/></Border>
                        <StackPanel Orientation="Horizontal">
                          <Border Style="{StaticResource Bolha}" Width="30" Height="30"><TextBlock Text="&#xE7F4;" Style="{StaticResource Icone}" FontSize="13"/></Border>
                          <TextBlock Text="Placa de Vídeo" Style="{StaticResource TituloCard}" Margin="10,0,0,0"/>
                        </StackPanel>
                      </DockPanel>
                      <UniformGrid DockPanel.Dock="Top" Columns="4" Margin="0,14,0,0">
                        <StackPanel Orientation="Horizontal">
                          <Border x:Name="bdGpuTemp" Style="{StaticResource Bolha}"><TextBlock x:Name="icGpuTemp" Text="&#xE9CA;" Style="{StaticResource Icone}"/></Border>
                          <StackPanel Margin="10,0,0,0" VerticalAlignment="Center">
                            <TextBlock Text="Temperatura" Style="{StaticResource Suave}" FontSize="11"/>
                            <TextBlock x:Name="txtGpuTemp" Text="--" FontSize="17" FontWeight="SemiBold"/>
                          </StackPanel>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal">
                          <Border Style="{StaticResource Bolha}"><TextBlock Text="&#xE964;" Style="{StaticResource Icone}"/></Border>
                          <StackPanel Margin="10,0,0,0" VerticalAlignment="Center">
                            <TextBlock Text="Memória (VRAM)" Style="{StaticResource Suave}" FontSize="11"/>
                            <TextBlock x:Name="txtGpuVram" Text="--" FontSize="17" FontWeight="SemiBold"/>
                          </StackPanel>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal">
                          <Border Style="{StaticResource Bolha}"><TextBlock Text="&#xE945;" Style="{StaticResource Icone}"/></Border>
                          <StackPanel Margin="10,0,0,0" VerticalAlignment="Center">
                            <TextBlock Text="Consumo" Style="{StaticResource Suave}" FontSize="11"/>
                            <TextBlock x:Name="txtGpuWatts" Text="--" FontSize="17" FontWeight="SemiBold"/>
                          </StackPanel>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal">
                          <Border Style="{StaticResource Bolha}"><TextBlock Text="&#xEC4A;" Style="{StaticResource Icone}"/></Border>
                          <StackPanel Margin="10,0,0,0" VerticalAlignment="Center">
                            <TextBlock x:Name="txtGpuFan" Text="Clock" Style="{StaticResource Suave}" FontSize="11"/>
                            <TextBlock x:Name="txtGpuClock" Text="--" FontSize="17" FontWeight="SemiBold"/>
                          </StackPanel>
                        </StackPanel>
                      </UniformGrid>
                      <ProgressBar x:Name="pbVram" DockPanel.Dock="Top" Height="6" Margin="0,14,0,0" Foreground="#7BC62D"/>
                      <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" Margin="0,8,0,0">
                        <Rectangle Width="14" Height="3" Fill="#A6F23A" VerticalAlignment="Center"/>
                        <TextBlock Text="Uso (%)" Style="{StaticResource Suave}" FontSize="11" Margin="6,0,16,0"/>
                        <Rectangle Width="14" Height="3" Fill="#FFB020" VerticalAlignment="Center"/>
                        <TextBlock Text="Temperatura (°C)" Style="{StaticResource Suave}" FontSize="11" Margin="6,0,16,0"/>
                        <TextBlock x:Name="txtGpuFonte" Style="{StaticResource Suave}" FontSize="11"/>
                      </StackPanel>
                      <Grid Margin="0,10,0,0" ClipToBounds="True">
                        <Line X1="0" X2="1" Stretch="Fill" VerticalAlignment="Top" Stroke="#1C1C20" StrokeDashArray="4 4"/>
                        <Line X1="0" X2="1" Stretch="Fill" VerticalAlignment="Center" Stroke="#1C1C20" StrokeDashArray="4 4"/>
                        <Canvas x:Name="cvGpu">
                          <Polygon x:Name="agGpu">
                            <Polygon.Fill><LinearGradientBrush StartPoint="0,0" EndPoint="0,1"><GradientStop Color="#44A6F23A" Offset="0"/><GradientStop Color="#00A6F23A" Offset="1"/></LinearGradientBrush></Polygon.Fill>
                          </Polygon>
                          <Polyline x:Name="lnGpuT" Stroke="#FFB020" StrokeThickness="1.6" StrokeDashArray="3 2"/>
                          <Polyline x:Name="lnGpu" Stroke="#A6F23A" StrokeThickness="2.2" StrokeLineJoin="Round">
                            <Polyline.Effect><DropShadowEffect Color="#A6F23A" BlurRadius="10" ShadowDepth="0" Opacity="0.6"/></Polyline.Effect>
                          </Polyline>
                        </Canvas>
                      </Grid>
                    </DockPanel>
                  </Grid>
                </Border>

                <!-- Resumo de segurança -->
                <Border Grid.Row="3" Grid.Column="2" Style="{StaticResource Card}">
                  <DockPanel>
                    <TextBlock DockPanel.Dock="Top" Text="Segurança" Style="{StaticResource TituloCard}"/>
                    <Button x:Name="btnSegMini" DockPanel.Dock="Bottom" Style="{StaticResource BtnSecundario}" Content="Ver detalhes  →" HorizontalAlignment="Left" Padding="14,6" FontSize="12"/>
                    <Grid VerticalAlignment="Center" Margin="0,8">
                      <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition/></Grid.ColumnDefinitions>
                      <Border x:Name="bdSegMini" Width="60" Height="60" CornerRadius="30" Background="#1C2418" BorderBrush="#2C3A22" BorderThickness="1">
                        <TextBlock x:Name="icSegMini" Text="&#xEA18;" Style="{StaticResource Icone}" FontSize="26"/>
                      </Border>
                      <StackPanel Grid.Column="1" Margin="14,0,0,0" VerticalAlignment="Center">
                        <TextBlock x:Name="txtSegMini" Text="Verificando..." FontSize="17" FontWeight="SemiBold" TextWrapping="Wrap"/>
                        <TextBlock x:Name="txtSegMiniSub" Style="{StaticResource Suave}" TextWrapping="Wrap" Margin="0,4,0,0"/>
                      </StackPanel>
                    </Grid>
                  </DockPanel>
                </Border>
              </Grid>
            </ScrollViewer>
          </TabItem>

          <!-- ===================== PROCESSOS ===================== -->
          <TabItem>
            <Border Style="{StaticResource Card}" Margin="21,0,21,21">
              <DockPanel>
                <DockPanel DockPanel.Dock="Top" Margin="0,0,0,12">
                  <Button x:Name="btnFinalizar" DockPanel.Dock="Right" Content="Finalizar tarefa" Margin="12,0,0,0" Background="#FF5A4D" Foreground="#0B0B0D"/>
                  <Border DockPanel.Dock="Right" Background="#141416" CornerRadius="11" Padding="3" BorderBrush="#1F1F23" BorderThickness="1">
                    <StackPanel Orientation="Horizontal">
                      <RadioButton x:Name="rbRam" Style="{StaticResource Seg}" Content="RAM" GroupName="ordem" IsChecked="True"/>
                      <RadioButton x:Name="rbCpu" Style="{StaticResource Seg}" Content="CPU" GroupName="ordem"/>
                      <RadioButton x:Name="rbNome" Style="{StaticResource Seg}" Content="Nome" GroupName="ordem"/>
                    </StackPanel>
                  </Border>
                  <StackPanel>
                    <TextBlock Text="Processos" FontSize="22" FontWeight="SemiBold"/>
                    <TextBlock x:Name="txtProcInfo" Style="{StaticResource Suave}"/>
                  </StackPanel>
                </DockPanel>
                <TextBlock DockPanel.Dock="Bottom" Style="{StaticResource Suave}" Margin="0,10,0,0" TextWrapping="Wrap"
                           Text="Use a busca no topo para filtrar. Processos com o mesmo nome são agrupados (ex.: todas as abas do navegador); 'Finalizar tarefa' fecha todas as instâncias do grupo."/>
                <DataGrid x:Name="dgProc" RowHeight="40">
                  <DataGrid.Columns>
                    <DataGridTemplateColumn Header="Processo" Width="*">
                      <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                          <StackPanel Orientation="Horizontal">
                            <Border Width="28" Height="28" CornerRadius="8" Background="#1C2418">
                              <Image Source="{Binding Icone}" Width="18" Height="18"/>
                            </Border>
                            <TextBlock Text="{Binding Nome}" Margin="12,0,0,0" VerticalAlignment="Center" FontWeight="SemiBold"/>
                          </StackPanel>
                        </DataTemplate>
                      </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                    <DataGridTextColumn Header="Instâncias" Binding="{Binding Qtd}" Width="100"/>
                    <DataGridTemplateColumn Header="CPU" Width="200">
                      <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                          <Grid>
                            <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="58"/></Grid.ColumnDefinitions>
                            <ProgressBar Value="{Binding Cpu}" Height="6"/>
                            <TextBlock Grid.Column="1" Text="{Binding Cpu, StringFormat={}{0:0.0}%}" HorizontalAlignment="Right"/>
                          </Grid>
                        </DataTemplate>
                      </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                    <DataGridTemplateColumn Header="Memória" Width="230">
                      <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                          <Grid>
                            <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="90"/></Grid.ColumnDefinitions>
                            <ProgressBar Value="{Binding RamPct}" Height="6" Foreground="#7BC62D"/>
                            <TextBlock Grid.Column="1" Text="{Binding RamMB, StringFormat={}{0:N0} MB}" HorizontalAlignment="Right"/>
                          </Grid>
                        </DataTemplate>
                      </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                  </DataGrid.Columns>
                </DataGrid>
              </DockPanel>
            </Border>
          </TabItem>

          <!-- ===================== ALERTAS ===================== -->
          <TabItem>
            <Grid Margin="14,0,14,14">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="1.15*"/>
                <ColumnDefinition/>
              </Grid.ColumnDefinitions>
              <Border Style="{StaticResource Card}" Padding="26,22">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                  <StackPanel>
                    <TextBlock Text="Configurar alertas" FontSize="22" FontWeight="SemiBold"/>
                    <TextBlock Text="Receba um aviso do Windows quando o PC passar destes limites." Style="{StaticResource Suave}" Margin="0,4,0,14"/>
                    <CheckBox x:Name="chkAlertas" Content="Alertas ativados"/>
                    <CheckBox x:Name="chkSom" Content="Tocar som no alerta"/>
                    <CheckBox x:Name="chkIniciar" Content="Iniciar junto com o Windows (minimizado na bandeja)"/>

                    <Grid Margin="0,18,0,0">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="210"/>
                        <ColumnDefinition/>
                        <ColumnDefinition Width="80"/>
                      </Grid.ColumnDefinitions>
                      <Grid.RowDefinitions>
                        <RowDefinition Height="44"/><RowDefinition Height="44"/><RowDefinition Height="44"/>
                        <RowDefinition Height="44"/><RowDefinition Height="44"/><RowDefinition Height="44"/>
                        <RowDefinition Height="44"/>
                      </Grid.RowDefinitions>

                      <TextBlock Text="Processador acima de" VerticalAlignment="Center"/>
                      <Slider x:Name="slCpu" Grid.Column="1" Minimum="30" Maximum="100" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center"/>
                      <TextBlock Grid.Column="2" Text="{Binding Value, ElementName=slCpu, StringFormat={}{0:0}%}" VerticalAlignment="Center" HorizontalAlignment="Right" FontSize="16" FontWeight="SemiBold" Foreground="#A6F23A"/>

                      <TextBlock Grid.Row="1" Text="Memória RAM acima de" VerticalAlignment="Center"/>
                      <Slider x:Name="slRam" Grid.Row="1" Grid.Column="1" Minimum="30" Maximum="100" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center"/>
                      <TextBlock Grid.Row="1" Grid.Column="2" Text="{Binding Value, ElementName=slRam, StringFormat={}{0:0}%}" VerticalAlignment="Center" HorizontalAlignment="Right" FontSize="16" FontWeight="SemiBold" Foreground="#A6F23A"/>

                      <TextBlock Grid.Row="2" Text="Atividade do disco acima de" VerticalAlignment="Center"/>
                      <Slider x:Name="slDsc" Grid.Row="2" Grid.Column="1" Minimum="30" Maximum="100" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center"/>
                      <TextBlock Grid.Row="2" Grid.Column="2" Text="{Binding Value, ElementName=slDsc, StringFormat={}{0:0}%}" VerticalAlignment="Center" HorizontalAlignment="Right" FontSize="16" FontWeight="SemiBold" Foreground="#A6F23A"/>

                      <TextBlock Grid.Row="3" Text="Espaço livre abaixo de" VerticalAlignment="Center"/>
                      <Slider x:Name="slEsp" Grid.Row="3" Grid.Column="1" Minimum="2" Maximum="50" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center"/>
                      <TextBlock Grid.Row="3" Grid.Column="2" Text="{Binding Value, ElementName=slEsp, StringFormat={}{0:0}%}" VerticalAlignment="Center" HorizontalAlignment="Right" FontSize="16" FontWeight="SemiBold" Foreground="#A6F23A"/>

                      <TextBlock Grid.Row="4" Text="Só alertar após" VerticalAlignment="Center"/>
                      <Slider x:Name="slSus" Grid.Row="4" Grid.Column="1" Minimum="2" Maximum="120" TickFrequency="2" IsSnapToTickEnabled="True" VerticalAlignment="Center"/>
                      <TextBlock Grid.Row="4" Grid.Column="2" Text="{Binding Value, ElementName=slSus, StringFormat={}{0:0} s}" VerticalAlignment="Center" HorizontalAlignment="Right" FontSize="16" FontWeight="SemiBold" Foreground="#A6F23A"/>

                      <TextBlock Grid.Row="5" Text="Repetir alerta a cada" VerticalAlignment="Center"/>
                      <Slider x:Name="slRep" Grid.Row="5" Grid.Column="1" Minimum="1" Maximum="60" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center"/>
                      <TextBlock Grid.Row="5" Grid.Column="2" Text="{Binding Value, ElementName=slRep, StringFormat={}{0:0} min}" VerticalAlignment="Center" HorizontalAlignment="Right" FontSize="16" FontWeight="SemiBold" Foreground="#A6F23A"/>

                      <TextBlock Grid.Row="6" Text="Temperatura da GPU acima de" VerticalAlignment="Center"/>
                      <Slider x:Name="slGpuT" Grid.Row="6" Grid.Column="1" Minimum="50" Maximum="100" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center"/>
                      <TextBlock Grid.Row="6" Grid.Column="2" Text="{Binding Value, ElementName=slGpuT, StringFormat={}{0:0} °C}" VerticalAlignment="Center" HorizontalAlignment="Right" FontSize="16" FontWeight="SemiBold" Foreground="#A6F23A"/>
                    </Grid>

                    <TextBlock Style="{StaticResource Suave}" TextWrapping="Wrap" Margin="0,14,0,0"
                               Text="O alerta só dispara se o valor ficar acima do limite pelo tempo definido (ignora picos rápidos). Fechar a janela mantém o Monitor ativo na bandeja, perto do relógio."/>
                    <StackPanel Orientation="Horizontal" Margin="0,18,0,0">
                      <Button x:Name="btnSalvar" Content="Salvar" Margin="0,0,10,0"/>
                      <Button x:Name="btnTestar" Style="{StaticResource BtnSecundario}" Content="Testar alerta"/>
                    </StackPanel>
                    <TextBlock x:Name="txtSalvo" Foreground="#A6F23A" Margin="0,10,0,0"/>
                  </StackPanel>
                </ScrollViewer>
              </Border>
              <Border Grid.Column="1" Style="{StaticResource Card}" Padding="22,20">
                <DockPanel>
                  <DockPanel DockPanel.Dock="Top" Margin="0,0,0,10">
                    <Button x:Name="btnLimparHist" DockPanel.Dock="Right" Style="{StaticResource Link}" Content="Limpar" VerticalAlignment="Center"/>
                    <TextBlock Text="Histórico de alertas" Style="{StaticResource TituloCard}"/>
                  </DockPanel>
                  <ListBox x:Name="lbHist" Background="Transparent" BorderThickness="0" Foreground="#F2F2F5"
                           ScrollViewer.HorizontalScrollBarVisibility="Disabled">
                    <ListBox.ItemTemplate>
                      <DataTemplate>
                        <TextBlock Text="{Binding}" TextWrapping="Wrap"/>
                      </DataTemplate>
                    </ListBox.ItemTemplate>
                  </ListBox>
                </DockPanel>
              </Border>
            </Grid>
          </TabItem>

          <!-- ===================== VIDA ÚTIL ===================== -->
          <TabItem>
            <DockPanel Margin="14,0,14,14">
              <Border DockPanel.Dock="Top" Style="{StaticResource Card}" Padding="22,16">
                <DockPanel>
                  <Button x:Name="btnAnalisar" DockPanel.Dock="Right" Content="Analisar peças"/>
                  <Border DockPanel.Dock="Right" Background="#141416" BorderBrush="#1F1F23" BorderThickness="1" CornerRadius="11" Padding="12,0" Margin="0,0,12,0">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="Ano de compra" Style="{StaticResource Suave}" VerticalAlignment="Center"/>
                      <TextBox x:Name="txtAno" Width="54" MaxLength="4" Margin="10,0,0,0" FontWeight="SemiBold" Foreground="#A6F23A"/>
                    </StackPanel>
                  </Border>
                  <StackPanel>
                    <TextBlock Text="Vida útil das peças" FontSize="22" FontWeight="SemiBold"/>
                    <TextBlock x:Name="txtVidaStatus" Style="{StaticResource Suave}" TextTrimming="CharacterEllipsis"/>
                  </StackPanel>
                </DockPanel>
              </Border>

              <!-- Teste de velocidade do disco -->
              <Border DockPanel.Dock="Top" Style="{StaticResource Card}" Padding="22,16">
                <StackPanel>
                  <DockPanel>
                    <Button x:Name="btnTestarDisco" DockPanel.Dock="Right" Content="Testar velocidade" VerticalAlignment="Center"/>
                    <WrapPanel x:Name="wpDiscos" DockPanel.Dock="Right" VerticalAlignment="Center" Margin="0,0,12,-8"/>
                    <StackPanel>
                      <TextBlock Text="Velocidade do disco" Style="{StaticResource TituloCard}"/>
                      <TextBlock x:Name="txtDiscoInfo" Style="{StaticResource Suave}" TextWrapping="Wrap"
                                 Text="Mede leitura e gravação com a ferramenta do próprio Windows (winsat). Leva de 5 s (SSD) a 1 min (HD)."/>
                    </StackPanel>
                  </DockPanel>
                  <UniformGrid x:Name="ugDisco" Columns="4" Margin="0,14,0,0" Visibility="Collapsed">
                    <StackPanel><TextBlock Text="Leitura sequencial" Style="{StaticResource Suave}"/><TextBlock x:Name="txtDiscoLeitura" FontSize="22" FontWeight="SemiBold"/><TextBlock x:Name="txtDiscoLeituraAnt" Style="{StaticResource Suave}" FontSize="11"/></StackPanel>
                    <StackPanel><TextBlock Text="Gravação sequencial" Style="{StaticResource Suave}"/><TextBlock x:Name="txtDiscoGravacao" FontSize="22" FontWeight="SemiBold"/><TextBlock x:Name="txtDiscoGravacaoAnt" Style="{StaticResource Suave}" FontSize="11"/></StackPanel>
                    <StackPanel><TextBlock Text="Leitura aleatória (arquivos pequenos)" Style="{StaticResource Suave}"/><TextBlock x:Name="txtDiscoAleatoria" FontSize="22" FontWeight="SemiBold"/></StackPanel>
                    <StackPanel><TextBlock Text="Tempo de resposta (95%)" Style="{StaticResource Suave}"/><TextBlock x:Name="txtDiscoLatencia" FontSize="22" FontWeight="SemiBold"/></StackPanel>
                  </UniformGrid>
                  <TextBlock x:Name="txtDiscoVeredito" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,12,0,0" Visibility="Collapsed"/>
                </StackPanel>
              </Border>
              <TextBlock DockPanel.Dock="Bottom" Style="{StaticResource Suave}" TextWrapping="Wrap" Margin="8,4,8,0"
                         Text="SSD/HD e bateria usam dados reais do sensor (SMART) — o desgaste do SSD exige abrir como administrador. CPU, RAM, placa de vídeo, placa-mãe e fonte não informam desgaste: o valor em verde-água é uma ESTIMATIVA pela idade comparada à vida típica da peça."/>
              <ScrollViewer VerticalScrollBarVisibility="Auto">
                <ItemsControl x:Name="icVida">
                  <ItemsControl.ItemsPanel>
                    <ItemsPanelTemplate><UniformGrid Columns="2"/></ItemsPanelTemplate>
                  </ItemsControl.ItemsPanel>
                  <ItemsControl.ItemTemplate>
                    <DataTemplate>
                      <Border Style="{StaticResource Card}">
                        <StackPanel>
                          <DockPanel>
                            <TextBlock DockPanel.Dock="Right" Text="{Binding SaudeTexto}" Foreground="{Binding Cor}" FontSize="20" FontWeight="SemiBold" VerticalAlignment="Center"/>
                            <Border Style="{StaticResource Bolha}">
                              <TextBlock Text="{Binding Glifo}" Style="{StaticResource Icone}"/>
                            </Border>
                            <TextBlock Text="{Binding Componente}" Style="{StaticResource Suave}" FontSize="13" VerticalAlignment="Center" Margin="10,0,0,0"/>
                          </DockPanel>
                          <TextBlock Text="{Binding Modelo}" FontSize="16" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,10,0,10"/>
                          <ProgressBar Value="{Binding Saude}" Foreground="{Binding Cor}" Height="8"/>
                          <TextBlock Text="{Binding Detalhes}" Style="{StaticResource Suave}" TextWrapping="Wrap" Margin="0,10,0,0"/>
                          <TextBlock Text="{Binding Observacao}" TextWrapping="Wrap" Margin="0,6,0,0" Foreground="#C8C8CE"/>
                        </StackPanel>
                      </Border>
                    </DataTemplate>
                  </ItemsControl.ItemTemplate>
                </ItemsControl>
              </ScrollViewer>
            </DockPanel>
          </TabItem>

          <!-- ===================== OTIMIZAR ===================== -->
          <TabItem>
            <DockPanel Margin="14,0,14,14">
              <Border DockPanel.Dock="Top" Style="{StaticResource Card}" Padding="22,16">
                <StackPanel>
                  <TextBlock Text="Otimizar o PC" FontSize="22" FontWeight="SemiBold"/>
                  <TextBlock Style="{StaticResource Suave}" TextWrapping="Wrap" Margin="0,4,0,0"
                             Text="Cada comando abre em uma janela própria para você acompanhar. Os marcados com Admin pedem permissão do Windows. Antes dos avançados, crie um ponto de restauração (primeiro item do nível Avançado)."/>
                </StackPanel>
              </Border>
              <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel>
                  <!-- Manutenção automática -->
                  <Border Style="{StaticResource Card}" Padding="22,18">
                    <StackPanel>
                      <DockPanel>
                        <CheckBox x:Name="chkManut" DockPanel.Dock="Right" Content="Ativada" VerticalAlignment="Center"/>
                        <StackPanel>
                          <TextBlock Text="Manutenção automática" Style="{StaticResource TituloCard}"/>
                          <TextBlock Style="{StaticResource Suave}" TextWrapping="Wrap" Margin="0,2,0,0"
                                     Text="Uma vez por semana o Windows roda a limpeza sozinho — mesmo com o Monitor fechado — e avisa quanto espaço liberou. Se o PC estiver desligado no horário, ela roda assim que ligar."/>
                        </StackPanel>
                      </DockPanel>
                      <WrapPanel Margin="0,14,0,0">
                        <TextBlock Text="Toda" VerticalAlignment="Center" Margin="0,0,10,8"/>
                        <Border Background="#141416" CornerRadius="11" Padding="3" BorderBrush="#1F1F23" BorderThickness="1" Margin="0,0,12,8">
                          <StackPanel x:Name="spManutDias" Orientation="Horizontal"/>
                        </Border>
                        <TextBlock Text="às" VerticalAlignment="Center" Margin="0,0,10,8"/>
                        <Border Width="80" Height="34" CornerRadius="11" Background="#141416" BorderBrush="#1F1F23" BorderThickness="1" Margin="0,0,0,8">
                          <TextBox x:Name="txtManutHora" Margin="12,0" MaxLength="5" FontFamily="Consolas" HorizontalContentAlignment="Center"/>
                        </Border>
                      </WrapPanel>
                      <CheckBox x:Name="chkManutTemp" Content="Apagar arquivos temporários com mais de 2 dias"/>
                      <CheckBox x:Name="chkManutLixeira" Content="Esvaziar da Lixeira só o que foi apagado há mais de 30 dias"/>
                      <CheckBox x:Name="chkManutSistema" Content="Limpar Temp do Windows e cache do Windows Update (precisa abrir o Monitor como admin para agendar)"/>
                      <DockPanel Margin="0,12,0,0">
                        <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
                          <Button x:Name="btnManutAgora" Style="{StaticResource BtnSecundario}" Content="Executar agora" Margin="0,0,10,0"/>
                          <Button x:Name="btnManutSalvar" Content="Salvar agendamento"/>
                        </StackPanel>
                        <TextBlock x:Name="txtManutStatus" Style="{StaticResource Suave}" TextWrapping="Wrap" VerticalAlignment="Center" Margin="0,0,12,0"/>
                      </DockPanel>
                    </StackPanel>
                  </Border>
                  <StackPanel x:Name="spOtimizar"/>
                </StackPanel>
              </ScrollViewer>
            </DockPanel>
          </TabItem>

          <!-- ===================== SEGURANÇA ===================== -->
          <TabItem>
            <DockPanel Margin="14,0,14,14">
              <Border DockPanel.Dock="Top" Style="{StaticResource Card}" Padding="22,16">
                <DockPanel>
                  <Button x:Name="btnVerificar" DockPanel.Dock="Right" Content="Verificar agora" VerticalAlignment="Center"/>
                  <Border x:Name="bdSegIcone" Width="54" Height="54" CornerRadius="27" Background="#1C2418" BorderBrush="#2C3A22" BorderThickness="1" Margin="0,0,16,0">
                    <TextBlock x:Name="icSegIcone" Text="&#xEA18;" Style="{StaticResource Icone}" FontSize="24"/>
                  </Border>
                  <StackPanel VerticalAlignment="Center">
                    <TextBlock x:Name="txtSegTitulo" Text="Segurança" FontSize="22" FontWeight="SemiBold"/>
                    <TextBlock x:Name="txtSegResumo" Style="{StaticResource Suave}" TextWrapping="Wrap" Text="Verificando proteção do sistema e processos..."/>
                  </StackPanel>
                </DockPanel>
              </Border>
              <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel>
                  <ItemsControl x:Name="icProtecao">
                    <ItemsControl.ItemsPanel>
                      <ItemsPanelTemplate><UniformGrid Columns="4"/></ItemsPanelTemplate>
                    </ItemsControl.ItemsPanel>
                    <ItemsControl.ItemTemplate>
                      <DataTemplate>
                        <Border Style="{StaticResource Card}">
                          <DockPanel>
                            <Button DockPanel.Dock="Bottom" Content="{Binding AcaoTexto}" Tag="{Binding Acao}" Visibility="{Binding AcaoVisivel}"
                                    Style="{StaticResource BtnSecundario}" HorizontalAlignment="Left" Padding="12,5" FontSize="12" Margin="0,12,0,0"/>
                            <StackPanel>
                              <StackPanel Orientation="Horizontal">
                                <Border Style="{StaticResource Bolha}"><TextBlock Text="{Binding Glifo}" Style="{StaticResource Icone}" Foreground="{Binding Cor}"/></Border>
                                <TextBlock Text="{Binding Titulo}" Style="{StaticResource Suave}" FontSize="13" VerticalAlignment="Center" Margin="10,0,0,0"/>
                              </StackPanel>
                              <TextBlock Text="{Binding Status}" Foreground="{Binding Cor}" FontSize="17" FontWeight="SemiBold" Margin="0,12,0,0" TextWrapping="Wrap"/>
                              <TextBlock Text="{Binding Detalhe}" Style="{StaticResource Suave}" TextWrapping="Wrap" Margin="0,4,0,0"/>
                            </StackPanel>
                          </DockPanel>
                        </Border>
                      </DataTemplate>
                    </ItemsControl.ItemTemplate>
                  </ItemsControl>

                  <DockPanel Margin="10,16,10,4">
                    <TextBlock x:Name="txtAchadosInfo" DockPanel.Dock="Right" Style="{StaticResource Suave}" VerticalAlignment="Center"/>
                    <TextBlock Text="Processos em execução" FontSize="18" FontWeight="SemiBold"/>
                  </DockPanel>
                  <Border x:Name="bdSemAchados" Style="{StaticResource Card}" Visibility="Collapsed">
                    <StackPanel Orientation="Horizontal">
                      <Border Style="{StaticResource Bolha}"><TextBlock Text="&#xE73E;" Style="{StaticResource Icone}"/></Border>
                      <StackPanel Margin="12,0,0,0" VerticalAlignment="Center">
                        <TextBlock Text="Nenhum processo suspeito encontrado" FontSize="15" FontWeight="SemiBold"/>
                        <TextBlock x:Name="txtSemAchados" Style="{StaticResource Suave}"/>
                      </StackPanel>
                    </StackPanel>
                  </Border>
                  <ItemsControl x:Name="icAchados">
                    <ItemsControl.ItemTemplate>
                      <DataTemplate>
                        <Border Style="{StaticResource Card}" Padding="0">
                          <Grid>
                            <Grid.ColumnDefinitions><ColumnDefinition Width="5"/><ColumnDefinition/></Grid.ColumnDefinitions>
                            <Border Background="{Binding Cor}" CornerRadius="18,0,0,18"/>
                            <StackPanel Grid.Column="1" Margin="18,14,20,14">
                              <DockPanel>
                                <Border DockPanel.Dock="Right" CornerRadius="8" Padding="10,3" Background="{Binding CorFundo}" VerticalAlignment="Center">
                                  <TextBlock Text="{Binding Severidade}" Foreground="{Binding Cor}" FontSize="11" FontWeight="SemiBold"/>
                                </Border>
                                <TextBlock Text="{Binding Titulo}" FontSize="15" FontWeight="SemiBold" TextWrapping="Wrap"/>
                              </DockPanel>
                              <TextBlock Text="{Binding Linha}" Style="{StaticResource Suave}" Margin="0,3,0,0"/>
                              <TextBlock Text="{Binding Caminho}" FontFamily="Consolas" FontSize="11" Foreground="#C8C8CE" TextWrapping="Wrap" Margin="0,6,0,0"/>
                              <TextBlock Text="{Binding Motivo}" TextWrapping="Wrap" Margin="0,6,0,0" Foreground="#C8C8CE"/>
                              <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                                <Button Content="Abrir local" Tag="abrir" Style="{StaticResource BtnSecundario}" Padding="12,5" FontSize="12" Margin="0,0,8,0"/>
                                <Button Content="Ver no VirusTotal" Tag="vt" Style="{StaticResource BtnSecundario}" Padding="12,5" FontSize="12" Margin="0,0,8,0"/>
                                <Button Content="Finalizar processo" Tag="fim" Padding="12,5" FontSize="12" Background="#FF5A4D"/>
                              </StackPanel>
                            </StackPanel>
                          </Grid>
                        </Border>
                      </DataTemplate>
                    </ItemsControl.ItemTemplate>
                  </ItemsControl>
                  <TextBlock Style="{StaticResource Suave}" TextWrapping="Wrap" Margin="10,10,10,0"
                             Text="Como funciona: confere se os processos do Windows (svchost, lsass, csrss...) estão na pasta certa, procura nomes que imitam processos do sistema, programas rodando de pastas suspeitas (Temp, Downloads, Lixeira) e arquivos sem assinatura digital. 'Ver no VirusTotal' envia só a impressão digital (hash) do arquivo, nunca o arquivo. A verificação roda sozinha a cada 15 minutos."/>
                </StackPanel>
              </ScrollViewer>
            </DockPanel>
          </TabItem>

          <!-- ===================== APLICATIVOS ===================== -->
          <TabItem>
            <DockPanel Margin="14,0,14,14">
              <Border DockPanel.Dock="Top" Style="{StaticResource Card}" Padding="22,16">
                <StackPanel>
                  <DockPanel>
                    <Button x:Name="btnAppsAtualizar" DockPanel.Dock="Right" Style="{StaticResource BtnSecundario}" Content="Atualizar lista" VerticalAlignment="Center"/>
                    <StackPanel>
                      <TextBlock Text="Aplicativos" FontSize="22" FontWeight="SemiBold"/>
                      <TextBlock x:Name="txtAppsInfo" Style="{StaticResource Suave}" Text="Carregando programas instalados..."/>
                    </StackPanel>
                  </DockPanel>
                  <DockPanel Margin="0,14,0,0">
                    <Border DockPanel.Dock="Right" Background="#141416" CornerRadius="11" Padding="3" BorderBrush="#1F1F23" BorderThickness="1" Margin="10,0,0,0">
                      <StackPanel Orientation="Horizontal">
                        <RadioButton x:Name="rbAppNome" Style="{StaticResource Seg}" Content="Nome" GroupName="appordem" IsChecked="True"/>
                        <RadioButton x:Name="rbAppTam" Style="{StaticResource Seg}" Content="Tamanho" GroupName="appordem"/>
                        <RadioButton x:Name="rbAppData" Style="{StaticResource Seg}" Content="Mais recentes" GroupName="appordem"/>
                      </StackPanel>
                    </Border>
                    <Border DockPanel.Dock="Right" Background="#141416" CornerRadius="11" Padding="3" BorderBrush="#1F1F23" BorderThickness="1" Margin="10,0,0,0">
                      <StackPanel Orientation="Horizontal">
                        <RadioButton x:Name="rbAppTodos" Style="{StaticResource Seg}" Content="Todos" GroupName="apptipo" IsChecked="True"/>
                        <RadioButton x:Name="rbAppProg" Style="{StaticResource Seg}" Content="Programas" GroupName="apptipo"/>
                        <RadioButton x:Name="rbAppStore" Style="{StaticResource Seg}" Content="Microsoft Store" GroupName="apptipo"/>
                      </StackPanel>
                    </Border>
                    <Border Height="38" CornerRadius="19" Background="#141416" BorderBrush="#1F1F23" BorderThickness="1">
                      <Grid Margin="14,0">
                        <TextBlock Text="&#xE721;" Style="{StaticResource Icone}" Foreground="#8E8E96" HorizontalAlignment="Left" FontSize="13"/>
                        <TextBox x:Name="txtBuscaApp" Margin="24,0,0,0"/>
                        <TextBlock x:Name="phBuscaApp" Text="Buscar aplicativo ou editor..." Foreground="#5E5E66" Margin="26,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
                      </Grid>
                    </Border>
                  </DockPanel>
                </StackPanel>
              </Border>

              <!-- Restos de programas desinstalados -->
              <ItemsControl x:Name="icRestos" DockPanel.Dock="Top">
                <ItemsControl.ItemTemplate>
                  <DataTemplate>
                    <Border Style="{StaticResource Card}" Padding="0">
                      <Border.Background>
                        <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                          <GradientStop Color="#2A2412" Offset="0"/>
                          <GradientStop Color="#0F0F11" Offset="1"/>
                        </LinearGradientBrush>
                      </Border.Background>
                      <Grid>
                        <Grid.ColumnDefinitions><ColumnDefinition Width="5"/><ColumnDefinition/></Grid.ColumnDefinitions>
                        <Border Background="#FFB020" CornerRadius="18,0,0,18"/>
                        <DockPanel Grid.Column="1" Margin="18,14,20,14">
                          <Button DockPanel.Dock="Right" Content="Limpar restos" Tag="limpar" Background="#FFB020" VerticalAlignment="Center" Margin="14,0,0,0"/>
                          <StackPanel>
                            <TextBlock Text="{Binding Titulo}" FontSize="15" FontWeight="SemiBold" Foreground="#FFC95C"/>
                            <TextBlock Text="{Binding Texto}" TextWrapping="Wrap" Margin="0,4,0,0" Foreground="#C8C8CE"/>
                            <TextBlock Text="{Binding Itens}" FontFamily="Consolas" FontSize="11" Foreground="#8E8E96" TextWrapping="Wrap" Margin="0,6,0,0"/>
                          </StackPanel>
                        </DockPanel>
                      </Grid>
                    </Border>
                  </DataTemplate>
                </ItemsControl.ItemTemplate>
              </ItemsControl>

              <Border Style="{StaticResource Card}" Padding="14,10">
                <DataGrid x:Name="dgApps" RowHeight="56">
                  <DataGrid.Columns>
                    <DataGridTemplateColumn Header="Aplicativo" Width="*">
                      <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                          <DockPanel>
                            <Border DockPanel.Dock="Left" Width="36" Height="36" CornerRadius="10" Background="#1C2418">
                              <Grid>
                                <TextBlock Text="&#xE71D;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" Foreground="#A6F23A" HorizontalAlignment="Center" VerticalAlignment="Center">
                                  <TextBlock.Style>
                                    <Style TargetType="TextBlock">
                                      <Setter Property="Visibility" Value="Collapsed"/>
                                      <Style.Triggers>
                                        <DataTrigger Binding="{Binding Icone}" Value="{x:Null}"><Setter Property="Visibility" Value="Visible"/></DataTrigger>
                                      </Style.Triggers>
                                    </Style>
                                  </TextBlock.Style>
                                </TextBlock>
                                <Image Source="{Binding Icone}" Width="22" Height="22"/>
                              </Grid>
                            </Border>
                            <Border DockPanel.Dock="Right" CornerRadius="8" Padding="8,2" Background="#33FFB020" VerticalAlignment="Center" Visibility="{Binding TagVisivel}" Margin="8,0,0,0">
                              <TextBlock Text="{Binding Tag}" FontSize="11" Foreground="#FFC95C"/>
                            </Border>
                            <StackPanel Margin="12,0,0,0" VerticalAlignment="Center">
                              <TextBlock Text="{Binding Nome}" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
                              <TextBlock Text="{Binding Sub}" Foreground="#8E8E96" FontSize="11" TextTrimming="CharacterEllipsis"/>
                            </StackPanel>
                          </DockPanel>
                        </DataTemplate>
                      </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                    <DataGridTextColumn Header="Tipo" Binding="{Binding Tipo}" Width="130"/>
                    <DataGridTextColumn Header="Tamanho" Binding="{Binding TamanhoTxt}" Width="100"/>
                    <DataGridTextColumn Header="Instalado em" Binding="{Binding DataTxt}" Width="110"/>
                    <DataGridTemplateColumn Header="" Width="130">
                      <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                          <Button Content="Desinstalar" Tag="desinstalar" Style="{StaticResource BtnSecundario}" Foreground="#FF7A6E" Padding="12,5" FontSize="12" HorizontalAlignment="Right"/>
                        </DataTemplate>
                      </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                  </DataGrid.Columns>
                </DataGrid>
              </Border>
            </DockPanel>
          </TabItem>

          <!-- ===================== APARÊNCIA ===================== -->
          <TabItem>
            <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="14,0,14,14">
              <StackPanel>
                <Border Style="{StaticResource Card}" Padding="22,16">
                  <DockPanel>
                    <Button x:Name="btnAplicarTema" DockPanel.Dock="Right" Content="Aplicar e reiniciar" VerticalAlignment="Center"/>
                    <StackPanel>
                      <TextBlock Text="Aparência" FontSize="22" FontWeight="SemiBold"/>
                      <TextBlock Style="{StaticResource Suave}" TextWrapping="Wrap"
                                 Text="Escolha a cor de destaque e o estilo. Ao aplicar, o Monitor reinicia em 2 segundos com o novo visual."/>
                    </StackPanel>
                  </DockPanel>
                </Border>

                <Border Style="{StaticResource Card}" Padding="26,22">
                  <StackPanel>
                    <TextBlock Text="Cor de destaque" Style="{StaticResource TituloCard}"/>
                    <TextBlock Text="Usada nos anéis, gráficos, botões e no menu." Style="{StaticResource Suave}" Margin="0,4,0,0"/>
                    <WrapPanel x:Name="wpCores" Margin="0,18,0,0"/>
                    <DockPanel Margin="0,20,0,0">
                      <TextBlock Text="Personalizada" VerticalAlignment="Center" Margin="0,0,14,0"/>
                      <Border Width="140" Height="38" CornerRadius="19" Background="#141416" BorderBrush="#1F1F23" BorderThickness="1" HorizontalAlignment="Left">
                        <TextBox x:Name="txtCorHex" Margin="14,0" MaxLength="7" FontFamily="Consolas"/>
                      </Border>
                      <Ellipse x:Name="elCorPrev" Width="30" Height="30" Margin="12,0,0,0" HorizontalAlignment="Left" Stroke="#2A2A2F" StrokeThickness="1"/>
                      <TextBlock x:Name="txtCorAviso" Style="{StaticResource Suave}" VerticalAlignment="Center" Margin="12,0,0,0" Text="Cor atual · ou digite outra, ex.: #FF6B00"/>
                    </DockPanel>
                  </StackPanel>
                </Border>

                <Border Style="{StaticResource Card}" Padding="26,22">
                  <StackPanel>
                    <TextBlock Text="Estilo" Style="{StaticResource TituloCard}"/>
                    <CheckBox x:Name="chkBrilho" Content="Brilho neon nos anéis, linhas e no chip" Margin="0,14,0,0"/>
                    <TextBlock Style="{StaticResource Suave}" Margin="50,0,0,0" TextWrapping="Wrap"
                               Text="Desligado, o visual fica mais limpo e plano (e o Monitor usa um pouco menos de placa de vídeo)."/>
                  </StackPanel>
                </Border>

                <Border Style="{StaticResource Card}" Padding="26,22">
                  <DockPanel>
                    <Image x:Name="imgIcone" DockPanel.Dock="Left" Width="72" Height="72" Margin="0,0,20,0" VerticalAlignment="Top"/>
                    <StackPanel>
                      <TextBlock Text="Ícone e atalhos" Style="{StaticResource TituloCard}"/>
                      <TextBlock Style="{StaticResource Suave}" TextWrapping="Wrap" Margin="0,4,0,0"
                                 Text="O ícone acompanha a cor de destaque. Crie um atalho para abrir o Monitor com dois cliques, sem a janela preta do PowerShell."/>
                      <StackPanel Orientation="Horizontal" Margin="0,14,0,0">
                        <Button x:Name="btnAtalhoDesktop" Style="{StaticResource BtnSecundario}" Content="Criar atalho na Área de Trabalho" Margin="0,0,10,0"/>
                        <Button x:Name="btnAtalhoIniciar" Style="{StaticResource BtnSecundario}" Content="Adicionar ao menu Iniciar"/>
                      </StackPanel>
                      <TextBlock x:Name="txtAtalho" Foreground="#A6F23A" Margin="0,10,0,0"/>
                    </StackPanel>
                  </DockPanel>
                </Border>

                <!-- Sobre e atualizações -->
                <Border Style="{StaticResource Card}" Padding="26,22">
                  <DockPanel>
                    <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
                      <Button x:Name="btnVerificarAtualizacao" Style="{StaticResource BtnSecundario}" Content="Verificar atualização" Margin="0,0,10,0"/>
                      <Button x:Name="btnInstalarAtualizacao" Content="Atualizar agora" Visibility="Collapsed"/>
                    </StackPanel>
                    <StackPanel>
                      <TextBlock x:Name="txtVersao" Style="{StaticResource TituloCard}"/>
                      <TextBlock x:Name="txtAtualizacao" Style="{StaticResource Suave}" TextWrapping="Wrap" Margin="0,4,0,0"/>
                    </StackPanel>
                  </DockPanel>
                </Border>

                <Border Style="{StaticResource Card}" Padding="26,22">
                  <StackPanel>
                    <TextBlock Text="Prévia" Style="{StaticResource TituloCard}"/>
                    <StackPanel Orientation="Horizontal" Margin="0,16,0,0">
                      <Border x:Name="bdPrevBotao" CornerRadius="11" Padding="18,8">
                        <TextBlock Text="Botão" FontWeight="SemiBold" Foreground="#0B0B0D"/>
                      </Border>
                      <Border x:Name="bdPrevMenu" CornerRadius="12" Padding="14,9" Margin="14,0,0,0" BorderThickness="1">
                        <TextBlock x:Name="txtPrevMenu" Text="Item do menu ativo"/>
                      </Border>
                      <Grid Width="200" Margin="18,0,0,0" VerticalAlignment="Center">
                        <Border Height="8" CornerRadius="4" Background="#1C1C20"/>
                        <Border x:Name="bdPrevBarra" Height="8" CornerRadius="4" Width="130" HorizontalAlignment="Left"/>
                      </Grid>
                      <Grid Width="54" Height="54" Margin="18,0,0,0">
                        <Ellipse x:Name="elPrevTrilha" StrokeThickness="7"/>
                        <Path x:Name="anPrev" StrokeThickness="7" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                      </Grid>
                    </StackPanel>
                  </StackPanel>
                </Border>
              </StackPanel>
            </ScrollViewer>
          </TabItem>

          <!-- ===================== JOGOS E FPS ===================== -->
          <TabItem>
            <ScrollViewer x:Name="svJogos" VerticalScrollBarVisibility="Auto" Padding="14,0,14,14">
              <StackPanel>
                <Border Style="{StaticResource Card}" Padding="22,16">
                  <DockPanel>
                    <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
                      <Button x:Name="btnFpsBaixar" Content="Baixar medidor de FPS (Intel, 1 MB)" Visibility="Collapsed" Margin="0,0,12,0"/>
                      <Button x:Name="btnFpsAdmin" Style="{StaticResource BtnSecundario}" Content="Abrir como admin" Visibility="Collapsed" Margin="0,0,12,0"/>
                      <CheckBox x:Name="chkOverlay" Content="FPS na tela do jogo" VerticalAlignment="Center" Margin="0,0,10,0"/>
                      <Border Background="#141416" CornerRadius="11" Padding="3" BorderBrush="#1F1F23" BorderThickness="1" Margin="0,0,18,0" VerticalAlignment="Center" ToolTip="Canto da tela">
                        <StackPanel Orientation="Horizontal">
                          <RadioButton x:Name="rbOvSE" Style="{StaticResource Seg}" GroupName="ovcanto" Content="↖" Tag="SuperiorEsquerdo" Padding="0"/>
                          <RadioButton x:Name="rbOvSD" Style="{StaticResource Seg}" GroupName="ovcanto" Content="↗" Tag="SuperiorDireito"/>
                          <RadioButton x:Name="rbOvIE" Style="{StaticResource Seg}" GroupName="ovcanto" Content="↙" Tag="InferiorEsquerdo"/>
                          <RadioButton x:Name="rbOvID" Style="{StaticResource Seg}" GroupName="ovcanto" Content="↘" Tag="InferiorDireito"/>
                        </StackPanel>
                      </Border>
                      <CheckBox x:Name="chkFps" Content="Medir FPS" VerticalAlignment="Center"/>
                    </StackPanel>
                    <StackPanel>
                      <TextBlock Text="Jogos e FPS" FontSize="22" FontWeight="SemiBold"/>
                      <TextBlock x:Name="txtFpsStatus" Style="{StaticResource Suave}" TextWrapping="Wrap"/>
                    </StackPanel>
                  </DockPanel>
                </Border>

                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="1.35*"/>
                    <ColumnDefinition/>
                  </Grid.ColumnDefinitions>

                  <!-- FPS ao vivo -->
                  <Border Style="{StaticResource Card}">
                    <StackPanel>
                      <DockPanel>
                        <Border DockPanel.Dock="Right" Style="{StaticResource Chip}"><TextBlock x:Name="txtFpsModo" FontSize="11" Text="aguardando jogo"/></Border>
                        <Border Width="36" Height="36" CornerRadius="10" Background="#1C2418">
                          <Grid>
                            <TextBlock x:Name="icJogoGlifo" Text="&#xE7FC;" Style="{StaticResource Icone}"/>
                            <Image x:Name="imgJogo" Width="22" Height="22"/>
                          </Grid>
                        </Border>
                        <StackPanel Margin="12,0,0,0" VerticalAlignment="Center">
                          <TextBlock x:Name="txtJogoNome" Text="Nenhum jogo detectado" FontSize="16" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
                          <TextBlock x:Name="txtJogoSub" Style="{StaticResource Suave}" Text="Abra um jogo — o Monitor detecta sozinho"/>
                        </StackPanel>
                      </DockPanel>
                      <Grid Margin="0,16,0,0">
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="Auto"/>
                          <ColumnDefinition/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Orientation="Horizontal">
                          <TextBlock x:Name="txtFps" Text="--" FontSize="64" FontWeight="SemiBold" Foreground="#A6F23A"/>
                          <TextBlock Text="FPS" FontSize="16" Foreground="#8E8E96" VerticalAlignment="Bottom" Margin="8,0,0,16"/>
                        </StackPanel>
                        <UniformGrid Grid.Column="1" Columns="3" Margin="24,0,0,0" VerticalAlignment="Center">
                          <StackPanel>
                            <TextBlock Text="Média (1 min)" Style="{StaticResource Suave}"/>
                            <TextBlock x:Name="txtFpsMedia" Text="--" FontSize="20" FontWeight="SemiBold"/>
                          </StackPanel>
                          <StackPanel>
                            <TextBlock Text="1% mais baixos" Style="{StaticResource Suave}"/>
                            <TextBlock x:Name="txtFpsLow" Text="--" FontSize="20" FontWeight="SemiBold"/>
                          </StackPanel>
                          <StackPanel>
                            <TextBlock Text="Tempo de quadro" Style="{StaticResource Suave}"/>
                            <TextBlock x:Name="txtFrameMs" Text="--" FontSize="20" FontWeight="SemiBold"/>
                          </StackPanel>
                        </UniformGrid>
                      </Grid>
                      <Grid Height="150" Margin="0,14,0,0" ClipToBounds="True">
                        <Line X1="0" X2="1" Stretch="Fill" VerticalAlignment="Top" Stroke="#1C1C20" StrokeDashArray="4 4"/>
                        <Line X1="0" X2="1" Stretch="Fill" VerticalAlignment="Center" Stroke="#1C1C20" StrokeDashArray="4 4"/>
                        <Canvas x:Name="cvFps">
                          <Polygon x:Name="agFps">
                            <Polygon.Fill><LinearGradientBrush StartPoint="0,0" EndPoint="0,1"><GradientStop Color="#44A6F23A" Offset="0"/><GradientStop Color="#00A6F23A" Offset="1"/></LinearGradientBrush></Polygon.Fill>
                          </Polygon>
                          <Polyline x:Name="lnFps" Stroke="#A6F23A" StrokeThickness="2.2" StrokeLineJoin="Round">
                            <Polyline.Effect><DropShadowEffect Color="#A6F23A" BlurRadius="10" ShadowDepth="0" Opacity="0.6"/></Polyline.Effect>
                          </Polyline>
                        </Canvas>
                        <Canvas x:Name="cvQuedas" IsHitTestVisible="False"/>
                        <TextBlock x:Name="txtFpsEscala" Style="{StaticResource Suave}" FontSize="10" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="2,2,0,0"/>
                      </Grid>
                      <TextBlock Style="{StaticResource Suave}" FontSize="11" Margin="0,6,0,0" Text="Últimos 2 minutos · marcas vermelhas = quedas de FPS detectadas"/>
                    </StackPanel>
                  </Border>

                  <!-- Gargalo e sinais -->
                  <Border Grid.Column="1" Style="{StaticResource Card}">
                    <StackPanel>
                      <TextBlock Text="O que está limitando agora" Style="{StaticResource TituloCard}"/>
                      <TextBlock x:Name="txtGargalo" FontSize="17" FontWeight="SemiBold" Margin="0,10,0,0" TextWrapping="Wrap" Text="--"/>
                      <TextBlock x:Name="txtGargaloDesc" Style="{StaticResource Suave}" TextWrapping="Wrap" Margin="0,4,0,0"/>
                      <UniformGrid Columns="2" Margin="0,14,0,0">
                        <Border Background="#141416" CornerRadius="12" Padding="12,9" Margin="0,0,6,6"><StackPanel><TextBlock Text="Processador" Style="{StaticResource Suave}"/><TextBlock x:Name="txtSCpu" Text="--" FontSize="16" FontWeight="SemiBold"/></StackPanel></Border>
                        <Border Background="#141416" CornerRadius="12" Padding="12,9" Margin="0,0,0,6"><StackPanel><TextBlock Text="Núcleo mais usado" Style="{StaticResource Suave}"/><TextBlock x:Name="txtSNucleo" Text="--" FontSize="16" FontWeight="SemiBold"/></StackPanel></Border>
                        <Border Background="#141416" CornerRadius="12" Padding="12,9" Margin="0,0,6,6"><StackPanel><TextBlock Text="Placa de vídeo" Style="{StaticResource Suave}"/><TextBlock x:Name="txtSGpu" Text="--" FontSize="16" FontWeight="SemiBold"/></StackPanel></Border>
                        <Border Background="#141416" CornerRadius="12" Padding="12,9" Margin="0,0,0,6"><StackPanel><TextBlock Text="Temperatura da GPU" Style="{StaticResource Suave}"/><TextBlock x:Name="txtSGpuT" Text="--" FontSize="16" FontWeight="SemiBold"/></StackPanel></Border>
                        <Border Background="#141416" CornerRadius="12" Padding="12,9" Margin="0,0,6,6"><StackPanel><TextBlock Text="Memória de vídeo" Style="{StaticResource Suave}"/><TextBlock x:Name="txtSVram" Text="--" FontSize="16" FontWeight="SemiBold"/></StackPanel></Border>
                        <Border Background="#141416" CornerRadius="12" Padding="12,9" Margin="0,0,0,6"><StackPanel><TextBlock Text="Memória RAM" Style="{StaticResource Suave}"/><TextBlock x:Name="txtSRam" Text="--" FontSize="16" FontWeight="SemiBold"/></StackPanel></Border>
                        <Border Background="#141416" CornerRadius="12" Padding="12,9" Margin="0,0,6,0"><StackPanel><TextBlock Text="Disco" Style="{StaticResource Suave}"/><TextBlock x:Name="txtSDsc" Text="--" FontSize="16" FontWeight="SemiBold"/></StackPanel></Border>
                        <Border Background="#141416" CornerRadius="12" Padding="12,9"><StackPanel><TextBlock Text="Mais pesado em 2º plano" Style="{StaticResource Suave}"/><TextBlock x:Name="txtSFundo" Text="--" FontSize="16" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/></StackPanel></Border>
                      </UniformGrid>
                    </StackPanel>
                  </Border>
                </Grid>

                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="1.35*"/>
                    <ColumnDefinition/>
                  </Grid.ColumnDefinitions>

                  <!-- Modo Jogo automático -->
                  <Border Style="{StaticResource Card}">
                    <StackPanel>
                      <DockPanel>
                        <CheckBox x:Name="chkModoJogo" DockPanel.Dock="Right" Content="Ativado" VerticalAlignment="Center"/>
                        <StackPanel>
                          <TextBlock Text="Modo Jogo automático" Style="{StaticResource TituloCard}"/>
                          <TextBlock Style="{StaticResource Suave}" TextWrapping="Wrap" Margin="0,2,0,0"
                                     Text="Quando um jogo abrir, o Monitor prepara o PC sozinho — e desfaz tudo quando você fechar o jogo."/>
                        </StackPanel>
                      </DockPanel>
                      <Border x:Name="bdModoStatus" CornerRadius="12" Padding="14,10" Margin="0,12,0,4" Background="#141416">
                        <TextBlock x:Name="txtModoStatus" TextWrapping="Wrap" Text="Aguardando um jogo abrir."/>
                      </Border>
                      <CheckBox x:Name="chkModoPlano" Content="Plano de energia de alto desempenho durante o jogo" Margin="0,8,0,2"/>
                      <CheckBox x:Name="chkModoIndex" Content="Pausar a indexação de arquivos (libera o disco)"/>
                      <CheckBox x:Name="chkModoReabrir" Content="Reabrir os programas fechados quando o jogo terminar"/>
                      <DockPanel Margin="0,12,0,6">
                        <Button x:Name="btnModoListar" DockPanel.Dock="Right" Style="{StaticResource Link}" Content="Atualizar lista" VerticalAlignment="Center"/>
                        <TextBlock Text="Fechar estes programas ao abrir um jogo:" FontWeight="SemiBold"/>
                      </DockPanel>
                      <WrapPanel x:Name="wpModoFechar"/>
                      <TextBlock Style="{StaticResource Suave}" FontSize="11" TextWrapping="Wrap" Margin="0,2,0,0"
                                 Text="Clique para marcar. Eles são fechados como se você clicasse no X; os que não fecharem em 5 s são encerrados (trabalho não salvo neles é perdido)."/>
                    </StackPanel>
                  </Border>

                  <!-- Conexão -->
                  <Border Grid.Column="1" Style="{StaticResource Card}">
                    <StackPanel>
                      <DockPanel>
                        <Border DockPanel.Dock="Right" Width="150" Height="32" CornerRadius="16" Background="#141416" BorderBrush="#1F1F23" BorderThickness="1" VerticalAlignment="Center">
                          <TextBox x:Name="txtPingAlvo" Margin="12,0" FontFamily="Consolas" FontSize="12" ToolTip="Endereço medido (ex.: IP do servidor do jogo)"/>
                        </Border>
                        <TextBlock Text="Conexão" Style="{StaticResource TituloCard}"/>
                      </DockPanel>
                      <Grid Margin="0,12,0,0">
                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition/></Grid.ColumnDefinitions>
                        <StackPanel Orientation="Horizontal">
                          <TextBlock x:Name="txtPing" Text="--" FontSize="40" FontWeight="SemiBold" Foreground="#A6F23A"/>
                          <TextBlock Text="ms" FontSize="15" Foreground="#8E8E96" VerticalAlignment="Bottom" Margin="6,0,0,10"/>
                        </StackPanel>
                        <UniformGrid Grid.Column="1" Columns="3" Margin="20,0,0,0" VerticalAlignment="Center">
                          <StackPanel><TextBlock Text="Variação" Style="{StaticResource Suave}"/><TextBlock x:Name="txtJitter" Text="--" FontSize="16" FontWeight="SemiBold"/></StackPanel>
                          <StackPanel><TextBlock Text="Perda" Style="{StaticResource Suave}"/><TextBlock x:Name="txtPerda" Text="--" FontSize="16" FontWeight="SemiBold"/></StackPanel>
                          <StackPanel><TextBlock Text="Roteador" Style="{StaticResource Suave}"/><TextBlock x:Name="txtPingRot" Text="--" FontSize="16" FontWeight="SemiBold"/></StackPanel>
                        </UniformGrid>
                      </Grid>
                      <Grid Height="80" Margin="0,10,0,0" ClipToBounds="True">
                        <Line X1="0" X2="1" Stretch="Fill" VerticalAlignment="Center" Stroke="#1C1C20" StrokeDashArray="4 4"/>
                        <Canvas x:Name="cvPing">
                          <Polygon x:Name="agPing">
                            <Polygon.Fill><LinearGradientBrush StartPoint="0,0" EndPoint="0,1"><GradientStop Color="#33A6F23A" Offset="0"/><GradientStop Color="#00A6F23A" Offset="1"/></LinearGradientBrush></Polygon.Fill>
                          </Polygon>
                          <Polyline x:Name="lnPing" Stroke="#A6F23A" StrokeThickness="1.8" StrokeLineJoin="Round"/>
                        </Canvas>
                        <Canvas x:Name="cvPerdas" IsHitTestVisible="False"/>
                      </Grid>
                      <TextBlock x:Name="txtPingEscala" Style="{StaticResource Suave}" FontSize="10" Margin="0,4,0,0"/>
                      <TextBlock x:Name="txtPingDiag" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,10,0,0"/>
                      <TextBlock x:Name="txtPingDiagDesc" Style="{StaticResource Suave}" TextWrapping="Wrap" Margin="0,2,0,0"/>
                    </StackPanel>
                  </Border>
                </Grid>

                <!-- Sessões de jogo -->
                <Border Style="{StaticResource Card}">
                  <StackPanel>
                    <DockPanel>
                      <Button x:Name="btnLimparSessoes" DockPanel.Dock="Right" Style="{StaticResource Link}" Content="Apagar histórico" VerticalAlignment="Center"/>
                      <TextBlock Text="Sessões de jogo" Style="{StaticResource TituloCard}"/>
                    </DockPanel>
                    <TextBlock x:Name="txtSemSessoes" Style="{StaticResource Suave}" Margin="0,8,0,0" TextWrapping="Wrap"
                               Text="Cada vez que você joga por mais de 1 minuto, o resumo aparece aqui — dá para comparar antes e depois de uma otimização."/>
                    <ItemsControl x:Name="icSessoes" Margin="0,6,0,0">
                      <ItemsControl.ItemTemplate>
                        <DataTemplate>
                          <Border Background="#141416" CornerRadius="12" Padding="14,10" Margin="0,4">
                            <Grid>
                              <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="2*"/>
                                <ColumnDefinition/><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/>
                              </Grid.ColumnDefinitions>
                              <StackPanel VerticalAlignment="Center">
                                <TextBlock Text="{Binding Jogo}" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
                                <TextBlock Text="{Binding Quando}" Foreground="#8E8E96" FontSize="11"/>
                              </StackPanel>
                              <StackPanel Grid.Column="1"><TextBlock Text="FPS médio" Foreground="#8E8E96" FontSize="11"/>
                                <TextBlock><Run Text="{Binding FpsTxt}" FontWeight="SemiBold" FontSize="15"/><Run Text="{Binding Delta}" Foreground="{Binding DeltaCor}" FontSize="11"/></TextBlock></StackPanel>
                              <StackPanel Grid.Column="2"><TextBlock Text="1% mais baixos" Foreground="#8E8E96" FontSize="11"/><TextBlock Text="{Binding LowTxt}" FontWeight="SemiBold" FontSize="15"/></StackPanel>
                              <StackPanel Grid.Column="3" Margin="0,0,8,0"><TextBlock Text="Quedas" Foreground="#8E8E96" FontSize="11"/><TextBlock Text="{Binding QuedasTxt}" FontWeight="SemiBold" FontSize="15"/>
                                <TextBlock Text="{Binding CausaTxt}" Foreground="#FFB020" FontSize="10" TextTrimming="CharacterEllipsis" ToolTip="{Binding CausaTxt}"/></StackPanel>
                              <StackPanel Grid.Column="4"><TextBlock Text="GPU máx." Foreground="#8E8E96" FontSize="11"/><TextBlock Text="{Binding TempTxt}" FontWeight="SemiBold" FontSize="15"/></StackPanel>
                              <StackPanel Grid.Column="5"><TextBlock Text="Ping médio" Foreground="#8E8E96" FontSize="11"/><TextBlock Text="{Binding PingTxt}" FontWeight="SemiBold" FontSize="15"/></StackPanel>
                            </Grid>
                          </Border>
                        </DataTemplate>
                      </ItemsControl.ItemTemplate>
                    </ItemsControl>
                  </StackPanel>
                </Border>

                <!-- Quedas registradas -->
                <Border Style="{StaticResource Card}">
                  <StackPanel>
                    <DockPanel>
                      <Button x:Name="btnLimparQuedas" DockPanel.Dock="Right" Style="{StaticResource Link}" Content="Limpar" VerticalAlignment="Center"/>
                      <TextBlock Text="Quedas de FPS detectadas" Style="{StaticResource TituloCard}"/>
                    </DockPanel>
                    <TextBlock x:Name="txtSemQuedas" Style="{StaticResource Suave}" Margin="0,8,0,0" TextWrapping="Wrap"
                               Text="Nenhuma queda registrada. Quando o FPS cair mais de 25% em relação ao normal do jogo, a causa provável aparece aqui."/>
                    <ItemsControl x:Name="icQuedas" Margin="0,6,0,0">
                      <ItemsControl.ItemTemplate>
                        <DataTemplate>
                          <Border Background="#141416" CornerRadius="12" Padding="14,10" Margin="0,4">
                            <DockPanel>
                              <Button DockPanel.Dock="Right" Content="Ver solução" Tag="{Binding Guia}" Style="{StaticResource BtnSecundario}" Padding="12,5" FontSize="12" VerticalAlignment="Center" Margin="12,0,0,0"/>
                              <TextBlock DockPanel.Dock="Left" Text="{Binding Hora}" FontFamily="Consolas" Foreground="#8E8E96" VerticalAlignment="Center" Margin="0,0,14,0"/>
                              <StackPanel>
                                <TextBlock FontWeight="SemiBold">
                                  <Run Text="{Binding Causa}" Foreground="#FFB020"/>
                                  <Run Text="{Binding Queda}" Foreground="#8E8E96"/>
                                </TextBlock>
                                <TextBlock Text="{Binding Detalhe}" Style="{StaticResource Suave}" TextWrapping="Wrap"/>
                              </StackPanel>
                            </DockPanel>
                          </Border>
                        </DataTemplate>
                      </ItemsControl.ItemTemplate>
                    </ItemsControl>
                  </StackPanel>
                </Border>

                <TextBlock Text="Resolver problemas comuns" FontSize="18" FontWeight="SemiBold" Margin="10,14,10,2"/>
                <TextBlock Style="{StaticResource Suave}" Margin="10,0,10,4" TextWrapping="Wrap"
                           Text="Cada problema explica por que acontece e traz comandos prontos. Os comandos abrem numa janela própria e sempre pedem confirmação."/>
                <StackPanel x:Name="spGuias"/>
              </StackPanel>
            </ScrollViewer>
          </TabItem>
        </TabControl>
      </Grid>
    </Grid>
  </Border>
</Window>
'@

# Aplica a cor de destaque (inclusive nas versões com transparência, ex.: #55A6F23A)
foreach ($k in $Paleta.Keys) {
    $XamlTexto = [regex]::Replace($XamlTexto, "(?i)(#(?:[0-9a-f]{2})?)$k(?![0-9a-f])", "`${1}$($Paleta[$k])")
}
# Visual limpo: sem o brilho neon em volta de anéis, linhas e chip
if (-not $Config.Brilho) {
    $XamlTexto = [regex]::Replace($XamlTexto, '(?s)<(\w+)\.Effect>\s*<DropShadowEffect[^>]*/>\s*</\1\.Effect>', '')
}
[xml]$Xaml = $XamlTexto

$window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $Xaml))
if (Test-Path $IconePath) {
    try { $window.Icon = [Windows.Media.Imaging.BitmapFrame]::Create((New-Object Uri $IconePath), 'None', 'OnLoad') } catch {}
}
$ui = @{}
$Xaml.SelectNodes('//*[@*[local-name()="Name"]]') | ForEach-Object {
    $nome = $_.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
    if ($nome) { $ui[$nome] = $window.FindName($nome) }
}

$conversorCor = New-Object Windows.Media.BrushConverter
function Get-Pincel([string]$hex) { $conversorCor.ConvertFromString($hex) }

$pulsar = [Windows.Media.Animation.Storyboard]$window.FindResource('Pulsar')
$sbPonto = $pulsar.Clone(); $sbPonto.Begin($ui.pontoVivo, $true)
$sbChip = $pulsar.Clone(); $sbChip.Begin($ui.brilhoChip, $true)
#endregion

#region Leitura de hardware
$cntCpu = $null; $cntDisco = $null
try {
    $cntCpu   = New-Object Diagnostics.PerformanceCounter('Processor', '% Processor Time', '_Total')
    $cntDisco = New-Object Diagnostics.PerformanceCounter('PhysicalDisk', '% Idle Time', '_Total')
    [void]$cntCpu.NextValue(); [void]$cntDisco.NextValue()
} catch {}

$infoPc     = New-Object Microsoft.VisualBasic.Devices.ComputerInfo
$Nucleos    = [Environment]::ProcessorCount
$nomeCpu    = (Get-CimInstance Win32_Processor | Select-Object -First 1).Name.Trim()
$HoraBoot   = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
$RamTotalMB = $infoPc.TotalPhysicalMemory / 1MB
$NomeUsuario = (Get-Culture).TextInfo.ToTitleCase($env:USERNAME.ToLower())
$ui.txtAvatar.Text = $NomeUsuario.Substring(0, 1)
$ui.txtCpuThreads.Text = "$Nucleos threads`n$($nomeCpu -replace ' \d+-Core Processor', '')"

$Hist = @{}
foreach ($k in 'Cpu', 'Ram', 'Dsc', 'Rx', 'Tx', 'Gpu', 'GpuT') { $Hist[$k] = New-Object System.Collections.Generic.List[double] }

# Placa de vídeo: lida em segundo plano (nvidia-smi leva ~100 ms, não pode travar a tela)
$GpuSync = [hashtable]::Synchronized(@{ Rodando = $true; Dados = $null })
$gpuPs = [powershell]::Create()
[void]$gpuPs.AddScript({
    param($s)
    function ConvertTo-Numero($x) {
        $d = 0.0
        if ([double]::TryParse("$x".Trim(), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { $d } else { $null }
    }
    $smi = (Get-Command nvidia-smi -ErrorAction SilentlyContinue).Source
    if (-not $smi -and (Test-Path "$env:ProgramFiles\NVIDIA Corporation\NVSMI\nvidia-smi.exe")) { $smi = "$env:ProgramFiles\NVIDIA Corporation\NVSMI\nvidia-smi.exe" }
    $nome = (Get-CimInstance Win32_VideoController | Where-Object { $_.Name -notmatch 'Basic|Virtual|Remote' } | Select-Object -First 1).Name
    while ($s.Rodando) {
        try {
            if ($smi) {
                $l = (& $smi --query-gpu=name,utilization.gpu,temperature.gpu,memory.used,memory.total,power.draw,fan.speed,clocks.gr --format=csv,noheader,nounits | Select-Object -First 1) -split ','
                $s.Dados = @{
                    Nome = $l[0].Trim(); Uso = ConvertTo-Numero $l[1]; Temp = ConvertTo-Numero $l[2]
                    VramUsada = ConvertTo-Numero $l[3]; VramTotal = ConvertTo-Numero $l[4]; Watts = ConvertTo-Numero $l[5]
                    Fan = ConvertTo-Numero $l[6]; Clock = ConvertTo-Numero $l[7]; Fonte = 'Sensores NVIDIA'
                }
                Start-Sleep -Milliseconds 1000
            } else {
                # Outras placas (AMD/Intel): contadores do Windows, sem temperatura
                $e = Get-CimInstance Win32_PerfFormattedData_GPUPerformanceCounters_GPUEngine -ErrorAction Stop
                $uso = ($e | Where-Object { $_.Name -like '*engtype_3D' } | Measure-Object UtilizationPercentage -Sum).Sum
                $m = (Get-CimInstance Win32_PerfFormattedData_GPUPerformanceCounters_GPUAdapterMemory | Measure-Object DedicatedUsage -Maximum).Maximum
                $s.Dados = @{
                    Nome = $nome; Uso = [Math]::Min(100, [double]$uso); Temp = $null; VramUsada = $m / 1MB; VramTotal = $null
                    Watts = $null; Fan = $null; Clock = $null; Fonte = 'Contadores do Windows'
                }
                Start-Sleep -Milliseconds 1500
            }
        } catch { Start-Sleep -Seconds 3 }
    }
}).AddArgument($GpuSync)
$gpuHandle = $gpuPs.BeginInvoke()

$script:procAnterior = @{}
$script:procSemAcesso = New-Object System.Collections.Generic.HashSet[int]
$script:procHora = [DateTime]::UtcNow
$script:Processos = @()
$script:Tick = 0
$IconeCache = @{}
$script:Caminhos = @{}
$script:CaminhosHora = [DateTime]::MinValue

# Caminho do .exe de todos os processos numa consulta só (bem mais rápido que um por um)
function Update-Caminhos {
    foreach ($w in Get-CimInstance Win32_Process -Property Name, ExecutablePath -ErrorAction SilentlyContinue) {
        if (-not $w.ExecutablePath) { continue }
        $n = [IO.Path]::GetFileNameWithoutExtension($w.Name)
        if (-not $script:Caminhos.ContainsKey($n)) { $script:Caminhos[$n] = $w.ExecutablePath }
    }
    $script:CaminhosHora = Get-Date
}

# Ícone do executável (guardado em cache por nome de processo)
function Get-IconeProcesso($nome) {
    if ($IconeCache.ContainsKey($nome)) { return $IconeCache[$nome] }
    if (-not $script:Caminhos.ContainsKey($nome)) {
        if (((Get-Date) - $script:CaminhosHora).TotalSeconds -lt 30) { return $null }
        Update-Caminhos
    }
    $img = $null
    try {
        $caminho = $script:Caminhos[$nome]
        if ($caminho) {
            $ico = [System.Drawing.Icon]::ExtractAssociatedIcon($caminho)
            $img = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon($ico.Handle, [Windows.Int32Rect]::Empty,
                [Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions())
            $img.Freeze()
            $ico.Dispose()
        }
    } catch {}
    $IconeCache[$nome] = $img
    $img
}

function Update-Processos {
    $agora = [DateTime]::UtcNow
    $decorrido = ($agora - $script:procHora).TotalMilliseconds
    $script:procHora = $agora
    $novo = @{}
    $grupos = @{}
    foreach ($p in [Diagnostics.Process]::GetProcesses()) {
        if ($p.Id -eq 0) { $p.Dispose(); continue }
        $cpu = 0.0
        if (-not $script:procSemAcesso.Contains($p.Id)) {
            try {
                $t = $p.TotalProcessorTime.TotalMilliseconds
                $novo[$p.Id] = $t
                if ($script:procAnterior.ContainsKey($p.Id) -and $decorrido -gt 0) {
                    $cpu = ($t - $script:procAnterior[$p.Id]) / $decorrido / $Nucleos * 100
                }
            } catch { [void]$script:procSemAcesso.Add($p.Id) }
        }
        $g = $grupos[$p.ProcessName]
        if (-not $g) {
            $g = [pscustomobject]@{ Nome = $p.ProcessName; Qtd = 0; Cpu = 0.0; RamMB = 0.0; RamPct = 0.0; Icone = $null }
            $grupos[$p.ProcessName] = $g
        }
        $g.Qtd += 1
        $g.Cpu += $cpu
        $g.RamMB += $p.WorkingSet64 / 1MB
        $p.Dispose()
    }
    $script:procAnterior = $novo
    # Total igual ao do Gerenciador de Tarefas (cada cópia conta); os grupos por nome são os "programas"
    $script:TotalProcessos = ($grupos.Values | Measure-Object Qtd -Sum).Sum
    $script:Processos = @($grupos.Values | Sort-Object RamMB -Descending | ForEach-Object {
        $_.Cpu = [Math]::Round([Math]::Max($_.Cpu, 0), 1)
        $_.RamMB = [Math]::Round($_.RamMB, 0)
        $_.RamPct = $_.RamMB / $RamTotalMB * 100
        $_.Icone = Get-IconeProcesso $_.Nome
        $_
    })
}

function Get-BytesRede {
    $rx = 0; $tx = 0
    foreach ($n in [Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
        if ($n.OperationalStatus -ne 'Up' -or $n.NetworkInterfaceType -in 'Loopback', 'Tunnel') { continue }
        try { $s = $n.GetIPv4Statistics(); $rx += $s.BytesReceived; $tx += $s.BytesSent } catch {}
    }
    , @($rx, $tx)
}
$script:redeAnterior = Get-BytesRede
$script:redeHora = [DateTime]::UtcNow

function Format-Tamanho($bytes) {
    if ($bytes -ge 1TB) { '{0:N2} TB' -f ($bytes / 1TB) } else { '{0:N0} GB' -f ($bytes / 1GB) }
}

function Format-Velocidade($bps) {
    if ($bps -ge 1MB) { '{0:N1} MB/s' -f ($bps / 1MB) } else { '{0:N0} KB/s' -f ($bps / 1KB) }
}

function Get-CorEstado($valor, $limite, $base) {
    if ($valor -ge $limite) { $Cor.Vermelho } elseif ($valor -ge $limite - 15) { $Cor.Ambar } else { $base }
}

function Set-StatusKpi($txt, $valor, $limite) {
    if ($valor -ge $limite) { $txt.Text = '● Crítico'; $txt.Foreground = Get-Pincel $Cor.Vermelho }
    elseif ($valor -ge $limite - 15) { $txt.Text = '● Elevado'; $txt.Foreground = Get-Pincel $Cor.Ambar }
    else { $txt.Text = '● Normal'; $txt.Foreground = Get-Pincel $Cor.Lima }
}

function Add-Historico($chave, $valor) {
    $l = $Hist[$chave]
    $l.Add($valor)
    if ($l.Count -gt $HistMax) { $l.RemoveAt(0) }
}

# Desenha os últimos $janela pontos como linha + área preenchida
function Update-Serie($canvas, $linha, $area, $lista, [double]$max, [int]$janela = $HistMax) {
    $w = $canvas.ActualWidth; $h = $canvas.ActualHeight
    if ($w -le 0 -or $h -le 0 -or $lista.Count -eq 0) { return }
    $inicio = [Math]::Max(0, $lista.Count - $janela)
    $n = $lista.Count - $inicio
    $passo = $w / ($janela - 1)
    $desl = $janela - $n
    $pl = New-Object Windows.Media.PointCollection
    $pa = New-Object Windows.Media.PointCollection
    for ($i = 0; $i -lt $n; $i++) {
        $x = ($desl + $i) * $passo
        $y = $h - [Math]::Min(1.0, $lista[$inicio + $i] / $max) * ($h - 8) - 4
        $pt = New-Object Windows.Point($x, $y)
        $pl.Add($pt); $pa.Add($pt)
    }
    $pa.Add((New-Object Windows.Point($w, $h)))
    $pa.Add((New-Object Windows.Point(($desl * $passo), $h)))
    $linha.Points = $pl
    if ($area) { $area.Points = $pa }
    $pl
}

# Arco circular começando no topo, no sentido horário
function Set-Arco($path, [double]$pct, [double]$c, [double]$r) {
    $pct = [Math]::Max(0.3, [Math]::Min(99.99, $pct))
    $ini = -90 * [Math]::PI / 180
    $varredura = 360 * $pct / 100
    $fim = (-90 + $varredura) * [Math]::PI / 180
    $fig = New-Object Windows.Media.PathFigure
    $fig.StartPoint = New-Object Windows.Point(($c + $r * [Math]::Cos($ini)), ($c + $r * [Math]::Sin($ini)))
    $p1 = New-Object Windows.Point(($c + $r * [Math]::Cos($fim)), ($c + $r * [Math]::Sin($fim)))
    $fig.Segments.Add((New-Object Windows.Media.ArcSegment($p1, (New-Object Windows.Size($r, $r)), 0, ($varredura -gt 180), 'Clockwise', $true)))
    $geo = New-Object Windows.Media.PathGeometry
    $geo.Figures.Add($fig)
    $path.Data = $geo
}

# Anéis animados: valor mostrado desliza até o alvo
$Aneis = @{
    Cpu = @{ Arco = $ui.anCpu; R = 78; Base = $Cor.Lima }
    Ram = @{ Arco = $ui.anRam; R = 60; Base = $Cor.LimaMedia }
    Dsc = @{ Arco = $ui.anDsc; R = 42; Base = '#' + $Paleta['5DB82A'] }
    Esp = @{ Arco = $ui.anEsp; R = 55; Base = $Cor.Lima }
    Gpu = @{ Arco = $ui.anGpu; R = 70; Base = $Cor.Lima }
}
foreach ($a in $Aneis.Values) { $a.Alvo = 0.0; $a.Atual = 0.0; $a.CorAtual = ''; Set-Arco $a.Arco 0 90 $a.R }

function Set-CorAnel($a, $hex) {
    if ($a.CorAtual -eq $hex) { return }
    $a.CorAtual = $hex
    $a.Arco.Stroke = Get-Pincel $hex
    if ($a.Arco.Effect) { $a.Arco.Effect.Color = [Windows.Media.ColorConverter]::ConvertFromString($hex) }
}

$timerAnim = New-Object Windows.Threading.DispatcherTimer
$timerAnim.Interval = [TimeSpan]::FromMilliseconds(33)
$timerAnim.add_Tick({
    foreach ($a in $Aneis.Values) {
        $dif = $a.Alvo - $a.Atual
        if ([Math]::Abs($dif) -lt 0.05) { continue }
        $a.Atual += $dif * 0.15
        Set-Arco $a.Arco $a.Atual 90 $a.R
    }
})

# Barras do histórico de CPU (12 barras de 10 segundos)
$NumBarras = 12
$Barras = @()
$rotBarra = $null
for ($i = 0; $i -lt $NumBarras; $i++) {
    $col = New-Object Windows.Controls.StackPanel
    $col.VerticalAlignment = 'Bottom'
    $col.Margin = '5,0'
    if ($i -eq $NumBarras - 1) {
        $selo = New-Object Windows.Controls.Border
        $selo.Background = Get-Pincel '#F2F2F5'; $selo.CornerRadius = 7; $selo.Padding = '6,2'; $selo.Margin = '-10,0,-10,6'
        $selo.HorizontalAlignment = 'Center'
        $rotBarra = New-Object Windows.Controls.TextBlock
        $rotBarra.FontSize = 11; $rotBarra.FontWeight = 'SemiBold'; $rotBarra.Foreground = Get-Pincel '#0B0B0D'
        $selo.Child = $rotBarra
        [void]$col.Children.Add($selo)
    }
    $b = New-Object Windows.Controls.Border
    $b.CornerRadius = 7
    $b.Background = $window.FindResource('BarraLima')
    $b.Height = 6
    if ($i -lt $NumBarras - 1) { $b.Opacity = 0.5 } elseif ($Config.Brilho) {
        $fx = New-Object Windows.Media.Effects.DropShadowEffect
        $fx.Color = [Windows.Media.ColorConverter]::ConvertFromString($Cor.Lima); $fx.BlurRadius = 16; $fx.ShadowDepth = 0; $fx.Opacity = 0.7
        $b.Effect = $fx
    }
    [void]$col.Children.Add($b)
    [void]$ui.ugBarras.Children.Add($col)
    $Barras += $b
    $r = New-Object Windows.Controls.TextBlock
    $r.Text = if ($i -eq $NumBarras - 1) { '0s' } elseif ($i % 3 -eq 0) { '-{0}s' -f (($NumBarras - $i) * 10) } else { '' }
    $r.HorizontalAlignment = 'Center'; $r.FontSize = 11; $r.Foreground = Get-Pincel $Cor.Suave
    [void]$ui.ugRotulos.Children.Add($r)
}

# Pontos do indicador de saúde (barra lateral)
$PontosSaude = @()
for ($i = 0; $i -lt 10; $i++) {
    $e = New-Object Windows.Shapes.Ellipse
    $e.Width = 9; $e.Height = 9; $e.Fill = Get-Pincel '#2A2A2F'
    [void]$ui.ugPontos.Children.Add($e)
    $PontosSaude += $e
}
#endregion

#region Alertas
$script:Acima = @{ Cpu = 0; Ram = 0; Dsc = 0; Esp = 0; GpuT = 0 }
$script:Em100 = @{ Dsc = 0; Cpu = 0; Gpu = 0 }
$script:DicaGuia = $null
$script:UltimoAlerta = @{}
$script:NaoLidos = 0

$tray = New-Object System.Windows.Forms.NotifyIcon
try { $tray.Icon = New-Object System.Drawing.Icon($IconePath, [System.Windows.Forms.SystemInformation]::SmallIconSize) }
catch { $tray.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Process -Id $PID).Path) }
$tray.Text = $AppNome
$tray.Visible = $true
$menuTray = New-Object System.Windows.Forms.ContextMenuStrip
[void]$menuTray.Items.Add('Abrir Monitor PC', $null, { Show-Janela })
[void]$menuTray.Items.Add('Sair', $null, { Exit-App })
$tray.ContextMenuStrip = $menuTray
$tray.add_MouseDoubleClick({ Show-Janela })
$tray.add_BalloonTipClicked({ Show-Janela })

function Update-Badge {
    if ($script:NaoLidos -gt 0) {
        $ui.txtBadge.Text = [string][Math]::Min(99, $script:NaoLidos)
        $ui.bdBadge.Visibility = 'Visible'
    } else {
        $ui.bdBadge.Visibility = 'Collapsed'
    }
}

function Send-Alerta($titulo, $mensagem) {
    $tray.ShowBalloonTip(8000, $titulo, $mensagem, [System.Windows.Forms.ToolTipIcon]::Warning)
    if ($Config.Som) { [System.Media.SystemSounds]::Exclamation.Play() }
    $linha = '{0:dd/MM HH:mm:ss}  ·  {1} — {2}' -f (Get-Date), $titulo, $mensagem
    $ui.lbHist.Items.Insert(0, $linha)
    if ($ui.lbHist.Items.Count -gt 200) { $ui.lbHist.Items.RemoveAt(200) }
    try { Add-Content -Path $LogPath -Value $linha -Encoding UTF8 } catch {}
    if ($ui.tabs.SelectedIndex -ne 2) { $script:NaoLidos++; Update-Badge }
}

function Test-Limite($chave, $valor, $limite, $titulo, $mensagem) {
    if ($valor -lt $limite) { $script:Acima[$chave] = 0; return }
    $script:Acima[$chave] += $Intervalo
    if ($script:Acima[$chave] -lt $Config.Sustentar) { return }
    $ultimo = $script:UltimoAlerta[$chave]
    if ($ultimo -and ((Get-Date) - $ultimo).TotalMinutes -lt $Config.IntervaloAlertaMin) { return }
    $script:UltimoAlerta[$chave] = Get-Date
    Send-Alerta $titulo $mensagem
}
#endregion

#region Atualização periódica
function Update-Tela {
    try {
        $script:Tick++
        if ($script:Tick % 2 -eq 1) { Update-Processos }
        $topCpu = @($script:Processos | Sort-Object Cpu -Descending | Select-Object -First 1)
        $topRam = @($script:Processos | Select-Object -First 5)
        $agora = Get-Date
        $ui.txtRelogio.Text = $agora.ToString('HH:mm')
        $saudacao = if ($agora.Hour -lt 12) { 'Bom dia' } elseif ($agora.Hour -lt 18) { 'Boa tarde' } else { 'Boa noite' }
        $ui.txtSaudacao.Text = "$saudacao, $NomeUsuario!"

        # CPU
        $cpu = if ($cntCpu) { [Math]::Min(100, [Math]::Round($cntCpu.NextValue())) } else { 0 }
        Add-Historico 'Cpu' $cpu
        $Aneis.Cpu.Alvo = $cpu
        Set-CorAnel $Aneis.Cpu (Get-CorEstado $cpu $Config.CpuLimite $Aneis.Cpu.Base)
        $ui.txtKCpu.Text = "$cpu%"; $ui.txtRCpu.Text = "$cpu%"; $ui.txtChip.Text = "$cpu%"
        Set-StatusKpi $ui.txtKCpuSt $cpu $Config.CpuLimite

        # RAM
        $ramTotal = [double]$infoPc.TotalPhysicalMemory
        $ramLivre = [double]$infoPc.AvailablePhysicalMemory
        $ramUso = [Math]::Round((1 - $ramLivre / $ramTotal) * 100)
        Add-Historico 'Ram' $ramUso
        $Aneis.Ram.Alvo = $ramUso
        Set-CorAnel $Aneis.Ram (Get-CorEstado $ramUso $Config.RamLimite $Aneis.Ram.Base)
        $ui.txtKRam.Text = "$ramUso%"; $ui.txtRRam.Text = "$ramUso%"
        Set-StatusKpi $ui.txtKRamSt $ramUso $Config.RamLimite
        $ui.txtRamGB.Text = '{0:N1} GB de {1:N1} GB  ·  {2}%' -f (($ramTotal - $ramLivre) / 1GB), ($ramTotal / 1GB), $ramUso

        # Disco (atividade)
        $dsc = if ($cntDisco) { [Math]::Max(0, [Math]::Min(100, [Math]::Round(100 - $cntDisco.NextValue()))) } else { 0 }
        Add-Historico 'Dsc' $dsc
        $Aneis.Dsc.Alvo = $dsc
        Set-CorAnel $Aneis.Dsc (Get-CorEstado $dsc $Config.DiscoLimite $Aneis.Dsc.Base)
        $ui.txtKDsc.Text = "$dsc%"; $ui.txtRDsc.Text = "$dsc%"
        Set-StatusKpi $ui.txtKDscSt $dsc $Config.DiscoLimite

        # Armazenamento: a rosca mostra a unidade mais cheia
        $piorLivre = 101; $piorDrive = ''; $piorInfo = $null
        foreach ($d in [IO.DriveInfo]::GetDrives() | Where-Object { $_.DriveType -eq 'Fixed' -and $_.IsReady }) {
            $livrePct = [Math]::Round($d.AvailableFreeSpace / $d.TotalSize * 100)
            if ($livrePct -lt $piorLivre) { $piorLivre = $livrePct; $piorDrive = $d.Name.TrimEnd('\'); $piorInfo = $d }
        }
        if ($piorInfo) {
            $espUso = 100 - $piorLivre
            $Aneis.Esp.Alvo = $espUso
            Set-CorAnel $Aneis.Esp (Get-CorEstado $espUso (100 - $Config.EspacoMinimo) $Cor.Lima)
            $ui.txtEspDrive.Text = $piorDrive
            $ui.txtEspLivre.Text = Format-Tamanho $piorInfo.AvailableFreeSpace
            $ui.txtEspTotal.Text = Format-Tamanho $piorInfo.TotalSize
            $ui.txtEspUsadoPct.Text = "$espUso%"
            $ui.txtEspLivrePct.Text = "$piorLivre%"
        }

        # Rede
        $redeAtual = Get-BytesRede
        $agoraUtc = [DateTime]::UtcNow
        $dt = [Math]::Max(0.1, ($agoraUtc - $script:redeHora).TotalSeconds)
        $rx = [Math]::Max(0, ($redeAtual[0] - $script:redeAnterior[0]) / $dt)
        $tx = [Math]::Max(0, ($redeAtual[1] - $script:redeAnterior[1]) / $dt)
        $script:redeAnterior = $redeAtual; $script:redeHora = $agoraUtc
        Add-Historico 'Rx' $rx; Add-Historico 'Tx' $tx
        $ui.txtRx.Text = Format-Velocidade $rx
        $ui.txtTx.Text = Format-Velocidade $tx
        $ultimos = @($Hist.Rx | Select-Object -Last 60) + @($Hist.Tx | Select-Object -Last 60)
        $escala = [Math]::Max(100KB, ($ultimos | Measure-Object -Maximum).Maximum * 1.2)
        $ui.txtRedeEscala.Text = "Escala: 0 – $(Format-Velocidade $escala)"
        [void](Update-Serie $ui.cvRede $ui.lnTx $null $Hist.Tx $escala 60)
        $pts = Update-Serie $ui.cvRede $ui.lnRx $ui.agRx $Hist.Rx $escala 60
        # Pontos sobre a linha de download (a cada 10 s + o último em destaque)
        $ui.cvPontos.Children.Clear()
        if ($pts) {
            for ($i = $pts.Count - 1; $i -ge 0; $i -= 10) {
                $ultimo = $i -eq $pts.Count - 1
                $tam = if ($ultimo) { 12 } else { 8 }
                $e = New-Object Windows.Shapes.Ellipse
                $e.Width = $tam; $e.Height = $tam
                $e.Fill = Get-Pincel $(if ($ultimo) { $Cor.Lima } else { '#0F0F11' })
                $e.Stroke = Get-Pincel $Cor.Lima; $e.StrokeThickness = 2
                [Windows.Controls.Canvas]::SetLeft($e, $pts[$i].X - $tam / 2)
                [Windows.Controls.Canvas]::SetTop($e, $pts[$i].Y - $tam / 2)
                [void]$ui.cvPontos.Children.Add($e)
            }
        }

        # Memória (gráfico de área)
        [void](Update-Serie $ui.cvRam $ui.lnRam $ui.agRam $Hist.Ram 100)

        # Barras do histórico de CPU
        $hc = $Hist.Cpu
        $altMax = [Math]::Max(20, $ui.ugBarras.ActualHeight - 30)
        for ($i = 0; $i -lt $NumBarras; $i++) {
            $fimFatia = $hc.Count - ($NumBarras - 1 - $i) * 10
            $iniFatia = [Math]::Max(0, $fimFatia - 10)
            $media = 0
            if ($fimFatia -gt 0) {
                $soma = 0.0
                for ($j = $iniFatia; $j -lt $fimFatia; $j++) { $soma += $hc[$j] }
                $media = $soma / ($fimFatia - $iniFatia)
            }
            $Barras[$i].Height = [Math]::Max(6, $media / 100 * $altMax)
            if ($i -eq $NumBarras - 1) { $rotBarra.Text = '{0:N0}%' -f $media }
        }
        $ui.txtCpuMedia.Text = '{0:N0}' -f ($hc | Measure-Object -Average).Average
        $ui.txtCpuPico.Text = '{0:N0}' -f ($hc | Measure-Object -Maximum).Maximum

        # Placa de vídeo
        $g = $GpuSync.Dados
        $gpuTemp = $null
        if ($g) {
            $gu = [Math]::Round([double]$g.Uso)
            Add-Historico 'Gpu' $gu
            $Aneis.Gpu.Alvo = $gu
            $ui.txtGpuUso.Text = "$gu%"
            $ui.txtGpuNome.Text = $g.Nome
            $ui.txtGpuFonte.Text = "·  $($g.Fonte)"
            if ($null -ne $g.Temp) {
                $gpuTemp = [Math]::Round($g.Temp)
                Add-Historico 'GpuT' $gpuTemp
                $corT = Get-CorEstado $gpuTemp $Config.GpuTempLimite $Cor.Lima
                $ui.txtGpuTemp.Text = "$gpuTemp °C"
                $ui.txtGpuTemp.Foreground = Get-Pincel $(if ($corT -eq $Cor.Lima) { '#F2F2F5' } else { $corT })
                $ui.icGpuTemp.Foreground = Get-Pincel $corT
                Set-CorAnel $Aneis.Gpu $corT
            } else {
                $ui.txtGpuTemp.Text = 'N/D'
            }
            if ($g.VramTotal) {
                $ui.txtGpuVram.Text = '{0:N1} / {1:N0} GB' -f ($g.VramUsada / 1024), ($g.VramTotal / 1024)
                $ui.pbVram.Value = $g.VramUsada / $g.VramTotal * 100
            } else {
                $ui.txtGpuVram.Text = '{0:N1} GB' -f ($g.VramUsada / 1024)
            }
            $ui.txtGpuWatts.Text = if ($null -ne $g.Watts) { '{0:N0} W' -f $g.Watts } else { 'N/D' }
            $ui.txtGpuClock.Text = if ($null -ne $g.Clock) { '{0:N0} MHz' -f $g.Clock } else { 'N/D' }
            $ui.txtGpuFan.Text = if ($null -ne $g.Fan) { "Clock · fan $([int]$g.Fan)%" } else { 'Clock' }
            [void](Update-Serie $ui.cvGpu $ui.lnGpu $ui.agGpu $Hist.Gpu 100)
            if ($Hist.GpuT.Count) { [void](Update-Serie $ui.cvGpu $ui.lnGpuT $null $Hist.GpuT 100) }
        }

        # Saúde geral (0-100)
        $up = $agora - $HoraBoot
        $pontos = 100
        if ($gpuTemp -gt 80) { $pontos -= ($gpuTemp - 80) * 1.5 }
        $pontos -= [Math]::Max(0, $cpu - 70) * 1.0
        $pontos -= [Math]::Max(0, $ramUso - 70) * 1.2
        $pontos -= [Math]::Max(0, $dsc - 80) * 0.8
        if ($piorInfo -and $piorLivre -lt 20) { $pontos -= (20 - $piorLivre) * 1.5 }
        if ($up.TotalDays -gt 7) { $pontos -= 8 }
        $saude = [int][Math]::Max(0, [Math]::Min(100, [Math]::Round($pontos)))
        $ui.txtSaudeNum.Text = $saude
        $ui.txtSaudeRot.Text = if ($saude -ge 85) { 'Excelente' } elseif ($saude -ge 70) { 'Bom' } elseif ($saude -ge 50) { 'Atenção' } else { 'Crítico' }
        for ($i = 0; $i -lt 10; $i++) {
            $PontosSaude[$i].Fill = Get-Pincel $(if ($i -lt [Math]::Round($saude / 10)) { $Cor.Lima } else { '#2A2A2F' })
        }
        $ui.txtUptime.Text = 'Ligado há {0} d {1} h {2} min' -f $up.Days, $up.Hours, $up.Minutes
        $ui.txtResumo.Text = if ($saude -ge 85) { 'Seu PC está rodando liso. Continue de olho nos limites de alerta.' }
            elseif ($saude -ge 70) { 'Tudo sob controle, mas vale dar uma olhada nas dicas ao lado.' }
            else { 'Seu PC está sob pressão agora. Veja o que mais está consumindo recursos.' }

        # Maiores consumidores
        if ($script:Tick % 2 -eq 1) {
            $ui.icTop.ItemsSource = @($topRam | ForEach-Object {
                [pscustomobject]@{
                    Nome = $_.Nome; Icone = $_.Icone; Texto = '{0:N0} MB' -f $_.RamMB
                    Sub = '{0} {1} · {2:N1}% CPU' -f $_.Qtd, $(if ($_.Qtd -eq 1) { 'instância' } else { 'instâncias' }), $_.Cpu
                } })
            if ($ui.tabs.SelectedIndex -eq 1) { Update-GradeProcessos }
        }

        # Dica do momento
        $r1 = $topRam | Select-Object -First 1
        $c1 = $topCpu | Select-Object -First 1
        # Quanto tempo seguido cada peça está em 100% (para não reagir a picos)
        $script:Em100.Dsc = if ($dsc -ge 95) { $script:Em100.Dsc + 1 } else { 0 }
        $script:Em100.Cpu = if ($cpu -ge 95) { $script:Em100.Cpu + 1 } else { 0 }
        $script:Em100.Gpu = if ($null -ne $gpuTemp -and $gpuTemp -ge $Config.GpuTempLimite) { $script:Em100.Gpu + 1 } else { 0 }
        $script:DicaGuia = $null
        if ($script:Em100.Dsc -ge 10) {
            $ui.txtDicaTitulo.Text = 'Disco em 100%'
            $ui.txtDica.Text = "Há $($script:Em100.Dsc) s seguidos. Isso causa travadas e lentidão — veja o que costuma causar e os comandos para resolver."
            $script:DicaGuia = 'disco'
        } elseif ($script:Em100.Cpu -ge 10) {
            $ui.txtDicaTitulo.Text = 'Processador em 100%'
            $ui.txtDica.Text = "$(if ($c1) { "$($c1.Nome) é o que mais usa agora ($($c1.Cpu)%). " })Veja como descobrir e resolver."
            $script:DicaGuia = 'cpu'
        } elseif ($script:Em100.Gpu -ge 10) {
            $ui.txtDicaTitulo.Text = "Placa de vídeo a $gpuTemp °C"
            $ui.txtDica.Text = 'Ela pode reduzir a velocidade sozinha para não superaquecer. Veja o diagnóstico e o que fazer.'
            $script:DicaGuia = 'gpu'
        } elseif ($up.TotalDays -gt 7) {
            $ui.txtDicaTitulo.Text = "Ligado há $($up.Days) dias"
            $ui.txtDica.Text = 'Reinicie o PC (Iniciar > Reiniciar) para limpar a memória. "Desligar" com Inicialização Rápida não faz isso.'
        } elseif ($piorInfo -and $piorLivre -lt 20) {
            $ui.txtDicaTitulo.Text = "Pouco espaço em $piorDrive"
            $ui.txtDica.Text = 'Limpe os temporários e o cache do Windows Update na aba Otimizar.'
        } elseif ($ramUso -ge 75 -and $r1) {
            $ui.txtDicaTitulo.Text = "$($r1.Nome) usa $('{0:N1}' -f ($r1.RamMB / 1024)) GB"
            $ui.txtDica.Text = 'Feche programas que não está usando para liberar memória RAM.'
        } elseif ($cpu -ge 70 -and $c1) {
            $ui.txtDicaTitulo.Text = "$($c1.Nome) puxa a CPU"
            $ui.txtDica.Text = "Está usando $($c1.Cpu)% do processador agora."
        } else {
            $ui.txtDicaTitulo.Text = 'Acelere o boot'
            $ui.txtDica.Text = 'Desative programas que abrem sozinhos com o Windows — é a otimização que mais faz diferença.'
        }
        $ui.btnDica.Content = if ($script:DicaGuia) { 'Ver solução  →' } else { 'Ir para Otimizar  →' }

        # Alertas
        if ($Config.AlertasAtivos) {
            Test-Limite 'Cpu' $cpu $Config.CpuLimite 'Processador alto' ("CPU em $cpu% (limite $($Config.CpuLimite)%). Maior uso: $($c1.Nome) ($($c1.Cpu)%).")
            Test-Limite 'Ram' $ramUso $Config.RamLimite 'Memória RAM alta' ("RAM em $ramUso% (limite $($Config.RamLimite)%). Maior uso: $($r1.Nome) ($('{0:N0}' -f $r1.RamMB) MB).")
            Test-Limite 'Dsc' $dsc $Config.DiscoLimite 'Disco sobrecarregado' ("Disco ocupado $dsc% do tempo (limite $($Config.DiscoLimite)%).")
            if ($piorInfo) {
                Test-Limite 'Esp' (100 - $piorLivre) (100 - $Config.EspacoMinimo) 'Pouco espaço em disco' ("A unidade $piorDrive tem só $piorLivre% livre. Use a aba Otimizar para limpar.")
            }
            if ($null -ne $gpuTemp) {
                Test-Limite 'GpuT' $gpuTemp $Config.GpuTempLimite 'Placa de vídeo quente' ("GPU a $gpuTemp °C (limite $($Config.GpuTempLimite) °C), com $([int]$g.Uso)% de uso. Verifique a ventilação do gabinete.")
            }
        }

        Update-Jogos $cpu $ramUso $dsc $g

        $ui.txtStatus.Text = "Atualizado às $($agora.ToString('HH:mm:ss')) · $($script:TotalProcessos) processos · $($script:Processos.Count) programas"
    } catch {
        $ui.txtStatus.Text = "Erro na leitura: $($_.Exception.Message)"
    }
}

function Update-GradeProcessos {
    $lista = $script:Processos
    $filtro = $ui.txtBusca.Text.Trim()
    if ($filtro) { $lista = @($lista | Where-Object { $_.Nome -like "*$filtro*" }) }
    $lista = if ($ui.rbCpu.IsChecked) { @($lista | Sort-Object Cpu -Descending) }
             elseif ($ui.rbNome.IsChecked) { @($lista | Sort-Object Nome) }
             else { @($lista | Sort-Object RamMB -Descending) }
    $ui.txtProcInfo.Text = if ($filtro) { "$($lista.Count) programa(s) para '$filtro' · $(($lista | Measure-Object Qtd -Sum).Sum) processos" }
                           else { "$($script:TotalProcessos) processos em execução, agrupados em $($lista.Count) programas" }
    $nomeSel = if ($ui.dgProc.SelectedItem) { $ui.dgProc.SelectedItem.Nome }
    $ui.dgProc.ItemsSource = $lista
    if ($nomeSel) {
        $item = $lista | Where-Object { $_.Nome -eq $nomeSel } | Select-Object -First 1
        if ($item) { $ui.dgProc.SelectedItem = $item }
    }
}
#endregion

#region Vida útil das peças
$ScriptAnalise = {
    param($AnoCompra, $CorOk)
    $itens = New-Object System.Collections.Generic.List[object]

    # $Estimado = porcentagem calculada pela idade (sem sensor), mostrada em outra cor
    function Add-Item($Componente, $Glifo, $Modelo, $Saude, $Detalhes, $Obs, [switch]$Estimado) {
        $cor = '#5E5E66'; $txt = 'Sem dados'; $valor = 0
        if ($null -ne $Saude) {
            $valor = [Math]::Max(0, [Math]::Min(100, [Math]::Round($Saude)))
            if ($Estimado) {
                $txt = "~$valor% est."
                $cor = '#5ED6B5'
            } else {
                $txt = "$valor%"
                $cor = if ($valor -ge 70) { $CorOk } elseif ($valor -ge 40) { '#FFB020' } else { '#FF5A4D' }
            }
        }
        $itens.Add([pscustomobject]@{
            Componente = $Componente; Glifo = [string][char]$Glifo; Modelo = $Modelo; Saude = [double]$valor
            SaudeTexto = $txt; Cor = $cor; Detalhes = $Detalhes; Observacao = $Obs
        })
    }

    # Idade do PC: ano informado pelo usuário ou data do BIOS
    $idade = $null; $fonteIdade = ''
    if ($AnoCompra -match '^\d{4}$') {
        $idade = [Math]::Max(0, (Get-Date).Year - [int]$AnoCompra + 0.5)
        $fonteIdade = "comprado em $AnoCompra"
    } else {
        $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
        if ($bios.ReleaseDate) {
            $idade = ((Get-Date) - $bios.ReleaseDate).TotalDays / 365.25
            $fonteIdade = 'pela data do BIOS'
        }
    }
    function Get-Estimativa($vidaAnos) { if ($null -eq $idade) { $null } else { 100 - ($idade / $vidaAnos * 100) } }
    $txtIdade = if ($null -ne $idade) { 'Idade ~{0:N1} anos ({1})' -f $idade, $fonteIdade } else { 'Idade desconhecida' }

    # Processador
    foreach ($c in Get-CimInstance Win32_Processor) {
        Add-Item 'Processador' 0xE950 $c.Name.Trim() (Get-Estimativa 15) `
            ("{0} núcleos / {1} threads • {2:N1} GHz • {3}" -f $c.NumberOfCores, $c.NumberOfLogicalProcessors, ($c.MaxClockSpeed / 1000), $txtIdade) `
            'Vida típica 10-15+ anos. O que mata CPU é calor: limpe a poeira e troque a pasta térmica a cada 2-3 anos.' -Estimado
    }

    # Memória RAM
    $mods = @(Get-CimInstance Win32_PhysicalMemory)
    if ($mods.Count) {
        $total = ($mods | Measure-Object Capacity -Sum).Sum
        $desc = ($mods | ForEach-Object { '{0:N0} GB {1} MHz {2}' -f ($_.Capacity / 1GB), $_.ConfiguredClockSpeed, "$($_.Manufacturer)".Trim() }) -join ' | '
        Add-Item 'Memória RAM' 0xE964 ('{0:N0} GB em {1} pente(s)' -f ($total / 1GB), $mods.Count) (Get-Estimativa 15) $desc `
            'RAM raramente falha. Se tiver telas azuis ou travamentos aleatórios, rode o Diagnóstico de Memória (aba Otimizar).' -Estimado
    }

    # Discos (SMART)
    foreach ($d in Get-PhysicalDisk) {
        $rel = $null
        try { $rel = $d | Get-StorageReliabilityCounter -ErrorAction Stop } catch {}
        $tipo = "$($d.MediaType)"; $barramento = "$($d.BusType)"; $status = "$($d.HealthStatus)"
        $ehSsd = $tipo -eq 'SSD' -or $barramento -eq 'NVMe'
        $det = @("$tipo $barramento • $('{0:N0}' -f ($d.Size / 1GB)) GB • Status do Windows: $status")
        if ($rel) {
            if ($rel.Temperature) { $det += "Temperatura: $($rel.Temperature) °C" }
            if ($rel.PowerOnHours) { $det += 'Tempo ligado: {0:N0} h' -f $rel.PowerOnHours }
            if ($rel.ReadErrorsTotal) { $det += "Erros de leitura: $($rel.ReadErrorsTotal)" }
        }
        $saude = $null; $obs = ''
        if ($ehSsd) {
            if ($rel -and $null -ne $rel.Wear) {
                $saude = 100 - $rel.Wear
                $obs = "Desgaste informado pelo próprio SSD: $($rel.Wear)%. SSDs costumam durar 5-10 anos de uso normal."
            } else {
                $obs = 'Abra como administrador para ler o desgaste real (SMART) do SSD.'
            }
        } else {
            if ($rel -and $rel.PowerOnHours) {
                $saude = 100 - ($rel.PowerOnHours / 40000 * 100)
                $obs = 'HDs costumam durar 30-50 mil horas ligados. Barulho de "clique" é sinal de falha: faça backup.'
            } else {
                $obs = 'Abra como administrador para ler as horas de uso (SMART) do HD.'
            }
        }
        if ($status -ne 'Healthy') {
            $saude = if ($null -ne $saude) { [Math]::Min($saude, 20) } else { 10 }
            $obs = 'ATENÇÃO: o Windows detectou problema neste disco. Faça backup agora!'
        }
        Add-Item ($(if ($ehSsd) { 'SSD' } else { 'HD' })) 0xEDA2 $d.FriendlyName $saude ($det -join ' • ') $obs
    }

    # Placa de vídeo
    foreach ($g in Get-CimInstance Win32_VideoController) {
        $obs = 'Vida típica 6-10 anos (menos se joga muito com temperatura alta). Mantenha os coolers limpos.'
        $dataDrv = ''
        if ($g.DriverDate) {
            $dataDrv = ' ({0:dd/MM/yyyy})' -f $g.DriverDate
            if (((Get-Date) - $g.DriverDate).TotalDays -gt 365) { $obs = 'Driver com mais de 1 ano — atualize pelo site do fabricante. ' + $obs }
        }
        Add-Item 'Placa de vídeo' 0xE7F4 $g.Name (Get-Estimativa 9) ("Driver $($g.DriverVersion)$dataDrv • $txtIdade") $obs -Estimado
    }

    # Placa-mãe
    $mb = Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue
    $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
    if ($mb) {
        $dataBios = if ($bios.ReleaseDate) { ' de {0:dd/MM/yyyy}' -f $bios.ReleaseDate } else { '' }
        Add-Item 'Placa-mãe' 0xEEA1 ("$($mb.Manufacturer) $($mb.Product)".Trim()) (Get-Estimativa 10) `
            ("BIOS $($bios.SMBIOSBIOSVersion)$dataBios • $txtIdade") `
            'Vida típica 8-12 anos. Capacitores estufados ou PC que não liga de primeira são sinais de desgaste.' -Estimado
    }

    # Fonte (não tem sensor)
    Add-Item 'Fonte' 0xE945 'Não é detectável por software' (Get-Estimativa 10) $txtIdade `
        'Vida típica 5-10 anos (fontes boas, 80 Plus, duram mais). Desligamentos sozinhos sob carga = suspeite da fonte.' -Estimado

    # Bateria (notebooks)
    $bat = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
    if ($bat) {
        $projetada = $null; $atual = $null; $ciclos = $null
        try { $projetada = (Get-CimInstance -Namespace root\wmi BatteryStaticData -ErrorAction Stop | Select-Object -First 1).DesignedCapacity } catch {}
        try { $atual = (Get-CimInstance -Namespace root\wmi BatteryFullChargedCapacity -ErrorAction Stop | Select-Object -First 1).FullChargedCapacity } catch {}
        try { $ciclos = (Get-CimInstance -Namespace root\wmi BatteryCycleCount -ErrorAction Stop | Select-Object -First 1).CycleCount } catch {}
        $saude = if ($projetada -and $atual) { $atual / $projetada * 100 } else { $null }
        $det = "Carga atual: $($bat.EstimatedChargeRemaining)%"
        if ($projetada -and $atual) { $det += " • Capacidade: $atual de $projetada mWh" }
        if ($ciclos) { $det += " • Ciclos: $ciclos" }
        Add-Item 'Bateria' 0xE945 $bat.Name $saude $det 'Abaixo de 60% de saúde vale trocar. Evite deixar sempre em 100% na tomada.'
    }

    # Windows
    $so = Get-CimInstance Win32_OperatingSystem
    $up = (Get-Date) - $so.LastBootUpTime
    $obsSo = 'Tudo certo.'
    if ($up.TotalDays -gt 7) { $obsSo = 'Ligado há mais de 7 dias: REINICIE (Iniciar > Reiniciar). "Desligar" com Inicialização Rápida não limpa a memória.' }
    Add-Item 'Windows' 0xE9D9 $so.Caption $null ('Instalado em {0:dd/MM/yyyy} • Último boot há {1} d {2} h' -f $so.InstallDate, $up.Days, $up.Hours) $obsSo

    $itens
}

$script:analise = $null
$timerAnalise = New-Object Windows.Threading.DispatcherTimer
$timerAnalise.Interval = [TimeSpan]::FromMilliseconds(300)
$timerAnalise.add_Tick({
    if (-not $script:analise -or -not $script:analise.Handle.IsCompleted) { return }
    $timerAnalise.Stop()
    try {
        $res = $script:analise.Ps.EndInvoke($script:analise.Handle)
        $ui.icVida.ItemsSource = @($res)
        $msg = 'Análise concluída às {0:HH:mm}.' -f (Get-Date)
        if (-not $EhAdmin) { $msg += ' Abra como admin para ler o desgaste real do SSD/HD.' }
        $ui.txtVidaStatus.Text = $msg
    } catch {
        $ui.txtVidaStatus.Text = "Erro na análise: $($_.Exception.Message)"
    } finally {
        $script:analise.Ps.Dispose()
        $script:analise = $null
        $ui.btnAnalisar.IsEnabled = $true
    }
})

function Start-Analise {
    if ($script:analise) { return }
    $ui.btnAnalisar.IsEnabled = $false
    $ui.txtVidaStatus.Text = 'Analisando as peças...'
    $ps = [powershell]::Create()
    [void]$ps.AddScript($ScriptAnalise).AddArgument($Config.AnoCompra).AddArgument($Cor.Lima)
    $script:analise = @{ Ps = $ps; Handle = $ps.BeginInvoke() }
    $timerAnalise.Start()
}
#endregion

#region Teste de velocidade do disco
$TestesDiscoPath = Join-Path $ConfigDir 'testes_disco.jsonl'
$script:DiscoEscolhido = $env:SystemDrive.Substring(0, 1)
$script:discoJob = $null

$ScriptDisco = {
    param($letra)
    $r = [ordered]@{ Letra = $letra; Nome = ''; Bus = ''; Midia = ''; Leitura = $null; Gravacao = $null; Aleatoria = $null; Latencia = $null; Erro = ''; Data = (Get-Date).ToString('s') }
    try {
        $disco = Get-Partition -DriveLetter $letra -ErrorAction Stop | Get-Disk
        $r.Nome = $disco.FriendlyName; $r.Bus = "$($disco.BusType)"
        $r.Midia = "$((Get-PhysicalDisk | Where-Object { $_.DeviceId -eq "$($disco.Number)" } | Select-Object -First 1).MediaType)"
    } catch {}
    $saida = @(& "$env:windir\System32\winsat.exe" disk -drive $letra 2>&1 | ForEach-Object { "$_" })
    function Get-Numero($linha) {
        if ("$linha" -match '(\d+[\.,]\d+)\s*(MB/s|ms)') { [double]::Parse($matches[1].Replace(',', '.'), [Globalization.CultureInfo]::InvariantCulture) }
    }
    $r.Leitura = Get-Numero ($saida | Where-Object { $_ -match 'Sequential 64\.0 Read' } | Select-Object -First 1)
    $r.Gravacao = Get-Numero ($saida | Where-Object { $_ -match 'Sequential 64\.0 Write' } | Select-Object -First 1)
    $r.Aleatoria = Get-Numero ($saida | Where-Object { $_ -match 'Random 16\.0 Read' } | Select-Object -First 1)
    $r.Latencia = Get-Numero ($saida | Where-Object { $_ -match '95' -and $_ -match 'ms' } | Select-Object -First 1)
    if ($null -eq $r.Leitura) { $r.Erro = ($saida | Select-Object -Last 3) -join ' ' }
    [pscustomobject]$r
}

function Get-VereditoDisco($r) {
    $min = $null
    if ($r.Bus -eq 'NVMe') { $tipo = 'SSD NVMe'; $min = 1500; $faixa = '1.500 a 7.000 MB/s' }
    elseif ($r.Midia -eq 'HDD') { $tipo = 'HD'; $min = 80; $faixa = '80 a 250 MB/s' }
    elseif ($r.Midia -eq 'SSD') { $tipo = 'SSD SATA'; $min = 350; $faixa = '450 a 550 MB/s' }
    elseif ($r.Bus -eq 'USB') { $tipo = 'disco USB' }
    else { $tipo = 'disco' }
    if ($null -eq $min) { return @("Velocidade medida no $tipo $($r.Nome).", '#F2F2F5') }
    if ($r.Leitura -ge $min) { return @("Rendendo bem para um $tipo (o normal é $faixa).", $Cor.Lima) }
    if ($r.Leitura -ge $min * 0.5) {
        return @("Abaixo do normal para um $tipo (o normal é $faixa). Feche programas que usam o disco e teste de novo; se continuar baixo, confira a saúde dele aqui embaixo e o espaço livre.", $Cor.Ambar)
    }
    @("Bem abaixo do normal para um $tipo (o normal é $faixa). Pode ser disco quase cheio, superaquecido ou desgastado — faça backup e verifique a saúde dele.", $Cor.Vermelho)
}

function Update-ListaDiscos {
    $ui.wpDiscos.Children.Clear()
    foreach ($d in [IO.DriveInfo]::GetDrives() | Where-Object { $_.DriveType -eq 'Fixed' -and $_.IsReady }) {
        $letra = $d.Name.Substring(0, 1)
        $rb = New-Object Windows.Controls.RadioButton
        $rb.Style = $window.FindResource('Seg')
        $rb.GroupName = 'disco'; $rb.Content = "${letra}:"; $rb.Tag = $letra; $rb.Margin = '0,0,4,8'
        $rb.IsChecked = $letra -eq $script:DiscoEscolhido
        $rb.add_Checked({ param($s, $e) $script:DiscoEscolhido = $s.Tag })
        [void]$ui.wpDiscos.Children.Add($rb)
    }
}

function Format-Velocidade2($mb) { if ($null -eq $mb) { '--' } elseif ($mb -ge 1000) { '{0:N2} GB/s' -f ($mb / 1000) } else { '{0:N0} MB/s' -f $mb } }

function Show-TesteDisco($r) {
    # Resultado anterior do mesmo disco, para comparar
    $ant = $null
    if (Test-Path $TestesDiscoPath) {
        $ant = Get-Content $TestesDiscoPath -Encoding UTF8 -ErrorAction SilentlyContinue | ForEach-Object { try { $_ | ConvertFrom-Json } catch {} } |
            Where-Object { $_.Letra -eq $r.Letra -and $_.Data -ne $r.Data -and $_.Leitura } | Select-Object -Last 1
    }
    $ui.ugDisco.Visibility = 'Visible'; $ui.txtDiscoVeredito.Visibility = 'Visible'
    $ui.txtDiscoLeitura.Text = Format-Velocidade2 $r.Leitura
    $ui.txtDiscoGravacao.Text = Format-Velocidade2 $r.Gravacao
    $ui.txtDiscoAleatoria.Text = Format-Velocidade2 $r.Aleatoria
    $ui.txtDiscoLatencia.Text = if ($null -ne $r.Latencia) { '{0:N2} ms' -f $r.Latencia } else { '--' }
    $quando = if ($ant) { ([datetime]$ant.Data).ToString('dd/MM/yyyy') } else { '' }
    $ui.txtDiscoLeituraAnt.Text = if ($ant) { "antes ($quando): $(Format-Velocidade2 $ant.Leitura)" } else { '' }
    $ui.txtDiscoGravacaoAnt.Text = if ($ant) { "antes ($quando): $(Format-Velocidade2 $ant.Gravacao)" } else { '' }
    $v = Get-VereditoDisco $r
    $ui.txtDiscoVeredito.Text = $v[0]; $ui.txtDiscoVeredito.Foreground = Get-Pincel $v[1]
    $ui.txtDiscoInfo.Text = "$($r.Letra): $($r.Nome) · testado em $(([datetime]$r.Data).ToString('dd/MM HH:mm'))"
}

$timerDisco = New-Object Windows.Threading.DispatcherTimer
$timerDisco.Interval = [TimeSpan]::FromMilliseconds(400)
$timerDisco.add_Tick({
    if (-not $script:discoJob) { $timerDisco.Stop(); return }
    $j = $script:discoJob
    if ($j.Arquivo) {
        # Teste rodando numa janela elevada à parte: espera o arquivo com o resultado
        if (-not (Test-Path $j.Arquivo)) {
            if (((Get-Date) - $j.Inicio).TotalSeconds -gt 150) {
                $timerDisco.Stop(); $script:discoJob = $null; $ui.btnTestarDisco.IsEnabled = $true
                $ui.txtDiscoInfo.Text = 'O teste não terminou a tempo. Tente de novo.'
            }
            return
        }
        Start-Sleep -Milliseconds 300   # garante que terminou de gravar
    } elseif (-not $j.Handle.IsCompleted) { return }
    $timerDisco.Stop()
    try {
        $r = if ($j.Arquivo) { Get-Content $j.Arquivo -Raw -Encoding UTF8 | ConvertFrom-Json } else { $j.Ps.EndInvoke($j.Handle) | Select-Object -Last 1 }
        if ($null -eq $r.Leitura) {
            $ui.txtDiscoInfo.Text = "Não foi possível medir: $($r.Erro)"
        } else {
            try { Add-Content -Path $TestesDiscoPath -Value ($r | ConvertTo-Json -Compress) -Encoding UTF8 } catch {}
            Show-TesteDisco $r
        }
    } catch {
        $ui.txtDiscoInfo.Text = "Erro no teste: $($_.Exception.Message)"
    } finally {
        if ($j.Ps) { $j.Ps.Dispose() }
        if ($j.Arquivo) { Remove-Item $j.Arquivo, $j.Script -ErrorAction SilentlyContinue }
        $script:discoJob = $null
        $ui.btnTestarDisco.IsEnabled = $true
    }
})

$ui.btnTestarDisco.add_Click({
    if ($script:discoJob) { return }
    $letra = $script:DiscoEscolhido
    if ($EhAdmin) {
        $ps = [powershell]::Create()
        [void]$ps.AddScript($ScriptDisco).AddArgument($letra)
        $script:discoJob = @{ Ps = $ps; Handle = $ps.BeginInvoke() }
    } else {
        # O winsat exige administrador: roda só o teste numa janela elevada escondida e lê o resultado de um arquivo
        $base = Join-Path $env:TEMP ('monitorpc_disco_' + [guid]::NewGuid().ToString('N'))
        $script = "$base.ps1"; $arquivo = "$base.json"
        $conteudo = "param(`$Letra, `$Saida)`n`$r = & {`n$($ScriptDisco.ToString())`n} `$Letra`n`$r | ConvertTo-Json | Set-Content -Path `$Saida -Encoding UTF8"
        [IO.File]::WriteAllText($script, $conteudo, (New-Object Text.UTF8Encoding $true))
        try {
            Start-Process powershell.exe -Verb RunAs -WindowStyle Hidden -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -Letra $letra -Saida `"$arquivo`"" -ErrorAction Stop
        } catch {
            Remove-Item $script -ErrorAction SilentlyContinue
            $ui.txtDiscoInfo.Text = 'O teste precisa de permissão de administrador (o Windows exige isso para medir o disco).'
            return
        }
        $script:discoJob = @{ Arquivo = $arquivo; Script = $script; Inicio = Get-Date }
    }
    $ui.btnTestarDisco.IsEnabled = $false
    $ui.txtDiscoInfo.Text = "Testando ${letra}: ... evite usar o PC pesado por alguns segundos."
    $timerDisco.Start()
})

Update-ListaDiscos
# Mostra o último teste já feito no disco do sistema
if (Test-Path $TestesDiscoPath) {
    $ultimo = Get-Content $TestesDiscoPath -Encoding UTF8 -ErrorAction SilentlyContinue | ForEach-Object { try { $_ | ConvertFrom-Json } catch {} } |
        Where-Object { $_.Leitura } | Select-Object -Last 1
    if ($ultimo) { Show-TesteDisco $ultimo }
}
#endregion

#region Segurança
$ScriptSeguranca = {
    param($CacheAssin, $CorOk)
    $win = $env:windir
    # Processos do Windows e a pasta (dentro de C:\Windows) de onde devem rodar
    $sistema = @{
        svchost = 'System32'; csrss = 'System32'; lsass = 'System32'; services = 'System32'; winlogon = 'System32'
        wininit = 'System32'; smss = 'System32'; spoolsv = 'System32'; taskhostw = 'System32'; dwm = 'System32'
        sihost = 'System32'; ctfmon = 'System32'; conhost = 'System32'; RuntimeBroker = 'System32'
        fontdrvhost = 'System32'; dllhost = 'System32|SysWOW64'; rundll32 = 'System32|SysWOW64'; explorer = ''
    }
    $nomesLegitimos = 'taskhost', 'lsaiso', 'conhost', 'svchost'
    $cores = @{ 'Crítico' = '#FF5A4D'; 'Atenção' = '#FFB020'; 'Info' = '#5ED6B5'; 'OK' = $CorOk }
    $ordem = @{ 'Crítico' = 0; 'Atenção' = 1; 'Info' = 2 }

    $protecao = New-Object System.Collections.Generic.List[object]
    $achados = [ordered]@{}

    function Add-Protecao($Titulo, $Glifo, $Nivel, $Status, $Detalhe, $Acao = '', $AcaoTexto = '') {
        $protecao.Add([pscustomobject]@{
            Titulo = $Titulo; Glifo = [string][char]$Glifo; Nivel = $Nivel; Cor = $cores[$Nivel]; Status = $Status
            Detalhe = $Detalhe; Acao = $Acao; AcaoTexto = $AcaoTexto; AcaoVisivel = $(if ($AcaoTexto) { 'Visible' } else { 'Collapsed' })
        })
    }
    # Achados iguais (mesmo arquivo, mesmo problema) viram um só, juntando os PIDs
    function Add-Achado($Nivel, $Titulo, $Nome, $ProcId, $Caminho, $Motivo) {
        $chave = "$Titulo|$Caminho"
        if ($achados.Contains($chave)) { $achados[$chave].Pids += $ProcId; return }
        $achados[$chave] = [pscustomobject]@{
            Severidade = $Nivel; Cor = $cores[$Nivel]; CorFundo = '#22' + $cores[$Nivel].TrimStart('#'); Titulo = $Titulo
            Nome = $Nome; Pids = @($ProcId); Linha = ''; Caminho = $Caminho; Motivo = $Motivo
        }
    }
    function Get-Distancia([string]$a, [string]$b) {
        if ([Math]::Abs($a.Length - $b.Length) -gt 1) { return 9 }
        # Duas letras vizinhas trocadas (scvhost) conta como 1 erro
        if ($a.Length -eq $b.Length) {
            $dif = @(0..($a.Length - 1) | Where-Object { $a[$_] -cne $b[$_] })
            if ($dif.Count -eq 2 -and $dif[1] -eq $dif[0] + 1 -and $a[$dif[0]] -ceq $b[$dif[1]] -and $a[$dif[1]] -ceq $b[$dif[0]]) { return 1 }
        }
        $ant = 0..$b.Length
        for ($i = 1; $i -le $a.Length; $i++) {
            $atual = @($i) + @(0) * $b.Length
            for ($j = 1; $j -le $b.Length; $j++) {
                $custo = if ($a[$i - 1] -ceq $b[$j - 1]) { 0 } else { 1 }
                $atual[$j] = [Math]::Min([Math]::Min($ant[$j] + 1, $atual[$j - 1] + 1), $ant[$j - 1] + $custo)
            }
            $ant = $atual
        }
        $ant[$b.Length]
    }
    # productState do Central de Segurança: byte do meio 0x10/0x11 = ligado; último byte 0x00 = atualizado
    function Test-Ligado($estado) { (($estado -shr 8) -band 0xFF) -in 0x10, 0x11 }

    # ---- Antivírus ----
    $avs = @(Get-CimInstance -Namespace root\SecurityCenter2 AntiVirusProduct -ErrorAction SilentlyContinue)
    # Registro "órfão": o programa foi desinstalado mas o Windows ainda tem ele cadastrado
    function Test-Existe($av) {
        $exe = [Environment]::ExpandEnvironmentVariables("$($av.pathToSignedProductExe)")
        -not $exe -or (Test-Path -LiteralPath $exe)
    }
    $orfaos = @($avs | Where-Object { -not (Test-Existe $_) })
    $ativos = @($avs | Where-Object { (Test-Ligado $_.productState) -and (Test-Existe $_) })
    $velhos = @($ativos | Where-Object { ($_.productState -band 0xFF) -ne 0 })
    if ($orfaos.Count -and $ativos.Count -le 1) {
        $nomes = ($orfaos | ForEach-Object { $_.displayName }) -join ', '
        Add-Protecao 'Antivírus' 0xEA18 'Atenção' 'Restos de antivírus antigo' `
            "$nomes já foi desinstalado, mas continua cadastrado no Windows. Ativo de verdade: $(if ($ativos.Count) { $ativos[0].displayName } else { 'nenhum' })." 'apps' 'Limpar restos'
    } elseif (-not $ativos.Count) {
        Add-Protecao 'Antivírus' 0xEA18 'Crítico' 'Nenhum ativo' 'Nenhum antivírus está protegendo o PC agora.' 'defender' 'Abrir Segurança do Windows'
    } elseif ($ativos.Count -gt 1) {
        Add-Protecao 'Antivírus' 0xEA18 'Atenção' "$($ativos.Count) ativos ao mesmo tempo" `
            ((($ativos | ForEach-Object { $_.displayName }) -join ' + ') + '. Dois antivírus juntos brigam entre si e deixam o PC lento — mantenha só um.') 'apps' 'Abrir Aplicativos'
    } elseif ($velhos.Count) {
        Add-Protecao 'Antivírus' 0xEA18 'Atenção' 'Desatualizado' "$($velhos[0].displayName) está com as definições de vírus atrasadas." 'defender' 'Abrir Segurança do Windows'
    } else {
        Add-Protecao 'Antivírus' 0xEA18 'OK' $ativos[0].displayName 'Ativo e atualizado.'
    }

    # ---- Windows Defender ----
    $mp = $null
    try { $mp = Get-MpComputerStatus -ErrorAction Stop } catch {}
    if ($mp -and $mp.AntivirusEnabled) {
        $diasScan = if ($mp.QuickScanEndTime) { ((Get-Date) - $mp.QuickScanEndTime).TotalDays } else { 999 }
        if (-not $mp.RealTimeProtectionEnabled) {
            Add-Protecao 'Windows Defender' 0xE9D9 'Crítico' 'Tempo real desligado' 'A proteção em tempo real está desativada.' 'defender' 'Abrir Segurança do Windows'
        } elseif ($mp.AntivirusSignatureAge -gt 7) {
            Add-Protecao 'Windows Defender' 0xE9D9 'Atenção' "Definições com $($mp.AntivirusSignatureAge) dias" 'A lista de vírus conhecidos está desatualizada.' 'atualizar' 'Atualizar definições'
        } elseif ($diasScan -gt 14) {
            Add-Protecao 'Windows Defender' 0xE9D9 'Atenção' 'Sem verificação recente' "Última verificação rápida há $([int]$diasScan) dias." 'scan' 'Verificação rápida'
        } else {
            Add-Protecao 'Windows Defender' 0xE9D9 'OK' 'Protegendo em tempo real' ('Última verificação {0:dd/MM HH:mm} · definições de {1:dd/MM}' -f $mp.QuickScanEndTime, $mp.AntivirusSignatureLastUpdated) 'scan' 'Verificação rápida'
        }
    } else {
        Add-Protecao 'Windows Defender' 0xE9D9 'Info' 'Em espera' 'Outro antivírus está no comando da proteção.'
    }

    # ---- Firewall ----
    $traducao = @{ Domain = 'Domínio'; Private = 'Privada'; Public = 'Pública' }
    $desligados = @(Get-NetFirewallProfile -ErrorAction SilentlyContinue | Where-Object { "$($_.Enabled)" -ne 'True' })
    $fwOutros = @(Get-CimInstance -Namespace root\SecurityCenter2 FirewallProduct -ErrorAction SilentlyContinue | Where-Object { Test-Ligado $_.productState })
    if (-not $desligados.Count) {
        Add-Protecao 'Firewall' 0xE7EF 'OK' 'Ativado' 'Firewall do Windows ligado em todas as redes.'
    } elseif ($fwOutros.Count) {
        Add-Protecao 'Firewall' 0xE7EF 'OK' "Por $($fwOutros[0].displayName)" 'O Firewall do Windows está desligado, mas outro firewall está ativo.'
    } else {
        $redes = ($desligados | ForEach-Object { $traducao["$($_.Name)"] }) -join ', '
        Add-Protecao 'Firewall' 0xE7EF 'Crítico' 'Desligado' "Desligado nas redes: $redes. O PC fica exposto a conexões de fora." 'firewall' 'Ativar firewall'
    }

    # ---- Controle de Conta de Usuário ----
    $uac = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -ErrorAction SilentlyContinue).EnableLUA
    if ($uac -eq 1) {
        Add-Protecao 'Controle de Conta (UAC)' 0xE72E 'OK' 'Ativado' 'O Windows pede permissão antes de mudanças no sistema.'
    } else {
        Add-Protecao 'Controle de Conta (UAC)' 0xE72E 'Crítico' 'Desativado' 'Programas podem alterar o sistema sem pedir permissão.' 'uac' 'Abrir configuração'
    }

    # ---- Processos ----
    $verificados = 0; $semCaminho = 0
    $pastasSuspeitas = @(
        @{ Caminho = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\'); Nome = 'Temp' }
        @{ Caminho = "$win\Temp"; Nome = 'Temp do Windows' }
        @{ Caminho = "$env:USERPROFILE\Downloads"; Nome = 'Downloads' }
        @{ Caminho = $env:PUBLIC; Nome = 'Usuários\Público' }
        @{ Caminho = "$env:SystemDrive\`$Recycle.Bin"; Nome = 'Lixeira' }
    )
    foreach ($p in Get-CimInstance Win32_Process -Property Name, ProcessId, ExecutablePath -ErrorAction SilentlyContinue) {
        if ($p.ProcessId -le 4) { continue }
        $nome = [IO.Path]::GetFileNameWithoutExtension($p.Name)
        $cam = $p.ExecutablePath
        if (-not $cam) { if ($sistema.ContainsKey($nome)) { $semCaminho++ }; continue }
        if ($cam.StartsWith('\\?\')) { $cam = $cam.Substring(4) }   # prefixo de caminho longo do Windows
        $verificados++
        $pasta = [IO.Path]::GetDirectoryName($cam)

        # 1) Processo do Windows rodando fora da pasta certa
        if ($sistema.ContainsKey($nome)) {
            $esperadas = if ($sistema[$nome]) { @($sistema[$nome] -split '\|' | ForEach-Object { Join-Path $win $_ }) } else { @($win) }
            if ($esperadas -notcontains $pasta) {
                Add-Achado 'Crítico' "$($p.Name) fora da pasta do Windows" $p.Name $p.ProcessId $cam `
                    "O $($p.Name) verdadeiro fica em $($esperadas[0]). Um arquivo com esse nome em outro lugar é um disfarce comum de vírus."
            }
            continue
        }
        # 2) Nome que imita um processo do Windows (ex.: svch0st, scvhost)
        if ($nome.Length -ge 5 -and $nomesLegitimos -notcontains $nome) {
            foreach ($s in $sistema.Keys) {
                if ((Get-Distancia $nome.ToLower() $s.ToLower()) -eq 1) {
                    Add-Achado 'Crítico' "Nome imita o $s.exe do Windows" $p.Name $p.ProcessId $cam `
                        "'$($p.Name)' é quase igual a '$s.exe'. Vírus costumam usar nomes assim para passar despercebidos."
                    break
                }
            }
        }
        # 3) Assinatura digital (fica em cache: cada arquivo só é conferido uma vez)
        if ($cam -like '*\WindowsApps\*') {
            $assin = @{ Status = 'Valid'; Assinante = 'Microsoft Store' }
        } else {
            if (-not $CacheAssin.ContainsKey($cam)) {
                try {
                    $sig = Get-AuthenticodeSignature -FilePath $cam -ErrorAction Stop
                    $quem = if ($sig.SignerCertificate) { $sig.SignerCertificate.GetNameInfo('SimpleName', $false) } else { '' }
                    $CacheAssin[$cam] = @{ Status = "$($sig.Status)"; Assinante = $quem }
                } catch { $CacheAssin[$cam] = @{ Status = 'Erro'; Assinante = '' } }
            }
            $assin = $CacheAssin[$cam]
        }
        $assinado = $assin.Status -eq 'Valid'
        # 4) Pasta suspeita
        $local = $pastasSuspeitas | Where-Object { $_.Caminho -and $cam.StartsWith($_.Caminho + '\', [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1

        if ($assin.Status -eq 'HashMismatch') {
            Add-Achado 'Crítico' 'Arquivo alterado depois de assinado' $p.Name $p.ProcessId $cam `
                'A assinatura digital não confere com o conteúdo: o arquivo foi modificado. Isso é sinal forte de infecção.'
        } elseif ($local -and -not $assinado) {
            Add-Achado 'Crítico' "Programa sem assinatura rodando de $($local.Nome)" $p.Name $p.ProcessId $cam `
                "Programas legítimos quase nunca rodam direto da pasta $($local.Nome) sem assinatura digital. Se você não abriu isso, verifique no VirusTotal."
        } elseif ($local) {
            Add-Achado 'Atenção' "Rodando da pasta $($local.Nome)" $p.Name $p.ProcessId $cam `
                "É assinado por $($assin.Assinante), mas programas costumam ser instalados antes de rodar. Se não reconhece, verifique."
        } elseif (-not $assinado) {
            if ($cam.StartsWith($win + '\', [StringComparison]::OrdinalIgnoreCase)) {
                Add-Achado 'Crítico' 'Arquivo do Windows sem assinatura' $p.Name $p.ProcessId $cam `
                    'Quase tudo dentro da pasta do Windows é assinado pela Microsoft. Um arquivo sem assinatura ali é suspeito.'
            } else {
                Add-Achado 'Info' 'Sem assinatura digital' $p.Name $p.ProcessId $cam `
                    'Não é necessariamente perigoso — muitos jogos e programas pequenos não são assinados. Se não reconhece o programa, verifique no VirusTotal.'
            }
        }
    }

    $lista = @($achados.Values | ForEach-Object {
        $_.Linha = '{0}  ·  PID {1}' -f $_.Nome, (($_.Pids | Select-Object -First 5) -join ', ') + $(if ($_.Pids.Count -gt 5) { " (+$($_.Pids.Count - 5))" } else { '' })
        $_
    } | Sort-Object { $ordem[$_.Severidade] })

    [pscustomobject]@{ Protecao = [object[]]$protecao.ToArray(); Achados = $lista; Verificados = $verificados; SemCaminho = $semCaminho; Hora = Get-Date }
}

$CacheAssinaturas = [hashtable]::Synchronized(@{})
$script:seg = $null
$script:SegCriticosAntes = -1

function Start-Seguranca {
    if ($script:seg) { return }
    $ui.btnVerificar.IsEnabled = $false
    $ui.txtSegResumo.Text = 'Verificando proteção do sistema e processos...'
    $ps = [powershell]::Create()
    [void]$ps.AddScript($ScriptSeguranca).AddArgument($CacheAssinaturas).AddArgument($Cor.Lima)
    $script:seg = @{ Ps = $ps; Handle = $ps.BeginInvoke() }
    $timerSeg.Start()
}

function Show-Seguranca($r) {
    $ui.icProtecao.ItemsSource = @($r.Protecao)
    $ui.icAchados.ItemsSource = @($r.Achados)
    $todos = @($r.Protecao) + @($r.Achados)
    $criticos = @($todos | Where-Object { $_.Nivel -eq 'Crítico' -or $_.Severidade -eq 'Crítico' })
    $atencao = @($todos | Where-Object { $_.Nivel -eq 'Atenção' -or $_.Severidade -eq 'Atenção' })
    $ui.bdSemAchados.Visibility = if (@($r.Achados).Count) { 'Collapsed' } else { 'Visible' }
    $ui.txtSemAchados.Text = "$($r.Verificados) programas conferidos."
    $ui.txtAchadosInfo.Text = '{0} verificados · {1} achado(s)' -f $r.Verificados, @($r.Achados).Count

    if ($criticos.Count) {
        $titulo = "$($criticos.Count) problema(s) crítico(s)"; $corG = $Cor.Vermelho; $glifo = 0xE7BA
    } elseif ($atencao.Count) {
        $titulo = "$($atencao.Count) ponto(s) de atenção"; $corG = $Cor.Ambar; $glifo = 0xEA18
    } else {
        $titulo = 'Tudo protegido'; $corG = $Cor.Lima; $glifo = 0xE73E
    }
    $resumo = 'Última verificação às {0:HH:mm} · {1} programas conferidos.' -f $r.Hora, $r.Verificados
    if ($r.SemCaminho -gt 0 -and -not $EhAdmin) {
        $resumo += " $($r.SemCaminho) processos do Windows só podem ser conferidos abrindo como administrador."
    }
    $ui.txtSegTitulo.Text = "Segurança — $titulo"
    $ui.txtSegResumo.Text = $resumo
    foreach ($par in @(@($ui.icSegIcone, $ui.bdSegIcone), @($ui.icSegMini, $ui.bdSegMini))) {
        $par[0].Text = [string][char]$glifo
        $par[0].Foreground = Get-Pincel $corG
        $par[1].BorderBrush = Get-Pincel $corG
    }
    $ui.txtSegMini.Text = $titulo
    $ui.txtSegMini.Foreground = Get-Pincel $corG
    $ui.txtSegMiniSub.Text = if ($criticos.Count) { ($criticos | Select-Object -First 2 | ForEach-Object { if ($_.Status) { "$($_.Titulo): $($_.Status)" } else { $_.Titulo } }) -join "`n" }
                             else { "$($r.Verificados) programas conferidos às $('{0:HH:mm}' -f $r.Hora)" }

    # Avisa só quando aparecem problemas críticos novos
    if ($Config.AlertasAtivos -and $script:SegCriticosAntes -ge 0 -and $criticos.Count -gt $script:SegCriticosAntes) {
        $c = $criticos | Select-Object -First 1
        Send-Alerta 'Problema de segurança' ("$($c.Titulo): $(if ($c.Status) { $c.Status } else { $c.Nome }). Veja a aba Segurança.")
    } elseif ($Config.AlertasAtivos -and $script:SegCriticosAntes -lt 0 -and $criticos.Count) {
        Send-Alerta 'Problema de segurança' "$titulo encontrado(s). Veja a aba Segurança."
    }
    $script:SegCriticosAntes = $criticos.Count
}

$timerSeg = New-Object Windows.Threading.DispatcherTimer
$timerSeg.Interval = [TimeSpan]::FromMilliseconds(500)
$timerSeg.add_Tick({
    if (-not $script:seg -or -not $script:seg.Handle.IsCompleted) { return }
    $timerSeg.Stop()
    try {
        $res = $script:seg.Ps.EndInvoke($script:seg.Handle) | Select-Object -Last 1
        Show-Seguranca $res
    } catch {
        $ui.txtSegResumo.Text = "Erro na verificação: $($_.Exception.Message)"
    } finally {
        $script:seg.Ps.Dispose()
        $script:seg = $null
        $ui.btnVerificar.IsEnabled = $true
    }
})

# Primeira verificação logo após abrir; depois a cada 15 minutos
$timerSegAgenda = New-Object Windows.Threading.DispatcherTimer
$timerSegAgenda.Interval = [TimeSpan]::FromSeconds(5)
$timerSegAgenda.add_Tick({
    $timerSegAgenda.Interval = [TimeSpan]::FromMinutes(15)
    Start-Seguranca
})

$AcoesSeguranca = @{
    firewall  = @{ Titulo = 'Ativar o Firewall do Windows'; Admin = $true
                   Descricao = 'Liga o firewall nas redes Domínio, Privada e Pública. Programas que já funcionam na internet continuam funcionando.'
                   Script = 'Set-NetFirewallProfile -All -Enabled True; Get-NetFirewallProfile | Format-Table Name, Enabled -AutoSize' }
    scan      = @{ Titulo = 'Verificação rápida do Windows Defender'; Admin = $true
                   Descricao = 'Procura vírus nos lugares mais comuns do sistema. Leva alguns minutos.'
                   Script = 'Write-Host "Verificando... aguarde."; Start-MpScan -ScanType QuickScan; Write-Host "Verificação concluída. Resultados em Segurança do Windows > Proteção contra vírus." -ForegroundColor Green' }
    atualizar = @{ Titulo = 'Atualizar definições do Defender'; Admin = $true
                   Descricao = 'Baixa a lista mais recente de vírus conhecidos.'
                   Script = 'Update-MpSignature; Get-MpComputerStatus | Format-List AntivirusSignatureLastUpdated, AntivirusSignatureVersion' }
}

function Invoke-AcaoSeguranca($acao) {
    switch ($acao) {
        'defender' { Start-Process 'windowsdefender:' }
        'apps'     { Select-Aba 6 }
        'uac'      { Start-Process 'UserAccountControlSettings.exe' }
        default    { if ($AcoesSeguranca.ContainsKey($acao)) { Invoke-Comando $AcoesSeguranca[$acao] } }
    }
}

function Invoke-AcaoAchado($acao, $item) {
    switch ($acao) {
        'abrir' {
            if (Test-Path $item.Caminho) { Start-Process explorer.exe "/select,`"$($item.Caminho)`"" }
        }
        'vt' {
            try {
                $hash = (Get-FileHash -Path $item.Caminho -Algorithm SHA256 -ErrorAction Stop).Hash
                Start-Process "https://www.virustotal.com/gui/file/$hash"
            } catch {
                [Windows.MessageBox]::Show("Não foi possível ler o arquivo para calcular o hash.`n$($_.Exception.Message)", $AppNome) | Out-Null
            }
        }
        'fim' {
            $resp = [Windows.MessageBox]::Show("Finalizar $($item.Nome) (PID $($item.Pids -join ', '))?`n`n$($item.Caminho)", $AppNome, 'YesNo', 'Warning')
            if ($resp -ne 'Yes') { return }
            $falhas = 0
            foreach ($id in $item.Pids) { try { Stop-Process -Id $id -Force -ErrorAction Stop } catch { $falhas++ } }
            if ($falhas) { [Windows.MessageBox]::Show("$falhas processo(s) não puderam ser finalizados. Tente abrir o Monitor como administrador.", $AppNome) | Out-Null }
            else { Start-Seguranca }
        }
    }
}

# Botões dentro das listas: um único tratador lê o Tag (ação) e o item da linha
$ui.icProtecao.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent, [Windows.RoutedEventHandler]{
    param($s, $e)
    if ($e.OriginalSource -is [Windows.Controls.Button]) { Invoke-AcaoSeguranca $e.OriginalSource.Tag }
})
$ui.icAchados.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent, [Windows.RoutedEventHandler]{
    param($s, $e)
    if ($e.OriginalSource -is [Windows.Controls.Button]) { Invoke-AcaoAchado $e.OriginalSource.Tag $e.OriginalSource.DataContext }
})
$ui.btnVerificar.add_Click({ Start-Seguranca })
#endregion

#region Aplicativos
$ScriptApps = {
    $lista = New-Object System.Collections.Generic.List[object]
    function Get-Icone($caminho) {
        try {
            $c = ([Environment]::ExpandEnvironmentVariables("$caminho") -split ',')[0].Trim().Trim('"')
            if (-not $c -or -not (Test-Path -LiteralPath $c)) { return $null }
            $ico = [System.Drawing.Icon]::ExtractAssociatedIcon($c)
            $img = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon($ico.Handle, [Windows.Int32Rect]::Empty,
                [Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions())
            $img.Freeze()
            $ico.Dispose()
            $img
        } catch { $null }
    }
    function Format-MB($mb) {
        if ($mb -ge 1024) { '{0:N1} GB' -f ($mb / 1024) } elseif ($mb -gt 0) { '{0:N0} MB' -f [Math]::Max(1, $mb) } else { '—' }
    }
    $avs = @(Get-CimInstance -Namespace root\SecurityCenter2 AntiVirusProduct -ErrorAction SilentlyContinue |
        Where-Object { $_.displayName -notmatch 'Windows Defender|Microsoft Defender' })

    # ---- Programas (registro do Windows) ----
    $chaves = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
              'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
              'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    $vistos = @{}
    foreach ($e in Get-ItemProperty $chaves -ErrorAction SilentlyContinue) {
        if (-not $e.DisplayName -or $e.SystemComponent -eq 1 -or $e.ParentKeyName -or "$($e.ReleaseType)" -match 'Update|Hotfix') { continue }
        $desinst = if ($e.UninstallString) { $e.UninstallString } else { $e.QuietUninstallString }
        if (-not $desinst) { continue }
        $id = "$($e.DisplayName)|$($e.DisplayVersion)"
        if ($vistos[$id]) { continue }
        $vistos[$id] = 1
        $data = $null
        if ("$($e.InstallDate)" -match '^\d{8}$') { try { $data = [datetime]::ParseExact("$($e.InstallDate)", 'yyyyMMdd', $null) } catch {} }
        $mb = if ($e.EstimatedSize) { [double]$e.EstimatedSize / 1024 } else { 0 }
        $ehAv = @($avs | Where-Object { $e.DisplayName -like "*$($_.displayName)*" }).Count -gt 0
        $lista.Add([pscustomobject]@{
            Nome = $e.DisplayName; Editor = "$($e.Publisher)"; Versao = "$($e.DisplayVersion)"
            Sub = (@("$($e.Publisher)", "$($e.DisplayVersion)") | Where-Object { $_ }) -join '  ·  '
            TamanhoMB = $mb; TamanhoTxt = Format-MB $mb; Data = $data; DataTxt = $(if ($data) { $data.ToString('dd/MM/yyyy') } else { '—' })
            Tipo = 'Programa'; Desinstalar = $desinst; Pacote = $null; Icone = Get-Icone $e.DisplayIcon
            Tag = $(if ($ehAv) { 'Antivírus' } else { '' }); TagVisivel = $(if ($ehAv) { 'Visible' } else { 'Collapsed' })
        })
    }

    # ---- Apps da Microsoft Store (só os que aparecem no menu Iniciar) ----
    $nomesIniciar = @{}
    foreach ($s in Get-StartApps) {
        if ($s.AppID -match '^(.+?)!' -and -not $nomesIniciar.ContainsKey($matches[1])) { $nomesIniciar[$matches[1]] = $s.Name }
    }
    foreach ($p in Get-AppxPackage -PackageTypeFilter Main -ErrorAction SilentlyContinue) {
        if ($p.IsFramework -or $p.NonRemovable -or "$($p.SignatureKind)" -eq 'System') { continue }
        $nome = $nomesIniciar[$p.PackageFamilyName]
        if (-not $nome) { continue }
        $editor = if ($p.Publisher -match 'CN=([^,]+)') { $matches[1].Trim('"') } else { '' }
        $data = $null
        try { $data = (Get-Item -LiteralPath $p.InstallLocation -ErrorAction Stop).CreationTime } catch {}
        $lista.Add([pscustomobject]@{
            Nome = $nome; Editor = $editor; Versao = "$($p.Version)"
            Sub = (@($editor, "$($p.Version)") | Where-Object { $_ }) -join '  ·  '
            TamanhoMB = 0; TamanhoTxt = '—'; Data = $data; DataTxt = $(if ($data) { $data.ToString('dd/MM/yyyy') } else { '—' })
            Tipo = 'Microsoft Store'; Desinstalar = $null; Pacote = $p.PackageFullName; Icone = $null
            Tag = ''; TagVisivel = 'Collapsed'
        })
    }

    # ---- Restos: antivírus registrados no Windows cujo programa já não existe ----
    $restos = New-Object System.Collections.Generic.List[object]
    $pastasProgramas = @(${env:ProgramFiles(x86)}, $env:ProgramFiles) | Where-Object { $_ }
    foreach ($a in $avs) {
        $exe = [Environment]::ExpandEnvironmentVariables("$($a.pathToSignedProductExe)")
        if (-not $exe -or (Test-Path -LiteralPath $exe)) { continue }
        # Pasta "raiz" do fabricante dentro de Arquivos de Programas (ex.: ...\360\Total Security\... -> 360)
        $topo = ''
        foreach ($pf in $pastasProgramas) {
            if ($exe.StartsWith($pf + '\', [StringComparison]::OrdinalIgnoreCase)) { $topo = ($exe.Substring($pf.Length + 1) -split '\\')[0] }
        }
        if ($topo -in 'Common Files', 'Microsoft', 'Windows Defender', 'WindowsApps') { $topo = '' }
        $semEspaco = $a.displayName -replace '\s', ''
        $pastas = @()
        foreach ($base in @($env:ProgramData, ${env:ProgramFiles(x86)}, $env:ProgramFiles, $env:LOCALAPPDATA, $env:APPDATA) | Where-Object { $_ }) {
            $pastas += @(Get-ChildItem -LiteralPath $base -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -ieq $semEspaco -or ($topo -and $_.Name -ieq $topo) } | ForEach-Object { $_.FullName })
        }
        $mb = 0
        foreach ($pa in $pastas) { $mb += (Get-ChildItem -LiteralPath $pa -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum / 1MB }
        $itens = @("• Registro no Central de Segurança do Windows ($($a.displayName))") + @($pastas | ForEach-Object { "• Pasta $_" })
        $restos.Add([pscustomobject]@{
            Nome = $a.displayName; Pastas = $pastas
            Titulo = "Restos de $($a.displayName)"
            Texto = "O programa já foi desinstalado, mas deixou restos. Por causa deles o Windows ainda acha que este antivírus está ativo." +
                    $(if ($pastas.Count) { " As pastas ocupam {0:N1} MB." -f $mb } else { '' })
            Itens = $itens -join "`n"
        })
    }

    [pscustomobject]@{ Apps = [object[]]$lista.ToArray(); Restos = [object[]]$restos.ToArray() }
}

$script:Apps = @()
$script:appsJob = $null

function Start-Apps {
    if ($script:appsJob) { return }
    $ui.btnAppsAtualizar.IsEnabled = $false
    $ui.txtAppsInfo.Text = 'Carregando programas instalados...'
    # STA: os ícones são imagens do WPF e precisam de uma thread STA para serem criados
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($ScriptApps)
    $script:appsJob = @{ Ps = $ps; Rs = $rs; Handle = $ps.BeginInvoke() }
    $timerApps.Start()
}

function Update-GradeApps {
    $lista = $script:Apps
    if ($ui.rbAppProg.IsChecked) { $lista = @($lista | Where-Object { $_.Tipo -eq 'Programa' }) }
    elseif ($ui.rbAppStore.IsChecked) { $lista = @($lista | Where-Object { $_.Tipo -eq 'Microsoft Store' }) }
    $filtro = $ui.txtBuscaApp.Text.Trim()
    if ($filtro) { $lista = @($lista | Where-Object { $_.Nome -like "*$filtro*" -or $_.Editor -like "*$filtro*" }) }
    $lista = if ($ui.rbAppTam.IsChecked) { @($lista | Sort-Object TamanhoMB -Descending) }
             elseif ($ui.rbAppData.IsChecked) { @($lista | Sort-Object { if ($_.Data) { $_.Data } else { [datetime]::MinValue } } -Descending) }
             else { @($lista | Sort-Object Nome) }
    $ui.dgApps.ItemsSource = $lista
    $totalGB = ($script:Apps | Measure-Object TamanhoMB -Sum).Sum / 1024
    $ui.txtAppsInfo.Text = if ($filtro) { "$($lista.Count) resultado(s) para '$filtro'" }
                           else { '{0} aplicativos instalados · programas ocupam cerca de {1:N1} GB' -f $script:Apps.Count, $totalGB }
}

$timerApps = New-Object Windows.Threading.DispatcherTimer
$timerApps.Interval = [TimeSpan]::FromMilliseconds(300)
$timerApps.add_Tick({
    if (-not $script:appsJob -or -not $script:appsJob.Handle.IsCompleted) { return }
    $timerApps.Stop()
    try {
        $res = $script:appsJob.Ps.EndInvoke($script:appsJob.Handle) | Select-Object -Last 1
        $script:Apps = @($res.Apps)
        $ui.icRestos.ItemsSource = @($res.Restos)
        Update-GradeApps
    } catch {
        $ui.txtAppsInfo.Text = "Erro ao listar: $($_.Exception.Message)"
    } finally {
        $script:appsJob.Ps.Dispose()
        $script:appsJob.Rs.Dispose()
        $script:appsJob = $null
        $ui.btnAppsAtualizar.IsEnabled = $true
    }
})

function Invoke-Desinstalar($app) {
    if ($app.Tipo -eq 'Microsoft Store') {
        Invoke-Comando @{
            Titulo = "Desinstalar $($app.Nome)"; Admin = $false
            Descricao = 'Remove o app da Microsoft Store desta conta. Dá para instalar de novo pela Loja quando quiser.'
            Script = "Remove-AppxPackage -Package '$($app.Pacote)'; Write-Host 'App removido.' -ForegroundColor Green"
        }
        return
    }
    $resp = [Windows.MessageBox]::Show("Desinstalar $($app.Nome)?`n`nO desinstalador do próprio programa vai abrir — siga as etapas dele. Se pedir permissão de administrador, confirme.",
        $AppNome, 'YesNo', 'Question')
    if ($resp -ne 'Yes') { return }
    $cmd = "$($app.Desinstalar)".Trim()
    try {
        if ($cmd -match '(?i)msiexec(\.exe)?\b.*?/[IX]\s*(\{[0-9A-F-]+\})') {
            Start-Process msiexec.exe -ArgumentList "/x $($matches[2])"
        } else {
            # Separa o executável dos argumentos: "C:\x\unins.exe" /arg  ou  C:\x\unins.exe /arg
            if ($cmd.StartsWith('"')) {
                $fim = $cmd.IndexOf('"', 1)
                $exe = $cmd.Substring(1, $fim - 1); $argumentos = $cmd.Substring($fim + 1).Trim()
            } else {
                $i = $cmd.ToLower().IndexOf('.exe')
                if ($i -ge 0) { $exe = $cmd.Substring(0, $i + 4); $argumentos = $cmd.Substring($i + 4).Trim() } else { $exe = $cmd; $argumentos = '' }
            }
            if ($argumentos) { Start-Process -FilePath $exe -ArgumentList $argumentos } else { Start-Process -FilePath $exe }
        }
        $ui.txtAppsInfo.Text = "Desinstalador de $($app.Nome) aberto. Quando terminar, clique em 'Atualizar lista'."
    } catch {
        [Windows.MessageBox]::Show("Não foi possível abrir o desinstalador.`n$($_.Exception.Message)`n`nTente por Configurações > Aplicativos.", $AppNome) | Out-Null
    }
}

function Invoke-LimparRestos($resto) {
    $nome = $resto.Nome -replace "'", "''"
    $linhas = @(
        "foreach (`$classe in 'AntiVirusProduct', 'AntiSpywareProduct', 'FirewallProduct') {"
        "    Get-CimInstance -Namespace root\SecurityCenter2 -ClassName `$classe -ErrorAction SilentlyContinue | Where-Object { `$_.displayName -eq '$nome' } | ForEach-Object { Remove-CimInstance -InputObject `$_; Write-Host ""Registro removido: `$classe"" -ForegroundColor Green }"
        "}"
    )
    foreach ($pa in $resto.Pastas) {
        $p = $pa -replace "'", "''"
        $linhas += "if (Test-Path -LiteralPath '$p') { Remove-Item -LiteralPath '$p' -Recurse -Force -ErrorAction SilentlyContinue; if (Test-Path -LiteralPath '$p') { Write-Host 'Alguns arquivos em uso ficaram em: $p (reinicie e rode de novo)' -ForegroundColor Yellow } else { Write-Host 'Pasta apagada: $p' -ForegroundColor Green } }"
    }
    Invoke-Comando @{
        Titulo = "Limpar restos de $($resto.Nome)"; Admin = $true
        Descricao = "Vai remover:`n$($resto.Itens)`n`nNão dá para desfazer, mas o programa já não está instalado. Depois, clique em 'Atualizar lista' e em 'Verificar agora' na aba Segurança."
        Script = $linhas -join "`n"
    }
}

$ui.dgApps.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent, [Windows.RoutedEventHandler]{
    param($s, $e)
    if ($e.OriginalSource -is [Windows.Controls.Button] -and $e.OriginalSource.Tag -eq 'desinstalar') { Invoke-Desinstalar $e.OriginalSource.DataContext }
})
$ui.icRestos.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent, [Windows.RoutedEventHandler]{
    param($s, $e)
    if ($e.OriginalSource -is [Windows.Controls.Button] -and $e.OriginalSource.Tag -eq 'limpar') { Invoke-LimparRestos $e.OriginalSource.DataContext }
})
$ui.btnAppsAtualizar.add_Click({ Start-Apps })
$ui.txtBuscaApp.add_TextChanged({
    $ui.phBuscaApp.Visibility = if ($ui.txtBuscaApp.Text) { 'Collapsed' } else { 'Visible' }
    Update-GradeApps
})
foreach ($rb in $ui.rbAppNome, $ui.rbAppTam, $ui.rbAppData, $ui.rbAppTodos, $ui.rbAppProg, $ui.rbAppStore) {
    $rb.add_Checked({ if ($ui.dgApps) { Update-GradeApps } })
}
#endregion

#region Aparência
$script:CorEscolhida = $Config.CorDestaque.ToUpper()
$Amostras = @()

# Mostra como a cor escolhida fica antes de aplicar
function Update-Previa {
    $pal = Get-Paleta $script:CorEscolhida
    $ui.bdPrevBotao.Background = Get-Pincel $script:CorEscolhida
    $ui.bdPrevMenu.Background = Get-Pincel ('#' + $pal['1A2614'])
    $ui.bdPrevMenu.BorderBrush = Get-Pincel ('#' + $pal['2F4A16'])
    $ui.txtPrevMenu.Foreground = Get-Pincel $script:CorEscolhida
    $ui.bdPrevBarra.Background = Get-Pincel $script:CorEscolhida
    $ui.elPrevTrilha.Stroke = Get-Pincel ('#' + $pal['1C2418'])
    $ui.anPrev.Stroke = Get-Pincel $script:CorEscolhida
    Set-Arco $ui.anPrev 68 27 23.5
    $ui.elCorPrev.Fill = Get-Pincel $script:CorEscolhida
    foreach ($a in $Amostras) {
        $a.Anel.BorderBrush = Get-Pincel $(if ($a.Hex -eq $script:CorEscolhida) { '#F2F2F5' } else { 'Transparent' })
    }
}

foreach ($nome in $Temas.Keys) {
    $hex = $Temas[$nome].ToUpper()
    $item = New-Object Windows.Controls.StackPanel
    $item.Margin = '0,0,22,12'
    $item.Cursor = 'Hand'
    $item.Tag = $hex
    $anel = New-Object Windows.Controls.Border
    $anel.Width = 60; $anel.Height = 60; $anel.CornerRadius = 30; $anel.BorderThickness = 3; $anel.Padding = 5
    $bola = New-Object Windows.Shapes.Ellipse
    $bola.Fill = Get-Pincel $hex
    $anel.Child = $bola
    $rot = New-Object Windows.Controls.TextBlock
    $rot.Text = $nome; $rot.HorizontalAlignment = 'Center'; $rot.Margin = '0,6,0,0'; $rot.Foreground = Get-Pincel '#C8C8CE'
    [void]$item.Children.Add($anel); [void]$item.Children.Add($rot)
    $item.add_MouseLeftButtonUp({
        param($s, $e)
        $script:CorEscolhida = $s.Tag
        $ui.txtCorHex.Text = $s.Tag
        Update-Previa
    })
    [void]$ui.wpCores.Children.Add($item)
    $Amostras += @{ Hex = $hex; Anel = $anel }
}

$ui.txtCorHex.Text = $script:CorEscolhida
$ui.chkBrilho.IsChecked = [bool]$Config.Brilho
$ui.txtCorHex.add_TextChanged({
    $t = $ui.txtCorHex.Text.Trim()
    if ($t -match '^#?[0-9A-Fa-f]{6}$') {
        $script:CorEscolhida = ('#' + $t.TrimStart('#')).ToUpper()
        $ui.txtCorAviso.Text = 'Cor válida'
        Update-Previa
    } else {
        $ui.txtCorAviso.Text = 'Digite um código como #FF6B00'
    }
})

function Restart-App {
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`" -Reiniciando"
    Exit-App
}

$ui.btnAplicarTema.add_Click({
    $Config.CorDestaque = $script:CorEscolhida
    $Config.Brilho = [bool]$ui.chkBrilho.IsChecked
    Save-Config
    Restart-App
})
if ($window.Icon) { $ui.imgIcone.Source = $window.Icon }
$AtalhoDesktop = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Monitor PC.lnk'
$AtalhoMenu = Join-Path ([Environment]::GetFolderPath('Programs')) 'Monitor PC.lnk'
function Update-TextoAtalhos {
    $tem = @()
    if (Test-Path $AtalhoDesktop) { $tem += 'Área de Trabalho' }
    if (Test-Path $AtalhoMenu) { $tem += 'menu Iniciar' }
    $ui.txtAtalho.Text = if ($tem.Count) { '✓ Atalho em: ' + ($tem -join ' e ') } else { '' }
}
$ui.btnAtalhoDesktop.add_Click({
    try { New-Atalho $AtalhoDesktop; Update-TextoAtalhos } catch { $ui.txtAtalho.Text = "Não foi possível criar: $($_.Exception.Message)" }
})
$ui.btnAtalhoIniciar.add_Click({
    try { New-Atalho $AtalhoMenu; Update-TextoAtalhos } catch { $ui.txtAtalho.Text = "Não foi possível criar: $($_.Exception.Message)" }
})
Update-TextoAtalhos

$ui.btnTema.add_Click({
    foreach ($m in $Menu) { $m.IsChecked = $false }
    $ui.tabs.SelectedIndex = 7
})
Update-Previa
#endregion

#region Atualização automática
# versao.json publicado: { "versao": "2.6.0", "notas": "...", "arquivo": "<url do MonitorPC.ps1>", "sha256": "<hash do arquivo>" }
$UrlAtualizacao = if ($Config.UrlAtualizacao) { $Config.UrlAtualizacao } else { $UrlAtualizacaoPadrao }
$script:Novidade = $null
$script:atuJob = $null
$ui.txtVersao.Text = "Monitor PC versão $VersaoApp"

$ScriptVerificar = {
    param($url)
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    try {
        # "?t=" evita receber uma cópia antiga guardada em cache (só faz sentido em endereços da internet)
        $alvo = if ($url -like 'http*') { "$url`?t=$([DateTime]::UtcNow.Ticks)" } else { $url }
        # Lê sempre como UTF-8 (acentos das notas) e ignora a marca BOM, se houver
        $wc = New-Object Net.WebClient
        $wc.Encoding = [Text.Encoding]::UTF8
        $texto = $wc.DownloadString($alvo)
        $info = ($texto -replace '^(﻿|ï»¿)', '') | ConvertFrom-Json
        if (-not $info.versao) { throw 'o arquivo de versão não está no formato esperado.' }
        @{ Ok = $true; Info = $info }
    } catch { @{ Ok = $false; Erro = $_.Exception.Message } }
}

function Start-VerificarAtualizacao([bool]$manual) {
    if (-not $UrlAtualizacao) {
        if ($manual) { $ui.txtAtualizacao.Text = 'Atualização automática ainda não configurada nesta cópia.' }
        return
    }
    if ($script:atuJob) { return }
    if ($manual) { $ui.txtAtualizacao.Text = 'Verificando...' }
    $ps = [powershell]::Create()
    [void]$ps.AddScript($ScriptVerificar).AddArgument($UrlAtualizacao)
    $script:atuJob = @{ Ps = $ps; Handle = $ps.BeginInvoke(); Manual = $manual }
    $timerAtu.Start()
}

$timerAtu = New-Object Windows.Threading.DispatcherTimer
$timerAtu.Interval = [TimeSpan]::FromMilliseconds(500)
$timerAtu.add_Tick({
    $j = $script:atuJob
    if (-not $j -or -not $j.Handle.IsCompleted) { return }
    $timerAtu.Stop()
    try {
        $r = $j.Ps.EndInvoke($j.Handle) | Select-Object -Last 1
        if (-not $r.Ok) {
            if ($j.Manual) { $ui.txtAtualizacao.Text = "Não foi possível verificar: $($r.Erro)" }
        } else {
            $nova = $null
            try { $nova = [version]$r.Info.versao } catch {}
            if ($nova -and $nova -gt [version]$VersaoApp) {
                $script:Novidade = $r.Info
                $ui.txtAtualizacao.Text = "Versão $($r.Info.versao) disponível." + $(if ($r.Info.notas) { " Novidades: $($r.Info.notas)" } else { '' })
                $ui.btnInstalarAtualizacao.Visibility = 'Visible'
                $ui.btnNovaVersao.Content = "Atualizar para $($r.Info.versao)"
                $ui.btnNovaVersao.Visibility = 'Visible'
            } else {
                $ui.txtAtualizacao.Text = "Você está na versão mais recente. Verificado às $((Get-Date).ToString('HH:mm'))."
            }
        }
    } catch {
        if ($j.Manual) { $ui.txtAtualizacao.Text = "Erro ao verificar: $($_.Exception.Message)" }
    } finally { $j.Ps.Dispose(); $script:atuJob = $null }
})

function Install-Atualizacao {
    $n = $script:Novidade
    if (-not $n) { return }
    $msg = "Atualizar o Monitor PC de $VersaoApp para $($n.versao)?" + $(if ($n.notas) { "`n`nNovidades: $($n.notas)" } else { '' }) +
           "`n`nSuas configurações continuam. Uma cópia da versão atual fica salva como MonitorPC.ps1.bak."
    if ([Windows.MessageBox]::Show($msg, $AppNome, 'YesNo', 'Question') -ne 'Yes') { return }
    $tmp = Join-Path $env:TEMP "MonitorPC_$($n.versao).ps1"
    try {
        $ui.txtAtualizacao.Text = 'Baixando...'
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        (New-Object Net.WebClient).DownloadFile($n.arquivo, $tmp)
        # Confere a impressão digital: só instala se for exatamente o arquivo publicado
        $hash = (Get-FileHash -Path $tmp -Algorithm SHA256).Hash
        if (-not $n.sha256 -or $hash -ne "$($n.sha256)".ToUpper()) { throw 'o arquivo baixado não confere com a versão publicada (SHA-256 diferente). Nada foi alterado.' }
        $err = $null; [void][Management.Automation.Language.Parser]::ParseFile($tmp, [ref]$null, [ref]$err)
        if ($err.Count) { throw 'o arquivo baixado está com erro. Nada foi alterado.' }
        Copy-Item -LiteralPath $PSCommandPath -Destination "$PSCommandPath.bak" -Force
        Copy-Item -LiteralPath $tmp -Destination $PSCommandPath -Force
        Remove-Item $tmp -ErrorAction SilentlyContinue
        Restart-App
    } catch {
        Remove-Item $tmp -ErrorAction SilentlyContinue
        $ui.txtAtualizacao.Text = "Não foi possível atualizar: $($_.Exception.Message)"
        [Windows.MessageBox]::Show("Não foi possível atualizar: $($_.Exception.Message)", $AppNome) | Out-Null
    }
}

$ui.btnVerificarAtualizacao.add_Click({ Start-VerificarAtualizacao $true })
$ui.btnInstalarAtualizacao.add_Click({ Install-Atualizacao })
$ui.btnNovaVersao.add_Click({ Install-Atualizacao })
$ui.txtAtualizacao.Text = if ($UrlAtualizacao) { 'Verifica sozinho ao abrir e uma vez por dia.' } else { 'Atualização automática ainda não configurada nesta cópia.' }
# Verifica 20 s depois de abrir e depois a cada 24 h
$timerAtuAgenda = New-Object Windows.Threading.DispatcherTimer
$timerAtuAgenda.Interval = [TimeSpan]::FromSeconds(20)
$timerAtuAgenda.add_Tick({ $timerAtuAgenda.Interval = [TimeSpan]::FromHours(24); Start-VerificarAtualizacao $false })
$timerAtuAgenda.Start()
#endregion

#region Comandos de otimização
$PreludioCmd = @'
$ErrorActionPreference = 'Continue'
function Limpar-Pasta($Caminho) {
    if (-not (Test-Path $Caminho)) { Write-Host "Pasta não encontrada: $Caminho" -ForegroundColor Yellow; return }
    $antes = (Get-ChildItem $Caminho -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
    Get-ChildItem $Caminho -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    $depois = (Get-ChildItem $Caminho -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
    Write-Host ('{0}: {1:N1} MB liberados (arquivos em uso foram mantidos)' -f $Caminho, (($antes - $depois) / 1MB)) -ForegroundColor Green
}
'@

$Comandos = @(
    # ---------- BÁSICO ----------
    @{ Nivel = 'Básico'; Titulo = 'Limpar arquivos temporários'; Admin = $false; Risco = 'Seguro'; Glifo = 0xE74D
       Descricao = 'Apaga a pasta TEMP do seu usuário (restos de instalações e programas). Arquivos em uso são ignorados.'
       Script = 'Limpar-Pasta $env:TEMP' }
    @{ Nivel = 'Básico'; Titulo = 'Esvaziar a Lixeira'; Admin = $false; Risco = 'Seguro'; Glifo = 0xE74D
       Descricao = 'Apaga de vez tudo o que está na Lixeira. Não dá para desfazer — confira antes se não tem nada importante lá.'
       Script = 'Clear-RecycleBin -Force -ErrorAction SilentlyContinue; Write-Host "Lixeira esvaziada." -ForegroundColor Green' }
    @{ Nivel = 'Básico'; Titulo = 'Limpar cache de DNS'; Admin = $false; Risco = 'Seguro'; Glifo = 0xEC05
       Descricao = 'Resolve sites que não abrem ou abrem lentos depois de mudanças na internet.'
       Script = 'ipconfig /flushdns' }
    @{ Nivel = 'Básico'; Titulo = 'Reiniciar o Explorer'; Admin = $false; Risco = 'Seguro'; Glifo = 0xE737
       Descricao = 'Reinicia a barra de tarefas e as pastas do Windows. Útil quando a barra trava ou o menu Iniciar não abre.'
       Script = 'Stop-Process -Name explorer -Force; Start-Sleep 2; if (-not (Get-Process explorer -ErrorAction SilentlyContinue)) { Start-Process explorer }; Write-Host "Explorer reiniciado." -ForegroundColor Green' }
    @{ Nivel = 'Básico'; Titulo = 'Gerenciar programas da inicialização'; Admin = $false; Risco = 'Seguro'; Glifo = 0xE7E8
       Descricao = 'Abre a tela do Windows para desligar programas que abrem sozinhos com o PC — a otimização que mais acelera o boot.'
       Script = 'Start-Process "ms-settings:startupapps"; Write-Host "Desative o que você não precisa ao ligar o PC (Discord, Steam, Spotify, etc.)."' }
    @{ Nivel = 'Básico'; Titulo = 'Limpeza de Disco do Windows'; Admin = $false; Risco = 'Seguro'; Glifo = 0xEDA2
       Descricao = 'Abre a ferramenta oficial de limpeza. Clique em "Limpar arquivos do sistema" para liberar ainda mais espaço.'
       Script = 'Start-Process cleanmgr.exe' }
    @{ Nivel = 'Básico'; Titulo = 'Limpar cache da Microsoft Store'; Admin = $false; Risco = 'Seguro'; Glifo = 0xE895
       Descricao = 'Corrige a Loja que não baixa ou não atualiza apps. A Loja abre sozinha no final.'
       Script = 'Start-Process wsreset.exe; Write-Host "Aguarde a Microsoft Store abrir."' }

    # ---------- INTERMEDIÁRIO ----------
    @{ Nivel = 'Intermediário'; Titulo = 'Limpar Temp do Windows'; Admin = $true; Risco = 'Seguro'; Glifo = 0xE74D
       Descricao = 'Apaga a pasta C:\Windows\Temp (temporários do sistema, geralmente maior que a do usuário).'
       Script = 'Limpar-Pasta "$env:windir\Temp"' }
    @{ Nivel = 'Intermediário'; Titulo = 'Limpar cache do Windows Update'; Admin = $true; Risco = 'Seguro'; Glifo = 0xE895
       Descricao = 'Apaga atualizações já baixadas/instaladas. Libera vários GB e corrige atualização travada.'
       Script = 'Stop-Service wuauserv, bits -Force -ErrorAction SilentlyContinue; Limpar-Pasta "$env:windir\SoftwareDistribution\Download"; Start-Service wuauserv, bits -ErrorAction SilentlyContinue' }
    @{ Nivel = 'Intermediário'; Titulo = 'Otimizar unidades (TRIM / desfragmentar)'; Admin = $true; Risco = 'Seguro'; Glifo = 0xEDA2
       Descricao = 'Faz TRIM nos SSDs e desfragmenta HDs automaticamente (o Windows escolhe o certo para cada um).'
       Script = 'Get-Volume | Where-Object { $_.DriveLetter -and $_.DriveType -eq "Fixed" } | ForEach-Object { Write-Host "Otimizando $($_.DriveLetter):" -ForegroundColor Cyan; Optimize-Volume -DriveLetter $_.DriveLetter -Verbose }' }
    @{ Nivel = 'Intermediário'; Titulo = 'Verificar o disco (somente leitura)'; Admin = $true; Risco = 'Seguro'; Glifo = 0xE9D9
       Descricao = 'Procura erros no sistema de arquivos do C: sem mexer em nada e sem reiniciar.'
       Script = 'chkdsk $env:SystemDrive /scan' }
    @{ Nivel = 'Intermediário'; Titulo = 'Plano de energia Alto desempenho'; Admin = $false; Risco = 'Moderado'; Glifo = 0xE945
       Descricao = 'Deixa o processador sempre pronto em vez de economizar energia. Em notebook gasta mais bateria.'
       Script = 'powercfg /setactive SCHEME_MIN; powercfg /getactivescheme' }
    @{ Nivel = 'Intermediário'; Titulo = 'Ajustar efeitos visuais'; Admin = $false; Risco = 'Seguro'; Glifo = 0xE7F4
       Descricao = 'Abre as opções de desempenho visual. Em PC fraco, escolha "Ajustar para obter melhor desempenho".'
       Script = 'Start-Process SystemPropertiesPerformance.exe' }
    @{ Nivel = 'Intermediário'; Titulo = 'Desativar hibernação'; Admin = $true; Risco = 'Moderado'; Glifo = 0xE7E8
       Descricao = 'Apaga o arquivo hiberfil.sys (vários GB). Desliga também a Inicialização Rápida. Para voltar: powercfg /h on.'
       Script = 'powercfg /h off; Write-Host "Hibernação desativada." -ForegroundColor Green' }

    # ---------- AVANÇADO ----------
    @{ Nivel = 'Avançado'; Titulo = 'Criar ponto de restauração'; Admin = $true; Risco = 'Seguro'; Glifo = 0xE81C
       Descricao = 'Salva o estado atual do Windows para poder voltar atrás se algo der errado. Faça isso ANTES dos itens abaixo.'
       Script = 'Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction SilentlyContinue; Checkpoint-Computer -Description "Monitor PC" -RestorePointType MODIFY_SETTINGS; Write-Host "Pronto. (O Windows só permite 1 ponto automático a cada 24 h.)" -ForegroundColor Green' }
    @{ Nivel = 'Avançado'; Titulo = 'Reparar arquivos do sistema (SFC)'; Admin = $true; Risco = 'Seguro'; Glifo = 0xE9D9
       Descricao = 'Verifica e repara arquivos corrompidos do Windows. Demora 10-20 minutos.'
       Script = 'sfc /scannow' }
    @{ Nivel = 'Avançado'; Titulo = 'Reparar imagem do Windows (DISM)'; Admin = $true; Risco = 'Seguro'; Glifo = 0xE9D9
       Descricao = 'Baixa da Microsoft e repara a base do Windows. Use quando o SFC não conseguir corrigir. Precisa de internet.'
       Script = 'DISM /Online /Cleanup-Image /RestoreHealth' }
    @{ Nivel = 'Avançado'; Titulo = 'Limpar componentes antigos (WinSxS)'; Admin = $true; Risco = 'Moderado'; Glifo = 0xE74D
       Descricao = 'Remove versões antigas de atualizações do Windows. Libera espaço, mas não dá mais para desinstalar essas atualizações.'
       Script = 'DISM /Online /Cleanup-Image /StartComponentCleanup' }
    @{ Nivel = 'Avançado'; Titulo = 'Corrigir disco no próximo boot (CHKDSK /f)'; Admin = $true; Risco = 'Moderado'; Glifo = 0xEDA2
       Descricao = 'Agenda a correção do C: para quando reiniciar (responda S). O boot seguinte pode demorar bastante.'
       Script = 'chkdsk $env:SystemDrive /f' }
    @{ Nivel = 'Avançado'; Titulo = 'Resetar a rede (Winsock e TCP/IP)'; Admin = $true; Risco = 'Moderado'; Glifo = 0xEC05
       Descricao = 'Resolve internet que conecta mas não navega. Precisa reiniciar o PC. Configurações de IP fixo/VPN podem ser perdidas.'
       Script = 'netsh winsock reset; netsh int ip reset; ipconfig /flushdns; Write-Host "Reinicie o PC para concluir." -ForegroundColor Yellow' }
    @{ Nivel = 'Avançado'; Titulo = 'Plano Desempenho Máximo'; Admin = $true; Risco = 'Moderado'; Glifo = 0xEC4A
       Descricao = 'Ativa o plano oculto "Desempenho Máximo" do Windows. Ótimo para desktops gamer; evite em notebook.'
       Script = '$saida = powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61; $guid = [regex]::Match("$saida", "[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}").Value; powercfg /setactive $guid; powercfg /getactivescheme' }
    @{ Nivel = 'Avançado'; Titulo = 'Diagnóstico de memória RAM'; Admin = $true; Risco = 'Seguro'; Glifo = 0xE964
       Descricao = 'Testa a RAM reiniciando o PC (leva uns 15 min). O resultado aparece ao entrar no Windows de novo.'
       Script = 'Start-Process mdsched.exe' }
)

function Invoke-Comando($cmd) {
    $aviso = "$($cmd.Titulo)`n`n$($cmd.Descricao)"
    if ($cmd.Admin) { $aviso += "`n`nO Windows vai pedir permissão de administrador." }
    $resp = [Windows.MessageBox]::Show($aviso, 'Executar comando?', 'YesNo', 'Question')
    if ($resp -ne 'Yes') { return }

    $titulo = $cmd.Titulo -replace "'", "''"
    $completo = "$PreludioCmd`n`$Host.UI.RawUI.WindowTitle = '$titulo'`nWrite-Host '== $titulo ==' -ForegroundColor Cyan`n$($cmd.Script)`nWrite-Host ''`nWrite-Host 'Concluído. Pode fechar esta janela.' -ForegroundColor Green"
    $codificado = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($completo))
    $argumentos = "-NoProfile -ExecutionPolicy Bypass -NoExit -EncodedCommand $codificado"
    try {
        if ($cmd.Admin -and -not $EhAdmin) {
            Start-Process powershell.exe -ArgumentList $argumentos -Verb RunAs
        } else {
            Start-Process powershell.exe -ArgumentList $argumentos
        }
    } catch {
        [Windows.MessageBox]::Show('O comando não foi executado (permissão negada ou cancelada).', $AppNome) | Out-Null
    }
}

function New-Selo($texto, $hex) {
    $b = New-Object Windows.Controls.Border
    $b.Background = Get-Pincel ('#22' + $hex.TrimStart('#'))
    $b.CornerRadius = 8
    $b.Padding = '8,2'
    $b.Margin = '10,0,0,0'
    $b.VerticalAlignment = 'Center'
    $t = New-Object Windows.Controls.TextBlock
    $t.Text = $texto; $t.FontSize = 11; $t.Foreground = Get-Pincel $hex
    $b.Child = $t
    $b
}

$Niveis = @(
    @{ Nome = 'Básico';        Cor = $Cor.Lima;     Texto = 'Seguro para qualquer pessoa. Pode usar sempre que quiser.' }
    @{ Nome = 'Intermediário'; Cor = $Cor.Ambar;    Texto = 'Mexe em configurações e pastas do sistema. Seguro, mas leia a descrição.' }
    @{ Nome = 'Avançado';      Cor = $Cor.Vermelho; Texto = 'Reparos profundos do Windows. Use quando houver problema de verdade.' }
)

foreach ($nivel in $Niveis) {
    $cab = New-Object Windows.Controls.StackPanel
    $cab.Orientation = 'Horizontal'
    $cab.Margin = '10,18,10,6'
    $t1 = New-Object Windows.Controls.TextBlock
    $t1.Text = $nivel.Nome; $t1.FontSize = 18; $t1.FontWeight = 'SemiBold'; $t1.VerticalAlignment = 'Center'
    $t2 = New-Object Windows.Controls.TextBlock
    $t2.Text = $nivel.Texto; $t2.Foreground = Get-Pincel $Cor.Suave; $t2.Margin = '14,0,0,0'; $t2.VerticalAlignment = 'Center'
    $ponto = New-Object Windows.Shapes.Ellipse
    $ponto.Width = 10; $ponto.Height = 10; $ponto.Fill = Get-Pincel $nivel.Cor; $ponto.Margin = '0,0,10,0'; $ponto.VerticalAlignment = 'Center'
    [void]$cab.Children.Add($ponto); [void]$cab.Children.Add($t1); [void]$cab.Children.Add($t2)
    [void]$ui.spOtimizar.Children.Add($cab)

    $grade = New-Object Windows.Controls.Primitives.UniformGrid
    $grade.Columns = 2
    foreach ($cmd in $Comandos | Where-Object { $_.Nivel -eq $nivel.Nome }) {
        $card = New-Object Windows.Controls.Border
        $card.Style = $window.FindResource('Card')
        $card.Padding = '18,14'
        $dp = New-Object Windows.Controls.DockPanel
        $btn = New-Object Windows.Controls.Button
        $btn.Content = 'Executar'
        $btn.Tag = $cmd
        $btn.VerticalAlignment = 'Center'
        $btn.Margin = '14,0,0,0'
        $btn.Padding = '14,7'
        $btn.Background = Get-Pincel $nivel.Cor
        $btn.add_Click({ param($s, $e) Invoke-Comando $s.Tag })
        [Windows.Controls.DockPanel]::SetDock($btn, 'Right')
        $bolha = New-Object Windows.Controls.Border
        $bolha.Style = $window.FindResource('Bolha')
        $bolha.Width = 40; $bolha.Height = 40; $bolha.CornerRadius = 12; $bolha.VerticalAlignment = 'Top'; $bolha.Margin = '0,0,14,0'
        $ic = New-Object Windows.Controls.TextBlock
        $ic.Style = $window.FindResource('Icone')
        $ic.Text = [string][char]$cmd.Glifo
        $ic.Foreground = Get-Pincel $nivel.Cor
        $bolha.Child = $ic
        [Windows.Controls.DockPanel]::SetDock($bolha, 'Left')
        $info = New-Object Windows.Controls.StackPanel
        $linhaTitulo = New-Object Windows.Controls.WrapPanel
        $tt = New-Object Windows.Controls.TextBlock
        $tt.Text = $cmd.Titulo; $tt.FontWeight = 'SemiBold'; $tt.FontSize = 14
        [void]$linhaTitulo.Children.Add($tt)
        if ($cmd.Admin) { [void]$linhaTitulo.Children.Add((New-Selo 'Admin' '#C9A2FF')) }
        $corRisco = if ($cmd.Risco -eq 'Seguro') { $Cor.Lima } else { $Cor.Ambar }
        [void]$linhaTitulo.Children.Add((New-Selo $cmd.Risco $corRisco))
        $td = New-Object Windows.Controls.TextBlock
        $td.Text = $cmd.Descricao; $td.TextWrapping = 'Wrap'; $td.Foreground = Get-Pincel $Cor.Suave; $td.Margin = '0,5,0,0'; $td.FontSize = 12
        [void]$info.Children.Add($linhaTitulo); [void]$info.Children.Add($td)
        [void]$dp.Children.Add($btn); [void]$dp.Children.Add($bolha); [void]$dp.Children.Add($info)
        $card.Child = $dp
        [void]$grade.Children.Add($card)
    }
    [void]$ui.spOtimizar.Children.Add($grade)
}
#endregion

#region Manutenção agendada (tela)
$NomeTarefa = 'Monitor PC - Manutenção semanal'
$DiasSemana = [ordered]@{ Sunday = 'Dom'; Monday = 'Seg'; Tuesday = 'Ter'; Wednesday = 'Qua'; Thursday = 'Qui'; Friday = 'Sex'; Saturday = 'Sáb' }
$NomesDias = @{ Sunday = 'domingo'; Monday = 'segunda'; Tuesday = 'terça'; Wednesday = 'quarta'; Thursday = 'quinta'; Friday = 'sexta'; Saturday = 'sábado' }
$RbDias = @{}
foreach ($d in $DiasSemana.Keys) {
    $rb = New-Object Windows.Controls.RadioButton
    $rb.Style = $window.FindResource('Seg'); $rb.GroupName = 'manutdia'; $rb.Content = $DiasSemana[$d]; $rb.Tag = $d
    $rb.IsChecked = $d -eq $Config.ManutDia
    [void]$ui.spManutDias.Children.Add($rb)
    $RbDias[$d] = $rb
}
$ui.chkManut.IsChecked = [bool]$Config.ManutAtiva
$ui.txtManutHora.Text = $Config.ManutHora
$ui.chkManutTemp.IsChecked = [bool]$Config.ManutTemp
$ui.chkManutLixeira.IsChecked = [bool]$Config.ManutLixeira
$ui.chkManutSistema.IsChecked = [bool]$Config.ManutSistema -and $EhAdmin
$ui.chkManutSistema.IsEnabled = $EhAdmin

function Update-StatusManut([string]$extra = '') {
    $partes = @()
    $tarefa = Get-ScheduledTask -TaskName $NomeTarefa -ErrorAction SilentlyContinue
    if ($tarefa) {
        $prox = (Get-ScheduledTaskInfo -TaskName $NomeTarefa -ErrorAction SilentlyContinue).NextRunTime
        $partes += "Agendada: toda $($NomesDias[$Config.ManutDia]) às $($Config.ManutHora)" + $(if ($prox) { " · próxima em $($prox.ToString('dd/MM HH:mm'))" } else { '' })
    } else {
        $partes += 'Não agendada.'
    }
    if (Test-Path $ManutLogPath) {
        $ult = Get-Content $ManutLogPath -Encoding UTF8 -ErrorAction SilentlyContinue | Select-Object -Last 1
        if ($ult) { $partes += "Última: $ult" }
    }
    if ($extra) { $partes += $extra }
    $ui.txtManutStatus.Text = $partes -join "`n"
}

function Save-Manut {
    $hora = $ui.txtManutHora.Text.Trim()
    if ($hora -notmatch '^([01]?\d|2[0-3]):[0-5]\d$') { Update-StatusManut 'Horário inválido — use o formato 12:00.'; return }
    $Config.ManutAtiva = [bool]$ui.chkManut.IsChecked
    $Config.ManutHora = $hora
    $Config.ManutDia = ($RbDias.Keys | Where-Object { $RbDias[$_].IsChecked } | Select-Object -First 1)
    if (-not $Config.ManutDia) { $Config.ManutDia = 'Sunday' }
    $Config.ManutTemp = [bool]$ui.chkManutTemp.IsChecked
    $Config.ManutLixeira = [bool]$ui.chkManutLixeira.IsChecked
    if ($EhAdmin) { $Config.ManutSistema = [bool]$ui.chkManutSistema.IsChecked }
    Save-Config
    try {
        if ($Config.ManutAtiva) {
            $acao = if ($script:TemLancador -or (Update-Lancador)) {
                New-ScheduledTaskAction -Execute $LancadorPath -Argument '-Manutencao' -WorkingDirectory $PSScriptRoot
            } else {
                New-ScheduledTaskAction -Execute "$env:windir\System32\WindowsPowerShell\v1.0\powershell.exe" -WorkingDirectory $PSScriptRoot `
                    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`" -Manutencao"
            }
            $gatilho = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $Config.ManutDia -At ([datetime]::ParseExact($hora, 'H:mm', $null))
            $opcoes = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
            # Com a limpeza do Windows marcada (e o Monitor como admin), a tarefa roda com privilégio elevado
            $nivel = if ($EhAdmin -and $Config.ManutSistema) { 'Highest' } else { 'Limited' }
            $quem = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel $nivel
            Register-ScheduledTask -TaskName $NomeTarefa -Action $acao -Trigger $gatilho -Settings $opcoes -Principal $quem `
                -Description 'Limpeza semanal do Monitor PC (temporários e, se marcado, Lixeira antiga e cache do Windows).' -Force -ErrorAction Stop | Out-Null
            Update-StatusManut '✓ Agendamento salvo.'
        } else {
            Unregister-ScheduledTask -TaskName $NomeTarefa -Confirm:$false -ErrorAction SilentlyContinue
            Update-StatusManut '✓ Manutenção automática desligada.'
        }
    } catch {
        $dica = if (-not $EhAdmin) { ' A tarefa pode ter sido criada como administrador: abra o Monitor como admin para alterar.' } else { '' }
        Update-StatusManut "Não foi possível agendar: $($_.Exception.Message)$dica"
    }
}

$timerManut = New-Object Windows.Threading.DispatcherTimer
$timerManut.Interval = [TimeSpan]::FromSeconds(15)
$timerManut.add_Tick({ $timerManut.Stop(); Update-StatusManut; $ui.btnManutAgora.IsEnabled = $true })
$ui.btnManutSalvar.add_Click({ Save-Manut })
$ui.chkManut.add_Click({ Save-Manut })
$ui.btnManutAgora.add_Click({
    $Config.ManutTemp = [bool]$ui.chkManutTemp.IsChecked; $Config.ManutLixeira = [bool]$ui.chkManutLixeira.IsChecked
    if ($EhAdmin) { $Config.ManutSistema = [bool]$ui.chkManutSistema.IsChecked }
    Save-Config
    try {
        if ($script:TemLancador -or (Update-Lancador)) { Start-Process $LancadorPath -ArgumentList '-Manutencao' }
        else { Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`" -Manutencao" -WindowStyle Hidden }
        $ui.btnManutAgora.IsEnabled = $false
        Update-StatusManut 'Limpando... o resultado aparece aqui e num aviso perto do relógio.'
        $timerManut.Start()
    } catch { Update-StatusManut "Não foi possível iniciar: $($_.Exception.Message)" }
})
Update-StatusManut
#endregion

#region Jogos e FPS
# FPS real vem do PresentMon (ferramenta oficial da Intel). Ele lê os eventos de "quadro apresentado"
# do Windows e precisa de permissão de administrador.
$PresentMonPath = Join-Path $PSScriptRoot 'ferramentas\PresentMon.exe'
$PresentMonUrl = 'https://github.com/GameTechDev/PresentMon/releases/download/v2.6.0/PresentMon-2.6.0-x64.exe'
$FpsSync = [hashtable]::Synchronized(@{ Rodando = $false; Atual = $null; Erro = ''; Pid = 0 })
$script:fpsJob = $null

# Programas que desenham na tela mas não são jogos
$NaoJogos = 'dwm', 'explorer', 'chrome', 'msedge', 'msedgewebview2', 'firefox', 'opera', 'brave', 'Discord', 'Spotify', 'Code',
    'Teams', 'ms-teams', 'WhatsApp', 'WhatsApp.Root', 'claude', 'ChatGPT', 'powershell', 'SearchHost', 'StartMenuExperienceHost',
    'TextInputHost', 'ShellExperienceHost', 'Widgets', 'WidgetBoard', 'PhoneExperienceHost', 'NVIDIA Overlay', 'steamwebhelper',
    'EpicWebHelper', 'EpicGamesLauncher', 'Notion', 'ApplicationFrameHost', 'LockApp', 'Monitor PC', 'obs64', 'XboxPcApp', 'GameBar',
    'vlc', 'mpv', 'PotPlayerMini64', 'PotPlayer', 'wmplayer', 'Video.UI', 'Microsoft.Media.Player', 'Netflix', 'Photos', 'mpc-hc64', 'Telegram'

$ScriptFps = {
    param($s, $exe, $excluir)
    $argumentos = '--output_stdout --stop_existing_session --session_name MonitorPC_FPS --no_console_stats --v1_metrics --no_track_input' +
        (($excluir | ForEach-Object { " --exclude `"$_.exe`"" }) -join '')
    $psi = New-Object Diagnostics.ProcessStartInfo($exe, $argumentos)
    $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true; $psi.CreateNoWindow = $true
    try { $p = [Diagnostics.Process]::Start($psi) } catch { $s.Erro = $_.Exception.Message; $s.Rodando = $false; return }
    $s.Pid = $p.Id
    $inv = [Globalization.CultureInfo]::InvariantCulture
    $cab = $null; $iApp = -1; $iPid = -1; $iMs = -1; $iModo = -1
    $quadros = @{}    # pid -> fila de @(segundo, ms)
    $nomes = @{}; $modos = @{}
    $relogio = [Diagnostics.Stopwatch]::StartNew()
    $ultimoCalc = 0.0
    while ($s.Rodando) {
        $linha = $p.StandardOutput.ReadLine()
        if ($null -eq $linha) { break }
        if (-not $cab) {
            $cab = $linha -split ','
            for ($i = 0; $i -lt $cab.Count; $i++) {
                switch -regex ($cab[$i]) {
                    '^Application$' { $iApp = $i }
                    '^ProcessID$' { $iPid = $i }
                    '^(msBetweenPresents|MsBetweenPresents|FrameTime)$' { $iMs = $i }
                    '^PresentMode$' { $iModo = $i }
                }
            }
            continue
        }
        $c = $linha -split ','
        if ($c.Count -le $iMs) { continue }
        $ms = 0.0
        if (-not [double]::TryParse($c[$iMs], [Globalization.NumberStyles]::Float, $inv, [ref]$ms) -or $ms -le 0) { continue }
        $id = $c[$iPid]
        $t = $relogio.Elapsed.TotalSeconds
        if (-not $quadros.ContainsKey($id)) { $quadros[$id] = New-Object System.Collections.Generic.Queue[object] }
        $quadros[$id].Enqueue(@($t, $ms))
        $nomes[$id] = $c[$iApp]
        if ($iModo -ge 0) { $modos[$id] = $c[$iModo] }

        if ($t - $ultimoCalc -ge 0.5) {
            $ultimoCalc = $t
            $melhor = $null; $melhorPontos = 0; $pontosAtual = 0
            foreach ($k in @($quadros.Keys)) {
                $q = $quadros[$k]
                while ($q.Count -and $q.Peek()[0] -lt $t - 10) { [void]$q.Dequeue() }
                if (-not $q.Count) { $quadros.Remove($k); continue }
                # Mesmo filtrando no PresentMon, confere aqui também que não é navegador/Discord/etc.
                if ([IO.Path]::GetFileNameWithoutExtension("$($nomes[$k])") -in $excluir) { continue }
                $ultimo = @($q | Where-Object { $_[0] -ge $t - 1 })
                if ($ultimo.Count -lt 5) { continue }
                # Jogos costumam usar "Hardware: ... Flip" (tela cheia/sem borda otimizada): pesa mais na escolha
                $pontos = $ultimo.Count * $(if ("$($modos[$k])" -like 'Hardware*') { 3 } else { 1 })
                if ($k -eq $s.AtualPid) { $pontosAtual = $pontos }
                if ($pontos -gt $melhorPontos) { $melhorPontos = $pontos; $melhor = $k }
            }
            # Só troca de jogo se o outro for claramente mais ativo (evita pular entre janelas)
            if ($s.AtualPid -and $pontosAtual -ge $melhorPontos * 0.6 -and $pontosAtual -gt 0) { $melhor = $s.AtualPid }
            $s.AtualPid = $melhor
            if ($melhor) {
                $q = $quadros[$melhor]
                $ultimo = @($q | Where-Object { $_[0] -ge $t - 1 } | ForEach-Object { $_[1] })
                $todos = @($q | ForEach-Object { $_[1] } | Sort-Object)
                $p99 = $todos[[Math]::Min($todos.Count - 1, [int][Math]::Floor($todos.Count * 0.99))]
                $media = ($ultimo | Measure-Object -Average).Average
                $s.Atual = @{
                    App = $nomes[$melhor]; Pid = [int]$melhor; Modo = "$($modos[$melhor])"; Hora = Get-Date
                    Fps = 1000.0 / $media; Low1 = 1000.0 / $p99; Ms = $media
                }
            }
        }
    }
    if ($p.HasExited -and $s.Rodando) { $s.Erro = ($p.StandardError.ReadToEnd() -split "`n" | Where-Object { $_ -match 'error' } | Select-Object -First 1) }
    try { if (-not $p.HasExited) { $p.Kill() } } catch {}
    $s.Rodando = $false
}

function Start-Fps {
    if ($script:fpsJob) { return }
    if (-not (Test-Path $PresentMonPath)) { return }
    if (-not $EhAdmin) { return }
    $FpsSync.Rodando = $true; $FpsSync.Erro = ''; $FpsSync.Atual = $null
    $ps = [powershell]::Create()
    [void]$ps.AddScript($ScriptFps).AddArgument($FpsSync).AddArgument($PresentMonPath).AddArgument($NaoJogos)
    $script:fpsJob = @{ Ps = $ps; Handle = $ps.BeginInvoke() }
}

function Stop-Fps {
    $FpsSync.Rodando = $false
    if ($FpsSync.Pid) { try { Stop-Process -Id $FpsSync.Pid -Force -ErrorAction Stop } catch {} }
    if ((Test-Path $PresentMonPath) -and $EhAdmin) {
        try { Start-Process $PresentMonPath -ArgumentList '--terminate_existing_session --session_name MonitorPC_FPS' -WindowStyle Hidden } catch {}
    }
    if ($script:fpsJob) { try { $script:fpsJob.Ps.Dispose() } catch {}; $script:fpsJob = $null }
    $FpsSync.Pid = 0
}

function Update-EstadoFps {
    $ui.btnFpsBaixar.Visibility = 'Collapsed'; $ui.btnFpsAdmin.Visibility = 'Collapsed'
    if (-not (Test-Path $PresentMonPath)) {
        $ui.txtFpsStatus.Text = 'Para medir o FPS é preciso o PresentMon, ferramenta gratuita e oficial da Intel (1 MB). Os diagnósticos abaixo funcionam mesmo sem ele.'
        $ui.btnFpsBaixar.Visibility = 'Visible'
    } elseif (-not $EhAdmin) {
        $ui.txtFpsStatus.Text = 'O Windows só libera a leitura de FPS para programas abertos como administrador.'
        $ui.btnFpsAdmin.Visibility = 'Visible'
    } elseif (-not $ui.chkFps.IsChecked) {
        $ui.txtFpsStatus.Text = 'Medição de FPS desligada.'
    } elseif ($FpsSync.Erro) {
        $ui.txtFpsStatus.Text = "O medidor parou: $($FpsSync.Erro)"
    } else {
        $ui.txtFpsStatus.Text = 'Medindo FPS em tempo real (Intel PresentMon). Quedas são analisadas automaticamente.'
    }
}

# ---- Diagnóstico das quedas ----
$Hist['Fps'] = New-Object System.Collections.Generic.List[double]
$script:Quedas = New-Object System.Collections.Generic.List[object]
$script:QuedaSeg = 0
$script:UltimaQueda = [DateTime]::MinValue
$script:ClockHist = New-Object System.Collections.Generic.List[double]
$script:Sinais = @{}

$cntNucleos = @()
try {
    $catProc = New-Object Diagnostics.PerformanceCounterCategory('Processor')
    foreach ($n in ($catProc.GetInstanceNames() | Where-Object { $_ -match '^\d+$' } | Sort-Object { [int]$_ })) {
        $cn = New-Object Diagnostics.PerformanceCounter('Processor', '% Processor Time', $n)
        [void]$cn.NextValue()
        $cntNucleos += $cn
    }
} catch {}

# Decide a causa mais provável a partir dos sinais do momento
function Get-Causa($si) {
    if ($si.Dsc -ge 85) {
        return @{ Guia = 'disco'; Causa = 'Disco em 100%'; Detalhe = "Disco ocupado $($si.Dsc)% do tempo: o jogo esperou o disco carregar dados. $($si.FundoTxt)" }
    }
    if ($si.Ram -ge 90) {
        return @{ Guia = 'ram'; Causa = 'Memória RAM cheia'; Detalhe = "RAM em $($si.Ram)%: o Windows passou a usar o disco como memória (bem mais lento)." }
    }
    if ($null -ne $si.VramPct -and $si.VramPct -ge 95) {
        return @{ Guia = 'vram'; Causa = 'Memória de vídeo cheia'; Detalhe = "VRAM em $([int]$si.VramPct)%: as texturas não cabem na placa de vídeo." }
    }
    if ($null -ne $si.GpuT -and ($si.GpuT -ge $Config.GpuTempLimite - 2 -or ($si.Gpu -ge 80 -and $si.ClockQueda -ge 15))) {
        return @{ Guia = 'gpu'; Causa = 'Placa de vídeo esquentando'; Detalhe = "GPU a $($si.GpuT) °C e clock $([int]$si.ClockQueda)% abaixo do normal: ela reduz a velocidade para não superaquecer." }
    }
    if ($si.FundoCpu -ge 15) {
        return @{ Guia = 'cpu'; Causa = 'Programa em segundo plano'; Detalhe = "$($si.FundoNome) usou $($si.FundoCpu)% do processador no momento da queda." }
    }
    if ($si.Cpu -ge 90) {
        return @{ Guia = 'cpu'; Causa = 'Processador no limite'; Detalhe = "CPU em $($si.Cpu)% no total." }
    }
    if ($si.Nucleo -ge 95 -and $si.Gpu -lt 85) {
        return @{ Guia = 'cpu'; Causa = 'Gargalo de processador'; Detalhe = "Um núcleo chegou a $($si.Nucleo)% enquanto a placa de vídeo ficou em $($si.Gpu)%: o jogo depende de um núcleo só." }
    }
    if ($si.Gpu -ge 95) {
        return @{ Guia = 'gpu'; Causa = 'Placa de vídeo no limite'; Detalhe = "GPU em $($si.Gpu)%: a cena ficou mais pesada do que a placa aguenta nessa configuração." }
    }
    @{ Guia = 'outro'; Causa = 'Engasgo do próprio jogo'; Detalhe = 'Nenhum componente estava no limite. Costuma ser carregamento de área nova ou compilação de shaders.' }
}

function Get-Gargalo($si) {
    if ($si.Dsc -ge 85) { return @('Disco em 100%', 'O disco está sobrecarregado — isso causa travadinhas ao carregar.', $Cor.Vermelho) }
    if ($si.Ram -ge 90) { return @('Memória RAM quase cheia', 'Feche programas abertos para o jogo não depender do disco.', $Cor.Ambar) }
    if ($null -ne $si.GpuT -and $si.GpuT -ge $Config.GpuTempLimite - 2) { return @('Placa de vídeo quente', "A $($si.GpuT) °C ela pode reduzir a velocidade sozinha.", $Cor.Ambar) }
    if ($si.Gpu -ge 95) { return @('Placa de vídeo (normal)', 'A GPU está trabalhando no máximo — é o esperado em jogos pesados. Para mais FPS, baixe a qualidade gráfica ou a resolução.', $Cor.Lima) }
    if ($si.Nucleo -ge 95 -and $si.Gpu -lt 85) { return @('Processador', 'Um núcleo está no máximo e a GPU sobrando: o jogo está limitado pelo processador.', $Cor.Ambar) }
    if ($si.Cpu -ge 90) { return @('Processador no limite', 'O processador inteiro está ocupado.', $Cor.Ambar) }
    @('Nada no limite', 'Sobra desempenho: o FPS deve estar limitado pelo próprio jogo (limite de FPS / V-Sync).', $Cor.Lima)
}

function Format-Sinal($txt, $valor, $limite) {
    $txt.Foreground = Get-Pincel $(if ($valor -ge $limite) { $Cor.Vermelho } elseif ($valor -ge $limite - 15) { $Cor.Ambar } else { '#F2F2F5' })
}

$script:FpsRetentar = [DateTime]::MinValue
function Update-Jogos($cpu, $ramUso, $dsc, $g) {
    # Se o medidor parou sozinho (antivírus segurando o arquivo, erro do Windows...), tenta de novo a cada 10 s
    if ($Config.MedirFps -and $script:fpsJob -and -not $FpsSync.Rodando) {
        try { $script:fpsJob.Ps.Dispose() } catch {}
        $script:fpsJob = $null
        $script:FpsRetentar = (Get-Date).AddSeconds(10)
    }
    if ($Config.MedirFps -and -not $script:fpsJob -and (Get-Date) -ge $script:FpsRetentar) { Start-Fps }

    # Sinais do momento
    $nucleo = 0
    foreach ($cn in $cntNucleos) { try { $v = $cn.NextValue(); if ($v -gt $nucleo) { $nucleo = $v } } catch {} }
    $nucleo = [int][Math]::Min(100, [Math]::Round($nucleo))
    $a = $FpsSync.Atual
    $temJogo = $a -and ((Get-Date) - $a.Hora).TotalSeconds -lt 2.5
    $nomeJogo = if ($temJogo) { [IO.Path]::GetFileNameWithoutExtension($a.App) } else { '' }
    $fundo = $script:Processos | Where-Object { $_.Nome -ne $nomeJogo -and $_.Nome -notin 'Idle', 'System', 'Memory Compression', 'dwm' } |
        Sort-Object Cpu -Descending | Select-Object -First 1
    $si = @{
        Cpu = $cpu; Nucleo = $nucleo; Ram = $ramUso; Dsc = $dsc
        Gpu = $(if ($g) { [int]$g.Uso } else { 0 }); GpuT = $(if ($g -and $null -ne $g.Temp) { [int]$g.Temp } else { $null })
        VramPct = $(if ($g -and $g.VramTotal) { $g.VramUsada / $g.VramTotal * 100 } else { $null })
        FundoNome = $fundo.Nome; FundoCpu = [double]$fundo.Cpu
        FundoTxt = $(if ($fundo -and $fundo.Cpu -ge 5) { "Mais pesado em segundo plano: $($fundo.Nome)." } else { '' })
        ClockQueda = 0
    }
    if ($g -and $g.Clock) {
        $script:ClockHist.Add([double]$g.Clock); if ($script:ClockHist.Count -gt 60) { $script:ClockHist.RemoveAt(0) }
        $maxClock = ($script:ClockHist | Measure-Object -Maximum).Maximum
        if ($maxClock -gt 0) { $si.ClockQueda = [Math]::Max(0, (1 - $g.Clock / $maxClock) * 100) }
    }
    $script:Sinais = $si

    # FPS e detecção de queda (roda sempre, mesmo com a aba fechada)
    if ($temJogo) {
        $fps = [Math]::Round($a.Fps)
        Add-Historico 'Fps' $fps
        $base = @($Hist.Fps | Select-Object -Last 30 | Sort-Object)
        $mediana = if ($base.Count -ge 10) { $base[[int]($base.Count / 2)] } else { 0 }
        if ($mediana -ge 25 -and $fps -lt $mediana * 0.75) { $script:QuedaSeg++ } else { $script:QuedaSeg = 0 }
        if ($script:QuedaSeg -eq 2 -and ((Get-Date) - $script:UltimaQueda).TotalSeconds -gt 8) {
            $script:UltimaQueda = Get-Date
            $causa = Get-Causa $si
            $script:Quedas.Insert(0, [pscustomobject]@{
                Hora = (Get-Date).ToString('HH:mm:ss'); Causa = $causa.Causa; Guia = $causa.Guia; Detalhe = $causa.Detalhe
                Queda = " · $nomeJogo caiu de $mediana para $fps FPS"; Indice = $Hist.Fps.Count
            })
            if ($script:Quedas.Count -gt 30) { $script:Quedas.RemoveAt(30) }
            $ui.icQuedas.ItemsSource = $null; $ui.icQuedas.ItemsSource = $script:Quedas
            $ui.txtSemQuedas.Visibility = 'Collapsed'
            try { Add-Content -Path (Join-Path $ConfigDir 'quedas.log') -Value ("{0:dd/MM HH:mm:ss}  {1}  {2}->{3} FPS  {4}: {5}" -f (Get-Date), $nomeJogo, $mediana, $fps, $causa.Causa, $causa.Detalhe) -Encoding UTF8 } catch {}
            if ($script:Sessao) { $script:Sessao.Quedas++; $script:Sessao.Causas += $causa.Causa }
        }
    }

    Update-SessaoJogo $temJogo $nomeJogo $a $si
    Update-Ping
    Update-Overlay $temJogo $a $si

    if ($ui.tabs.SelectedIndex -ne 8) { return }

    # Tela
    Update-EstadoFps
    if ($temJogo) {
        $ui.txtJogoNome.Text = $nomeJogo
        $ui.txtJogoSub.Text = "PID $($a.Pid) · medindo há $(@($Hist.Fps).Count) s"
        $ui.txtFpsModo.Text = if ($a.Modo -like 'Hardware*') { 'Tela cheia otimizada' } elseif ($a.Modo) { 'Janela (composto)' } else { 'ativo' }
        $ui.imgJogo.Source = Get-IconeProcesso $nomeJogo
        $ui.icJogoGlifo.Visibility = if ($ui.imgJogo.Source) { 'Collapsed' } else { 'Visible' }
        $ui.txtFps.Text = [string][Math]::Round($a.Fps)
        $ultimos = @($Hist.Fps | Select-Object -Last 60)
        $ui.txtFpsMedia.Text = '{0:N0}' -f ($ultimos | Measure-Object -Average).Average
        $ui.txtFpsLow.Text = '{0:N0}' -f $a.Low1
        $ui.txtFrameMs.Text = '{0:N1} ms' -f $a.Ms
        $escala = [Math]::Max(60, (($Hist.Fps | Measure-Object -Maximum).Maximum) * 1.15)
        $ui.txtFpsEscala.Text = '{0:N0} FPS' -f $escala
        $pts = Update-Serie $ui.cvFps $ui.lnFps $ui.agFps $Hist.Fps $escala
        # Marcas das quedas sobre o gráfico
        $ui.cvQuedas.Children.Clear()
        if ($pts) {
            $inicio = $Hist.Fps.Count - @($pts).Count
            foreach ($q in $script:Quedas) {
                $i = $q.Indice - 1 - $inicio
                if ($i -lt 0 -or $i -ge @($pts).Count) { continue }
                $e = New-Object Windows.Shapes.Ellipse
                $e.Width = 10; $e.Height = 10; $e.Fill = Get-Pincel $Cor.Vermelho
                [Windows.Controls.Canvas]::SetLeft($e, $pts[$i].X - 5); [Windows.Controls.Canvas]::SetTop($e, $pts[$i].Y - 5)
                [void]$ui.cvQuedas.Children.Add($e)
            }
        }
    } else {
        $ui.txtJogoNome.Text = 'Nenhum jogo detectado'
        $ui.txtJogoSub.Text = if ($FpsSync.Rodando) { 'Abra um jogo — o Monitor detecta sozinho' } else { 'Medição de FPS parada' }
        $ui.txtFpsModo.Text = 'aguardando jogo'
        $ui.txtFps.Text = '--'; $ui.txtFpsMedia.Text = '--'; $ui.txtFpsLow.Text = '--'; $ui.txtFrameMs.Text = '--'
        $ui.imgJogo.Source = $null; $ui.icJogoGlifo.Visibility = 'Visible'
    }
    $gg = Get-Gargalo $si
    $ui.txtGargalo.Text = $gg[0]; $ui.txtGargalo.Foreground = Get-Pincel $gg[2]; $ui.txtGargaloDesc.Text = $gg[1]
    $ui.txtSCpu.Text = "$cpu%"; Format-Sinal $ui.txtSCpu $cpu 90
    $ui.txtSNucleo.Text = "$nucleo%"; Format-Sinal $ui.txtSNucleo $nucleo 95
    $ui.txtSGpu.Text = "$($si.Gpu)%"; Format-Sinal $ui.txtSGpu $si.Gpu 101
    $ui.txtSGpuT.Text = if ($null -ne $si.GpuT) { "$($si.GpuT) °C" } else { 'N/D' }
    if ($null -ne $si.GpuT) { Format-Sinal $ui.txtSGpuT $si.GpuT $Config.GpuTempLimite }
    $ui.txtSVram.Text = if ($null -ne $si.VramPct) { '{0:N0}%' -f $si.VramPct } else { 'N/D' }
    if ($null -ne $si.VramPct) { Format-Sinal $ui.txtSVram $si.VramPct 95 }
    $ui.txtSRam.Text = "$ramUso%"; Format-Sinal $ui.txtSRam $ramUso 90
    $ui.txtSDsc.Text = "$dsc%"; Format-Sinal $ui.txtSDsc $dsc 85
    $ui.txtSFundo.Text = if ($fundo) { '{0} ({1:N0}%)' -f $fundo.Nome, $fundo.Cpu } else { '--' }
}

# ---- Guia de soluções ----
function Get-Cmd($titulo) { $Comandos | Where-Object { $_.Titulo -eq $titulo } | Select-Object -First 1 }
$CmdTopDisco = @{ Titulo = 'Ver o que está usando o disco'; Admin = $false; Risco = 'Seguro'; Glifo = 0xEDA2
    Descricao = 'Lista os 10 programas que mais leem/gravam no disco agora.'
    Script = 'Get-CimInstance Win32_PerfFormattedData_PerfProc_Process | Where-Object { $_.Name -notin "_Total","Idle" } | Sort-Object IODataBytesPersec -Descending | Select-Object -First 10 @{n="Processo";e={$_.Name}}, @{n="MB/s";e={[math]::Round($_.IODataBytesPersec/1MB,1)}} | Format-Table -AutoSize' }
$CmdTopCpu = @{ Titulo = 'Ver o que está usando o processador'; Admin = $false; Risco = 'Seguro'; Glifo = 0xE950
    Descricao = 'Lista os 10 programas que mais usam a CPU neste momento.'
    Script = '$n = [Environment]::ProcessorCount; Get-CimInstance Win32_PerfFormattedData_PerfProc_Process | Where-Object { $_.Name -notin "_Total","Idle" } | Sort-Object PercentProcessorTime -Descending | Select-Object -First 10 @{n="Processo";e={$_.Name}}, @{n="CPU %";e={[math]::Round($_.PercentProcessorTime/$n,1)}} | Format-Table -AutoSize' }
$CmdTopRam = @{ Titulo = 'Ver o que está usando a memória'; Admin = $false; Risco = 'Seguro'; Glifo = 0xE964
    Descricao = 'Lista os 10 programas que mais ocupam a memória RAM.'
    Script = 'Get-Process | Group-Object ProcessName | ForEach-Object { [pscustomobject]@{ Programa = $_.Name; Copias = $_.Count; "RAM (MB)" = [math]::Round(($_.Group | Measure-Object WorkingSet64 -Sum).Sum/1MB) } } | Sort-Object "RAM (MB)" -Descending | Select-Object -First 10 | Format-Table -AutoSize' }
$CmdIndexacao = @{ Titulo = 'Pausar a indexação de arquivos'; Admin = $true; Risco = 'Seguro'; Glifo = 0xE721
    Descricao = 'Para o Windows Search até o próximo reinício. Ele costuma ler o disco sem parar depois de atualizações.'
    Script = 'Stop-Service WSearch -Force; Write-Host "Indexação pausada até reiniciar o PC." -ForegroundColor Green' }
$CmdSysMain = @{ Titulo = 'Desativar o SysMain (Superfetch)'; Admin = $true; Risco = 'Moderado'; Glifo = 0xEDA2
    Descricao = 'Serviço que pré-carrega programas e é a causa clássica de disco em 100% em HD. Em SSD a diferença é pequena. Reversível.'
    Script = 'Stop-Service SysMain -Force; Set-Service SysMain -StartupType Disabled; Write-Host "SysMain desativado. Para reativar: Set-Service SysMain -StartupType Automatic; Start-Service SysMain" -ForegroundColor Green' }
$CmdGpuDiag = @{ Titulo = 'Diagnóstico da placa de vídeo'; Admin = $false; Risco = 'Seguro'; Glifo = 0xE7F4
    Descricao = 'Mostra temperatura, consumo, clocks e o MOTIVO de a placa estar se limitando (calor, energia ou sem carga). Placas NVIDIA.'
    Script = 'if (Get-Command nvidia-smi -ErrorAction SilentlyContinue) { nvidia-smi -q -d PERFORMANCE,TEMPERATURE,POWER,CLOCK; Write-Host ""; Write-Host "Em Clocks Event Reasons, Active = é isso que está limitando a placa." -ForegroundColor Yellow } else { Write-Host "nvidia-smi não encontrado (placa não é NVIDIA ou driver sem a ferramenta)." -ForegroundColor Yellow }' }
$CmdShaders = @{ Titulo = 'Limpar cache de shaders'; Admin = $false; Risco = 'Seguro'; Glifo = 0xE74D
    Descricao = 'Apaga o cache de shaders do DirectX/NVIDIA/AMD. Resolve travadinhas depois de atualizar driver. O primeiro carregamento do jogo fica um pouco mais lento.'
    Script = 'Limpar-Pasta "$env:LOCALAPPDATA\D3DSCache"; Limpar-Pasta "$env:LOCALAPPDATA\NVIDIA\DXCache"; Limpar-Pasta "$env:LOCALAPPDATA\NVIDIA\GLCache"; Limpar-Pasta "$env:LOCALAPPDATA\AMD\DxCache"' }
$CmdModoJogo = @{ Titulo = 'Abrir Modo de Jogo do Windows'; Admin = $false; Risco = 'Seguro'; Glifo = 0xE7FC
    Descricao = 'Confira se o Modo de Jogo está ATIVADO: ele dá prioridade ao jogo e segura atualizações durante a partida.'
    Script = 'Start-Process "ms-settings:gaming-gamemode"' }
$CmdGraficos = @{ Titulo = 'Abrir configurações de gráficos'; Admin = $false; Risco = 'Seguro'; Glifo = 0xE7F4
    Descricao = 'Ative o "Agendamento de GPU acelerado por hardware" e escolha "Alto desempenho" para o jogo.'
    Script = 'Start-Process "ms-settings:display-advancedgraphics"' }

$CmdWifi = @{ Titulo = 'Ver o sinal do Wi-Fi'; Admin = $false; Risco = 'Seguro'; Glifo = 0xEC05
    Descricao = 'Mostra a força do sinal (abaixo de 70% já pode causar lag), a velocidade e o canal do Wi-Fi.'
    Script = 'netsh wlan show interfaces; Write-Host ""; Write-Host "Veja a linha Sinal: abaixo de 70% vale aproximar do roteador ou usar cabo." -ForegroundColor Yellow' }
$CmdRota = @{ Titulo = 'Rastrear o caminho até o servidor'; Admin = $false; Risco = 'Seguro'; Glifo = 0xEC05
    Descricao = 'Mostra cada "salto" da sua casa até o servidor medido e onde o tempo de resposta sobe.'
    Script = "tracert -d -h 15 $($Config.PingAlvo); Write-Host ''; Write-Host 'Se o tempo já é alto nas primeiras linhas, o problema é a rede da sua casa; se sobe no meio, é o provedor.' -ForegroundColor Yellow" }

$Guias = [ordered]@{
    disco = @{ Titulo = 'Disco (HD/SSD) em 100%'; Cor = $Cor.Vermelho
        Texto = 'Causa travadas de vários segundos e lentidão geral. Os culpados mais comuns: Windows Update baixando, antivírus fazendo verificação, indexação de arquivos, Steam/Epic baixando jogo e, em HD, o serviço SysMain. Em HD velho também pode ser sinal de desgaste (veja a aba Vida útil).'
        Comandos = @($CmdTopDisco, $CmdIndexacao, $CmdSysMain, (Get-Cmd 'Verificar o disco (somente leitura)'), (Get-Cmd 'Otimizar unidades (TRIM / desfragmentar)')) }
    cpu = @{ Titulo = 'Processador (CPU) em 100%'; Cor = $Cor.Ambar
        Texto = 'O jogo disputa o processador com outros programas. Feche navegador, Discord com tela compartilhada e gravação de vídeo. Se a CPU fica alta mesmo sem nada aberto, pode ser vírus minerador — rode a verificação do Defender. Em jogos que usam um núcleo só, limitar o FPS no jogo deixa tudo mais estável.'
        Comandos = @($CmdTopCpu, (Get-Cmd 'Plano Desempenho Máximo'), (Get-Cmd 'Gerenciar programas da inicialização'), $AcoesSeguranca['scan']) }
    gpu = @{ Titulo = 'Placa de vídeo em 100% ou esquentando'; Cor = $Cor.Ambar
        Texto = 'GPU em 100% durante o jogo é NORMAL — significa que ela está sendo bem aproveitada. O problema é quando ela esquenta demais e reduz o clock sozinha (thermal throttling): limpe a poeira, melhore o fluxo de ar do gabinete e confira se as ventoinhas giram. Para mais FPS, baixe sombras, resolução ou use DLSS/FSR.'
        Comandos = @($CmdGpuDiag, $CmdShaders, $CmdModoJogo, $CmdGraficos) }
    ram = @{ Titulo = 'Memória RAM cheia'; Cor = $Cor.Ambar
        Texto = 'Quando a RAM enche, o Windows usa o disco como memória extra e o jogo engasga. Navegadores com muitas abas são os maiores consumidores. Feche o que não estiver usando antes de jogar.'
        Comandos = @($CmdTopRam, (Get-Cmd 'Reiniciar o Explorer'), (Get-Cmd 'Diagnóstico de memória RAM')) }
    vram = @{ Titulo = 'Memória de vídeo (VRAM) cheia'; Cor = $Cor.Ambar
        Texto = 'As texturas do jogo não cabem na placa de vídeo e passam a ir e voltar pela RAM, causando travadas. Baixe a qualidade das TEXTURAS e da resolução; feche vídeos/streams abertos no navegador enquanto joga.'
        Comandos = @($CmdShaders, $CmdGraficos) }
    rede = @{ Titulo = 'Lag na internet (ping alto ou perda de pacotes)'; Cor = $Cor.Ambar
        Texto = 'Lag não é o mesmo que FPS baixo: o jogo roda liso, mas os outros jogadores "teleportam". Se o ping até o ROTEADOR já está alto ou perde pacotes, o problema é o Wi-Fi/cabo dentro de casa — use cabo de rede ou aproxime-se do roteador. Se o roteador está bem e só a internet falha, é o provedor ou alguém baixando/assistindo vídeo na mesma rede.'
        Comandos = @($CmdWifi, $CmdRota, (Get-Cmd 'Limpar cache de DNS'), (Get-Cmd 'Resetar a rede (Winsock e TCP/IP)')) }
    outro = @{ Titulo = 'Jogo engasgando sem motivo claro'; Cor = $Cor.LimaMedia
        Texto = 'Se nenhum componente está no limite, costuma ser o próprio jogo: compilação de shaders nas primeiras partidas, carregamento de área nova ou arquivos corrompidos. Verifique os arquivos do jogo na Steam/Epic (Propriedades > Verificar integridade), atualize o driver de vídeo e limpe o cache de shaders.'
        Comandos = @($CmdShaders, $CmdModoJogo, (Get-Cmd 'Reparar arquivos do sistema (SFC)')) }
}

$CardsGuia = @{}
foreach ($chave in $Guias.Keys) {
    $gd = $Guias[$chave]
    $card = New-Object Windows.Controls.Border
    $card.Style = $window.FindResource('Card')
    $card.Padding = '22,18'
    $sp = New-Object Windows.Controls.StackPanel
    $cab = New-Object Windows.Controls.StackPanel
    $cab.Orientation = 'Horizontal'
    $ponto = New-Object Windows.Shapes.Ellipse
    $ponto.Width = 10; $ponto.Height = 10; $ponto.Fill = Get-Pincel $gd.Cor; $ponto.Margin = '0,0,10,0'; $ponto.VerticalAlignment = 'Center'
    $tt = New-Object Windows.Controls.TextBlock
    $tt.Text = $gd.Titulo; $tt.FontSize = 16; $tt.FontWeight = 'SemiBold'
    [void]$cab.Children.Add($ponto); [void]$cab.Children.Add($tt)
    $tx = New-Object Windows.Controls.TextBlock
    $tx.Text = $gd.Texto; $tx.TextWrapping = 'Wrap'; $tx.Foreground = Get-Pincel '#C8C8CE'; $tx.Margin = '0,8,0,10'
    [void]$sp.Children.Add($cab); [void]$sp.Children.Add($tx)
    foreach ($cmd in $gd.Comandos | Where-Object { $_ }) {
        $linha = New-Object Windows.Controls.DockPanel
        $linha.Margin = '0,4'
        $btn = New-Object Windows.Controls.Button
        $btn.Content = 'Executar'; $btn.Tag = $cmd; $btn.Padding = '14,6'; $btn.FontSize = 12; $btn.VerticalAlignment = 'Center'; $btn.Margin = '12,0,0,0'
        $btn.add_Click({ param($s, $e) Invoke-Comando $s.Tag })
        [Windows.Controls.DockPanel]::SetDock($btn, 'Right')
        $info = New-Object Windows.Controls.StackPanel
        $l1 = New-Object Windows.Controls.TextBlock
        $l1.Text = $cmd.Titulo + $(if ($cmd.Admin) { '  ·  admin' } else { '' }); $l1.FontWeight = 'SemiBold'
        $l2 = New-Object Windows.Controls.TextBlock
        $l2.Text = $cmd.Descricao; $l2.TextWrapping = 'Wrap'; $l2.FontSize = 12; $l2.Foreground = Get-Pincel $Cor.Suave
        [void]$info.Children.Add($l1); [void]$info.Children.Add($l2)
        [void]$linha.Children.Add($btn); [void]$linha.Children.Add($info)
        [void]$sp.Children.Add($linha)
    }
    $card.Child = $sp
    [void]$ui.spGuias.Children.Add($card)
    $CardsGuia[$chave] = $card
}

# Leva até a solução e destaca o card por alguns segundos
$script:CardDestacado = $null
$timerDestaque = New-Object Windows.Threading.DispatcherTimer
$timerDestaque.Interval = [TimeSpan]::FromSeconds(3)
$timerDestaque.add_Tick({ $timerDestaque.Stop(); if ($script:CardDestacado) { $script:CardDestacado.BorderBrush = Get-Pincel '#1F1F23' } })
function Show-Guia($chave) {
    Select-Aba 8
    $card = $CardsGuia[$chave]
    if (-not $card) { return }
    $window.Dispatcher.BeginInvoke([Action]{ $card.BringIntoView() }, 'Background') | Out-Null
    if ($script:CardDestacado) { $script:CardDestacado.BorderBrush = Get-Pincel '#1F1F23' }
    $card.BorderBrush = Get-Pincel $Cor.Lima
    $script:CardDestacado = $card
    $timerDestaque.Stop(); $timerDestaque.Start()
}

$ui.icQuedas.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent, [Windows.RoutedEventHandler]{
    param($s, $e)
    if ($e.OriginalSource -is [Windows.Controls.Button]) { Show-Guia $e.OriginalSource.Tag }
})
$ui.btnLimparQuedas.add_Click({ $script:Quedas.Clear(); $ui.icQuedas.ItemsSource = $null; $ui.txtSemQuedas.Visibility = 'Visible' })
$ui.btnFpsAdmin.add_Click({ $ui.btnAdmin.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Primitives.ButtonBase]::ClickEvent))) })
$ui.chkFps.IsChecked = [bool]$Config.MedirFps
$ui.chkFps.add_Click({
    $Config.MedirFps = [bool]$ui.chkFps.IsChecked
    Save-Config
    if ($Config.MedirFps) { Start-Fps } else { Stop-Fps }
    Update-EstadoFps
})
$ui.btnFpsBaixar.add_Click({
    $r = [Windows.MessageBox]::Show("Baixar o Intel PresentMon (1 MB) do GitHub oficial da Intel?`n`n$PresentMonUrl`n`nEle é gratuito, de código aberto e assinado pela Intel. Fica na pasta 'ferramentas' do Monitor.", $AppNome, 'YesNo', 'Question')
    if ($r -ne 'Yes') { return }
    try {
        $ui.txtFpsStatus.Text = 'Baixando...'
        New-Item -ItemType Directory -Path (Split-Path $PresentMonPath) -Force | Out-Null
        $tmp = "$PresentMonPath.baixando"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        (New-Object Net.WebClient).DownloadFile($PresentMonUrl, $tmp)
        # Só aceita se for realmente assinado pela Intel
        $sig = Get-AuthenticodeSignature $tmp
        if ($sig.Status -ne 'Valid' -or $sig.SignerCertificate.Subject -notmatch 'O=Intel Corporation') {
            Remove-Item $tmp -Force
            throw 'o arquivo baixado não tem a assinatura digital da Intel e foi descartado.'
        }
        Move-Item $tmp $PresentMonPath -Force
        if ($Config.MedirFps) { Start-Fps }
    } catch {
        [Windows.MessageBox]::Show("Não foi possível baixar: $($_.Exception.Message)", $AppNome) | Out-Null
    }
    Update-EstadoFps
})
if ($Config.MedirFps) { Start-Fps }

# ---- FPS na tela (overlay) ----
# Janelinha transparente, sempre no topo, que deixa o clique passar para o jogo.
# Aparece em jogos em janela/tela cheia sem borda; na tela cheia exclusiva o próprio jogo a cobre.
Add-Type -Namespace MonitorPC -Name Janela -MemberDefinition @'
[DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr h, int i);
[DllImport("user32.dll")] public static extern int SetWindowLong(IntPtr h, int i, int v);
[DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr depois, int x, int y, int cx, int cy, uint f);
'@
$XamlOverlay = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent" Topmost="True" ShowInTaskbar="False"
        ShowActivated="False" SizeToContent="WidthAndHeight" ResizeMode="NoResize" Focusable="False" IsHitTestVisible="False"
        FontFamily="Segoe UI" TextOptions.TextFormattingMode="Display">
  <Border CornerRadius="10" Background="#C80B0B0D" BorderBrush="#66A6F23A" BorderThickness="1" Padding="12,7">
    <StackPanel>
      <StackPanel Orientation="Horizontal">
        <TextBlock x:Name="ovFps" Text="--" FontSize="26" FontWeight="Bold" Foreground="#A6F23A"/>
        <TextBlock Text="FPS" FontSize="11" Foreground="#9A9AA2" VerticalAlignment="Bottom" Margin="5,0,0,6"/>
        <TextBlock x:Name="ovLow" FontSize="11" Foreground="#C8C8CE" VerticalAlignment="Bottom" Margin="10,0,0,6"/>
      </StackPanel>
      <TextBlock x:Name="ovGpu" FontFamily="Consolas" FontSize="12" Foreground="#E6E6EB"/>
      <TextBlock x:Name="ovCpu" FontFamily="Consolas" FontSize="12" Foreground="#E6E6EB"/>
      <TextBlock x:Name="ovPing" FontFamily="Consolas" FontSize="12" Foreground="#E6E6EB"/>
    </StackPanel>
  </Border>
</Window>
'@
foreach ($k in $Paleta.Keys) { $XamlOverlay = [regex]::Replace($XamlOverlay, "(?i)(#(?:[0-9a-f]{2})?)$k(?![0-9a-f])", "`${1}$($Paleta[$k])") }
$Overlay = [Windows.Markup.XamlReader]::Parse($XamlOverlay)
$ov = @{}
foreach ($n in 'ovFps', 'ovLow', 'ovGpu', 'ovCpu', 'ovPing') { $ov[$n] = $Overlay.FindName($n) }
$Overlay.add_SourceInitialized({
    $h = (New-Object Windows.Interop.WindowInteropHelper($Overlay)).Handle
    # WS_EX_TRANSPARENT (clique atravessa) + WS_EX_TOOLWINDOW (fora do Alt+Tab) + WS_EX_NOACTIVATE (não rouba o foco do jogo)
    [void][MonitorPC.Janela]::SetWindowLong($h, -20, ([MonitorPC.Janela]::GetWindowLong($h, -20) -bor 0x20 -bor 0x80 -bor 0x08000000))
})
function Set-PosicaoOverlay {
    $w = [Windows.SystemParameters]::PrimaryScreenWidth; $hT = [Windows.SystemParameters]::PrimaryScreenHeight
    $m = 18
    $lw = if ($Overlay.ActualWidth) { $Overlay.ActualWidth } else { 170 }
    $lh = if ($Overlay.ActualHeight) { $Overlay.ActualHeight } else { 90 }
    switch ($Config.OverlayCanto) {
        'SuperiorDireito'  { $Overlay.Left = $w - $lw - $m; $Overlay.Top = $m }
        'InferiorEsquerdo' { $Overlay.Left = $m; $Overlay.Top = $hT - $lh - $m - 40 }
        'InferiorDireito'  { $Overlay.Left = $w - $lw - $m; $Overlay.Top = $hT - $lh - $m - 40 }
        default            { $Overlay.Left = $m; $Overlay.Top = $m }
    }
}
$Overlay.add_SizeChanged({ Set-PosicaoOverlay })

function Update-Overlay($temJogo, $a, $si) {
    if (-not ($Config.Overlay -and $temJogo)) {
        if ($Overlay.IsVisible) { $Overlay.Hide() }
        return
    }
    $ov.ovFps.Text = [string][Math]::Round($a.Fps)
    $ov.ovFps.Foreground = Get-Pincel $(if ($script:QuedaSeg -gt 0) { $Cor.Ambar } else { $Cor.Lima })
    $ov.ovLow.Text = '1% {0:N0}' -f $a.Low1
    $ov.ovGpu.Text = 'GPU {0,3}%' -f $si.Gpu + $(if ($null -ne $si.GpuT) { '  {0} °C' -f $si.GpuT } else { '' })
    $ov.ovCpu.Text = 'CPU {0,3}%  RAM {1}%' -f $si.Cpu, $si.Ram
    $ultPing = $PingSync.Ultimo
    $ov.ovPing.Text = if ($ultPing -and $null -ne $ultPing.Internet) { 'Ping {0} ms' -f [int]$ultPing.Internet } else { 'Ping --' }
    if (-not $Overlay.IsVisible) { $Overlay.Show(); Set-PosicaoOverlay }
    # Reafirma "sempre no topo" (alguns jogos tomam a frente ao trocar de tela)
    $h = (New-Object Windows.Interop.WindowInteropHelper($Overlay)).Handle
    [void][MonitorPC.Janela]::SetWindowPos($h, [IntPtr](-1), 0, 0, 0, 0, 0x0001 -bor 0x0002 -bor 0x0010)
}

$ui.chkOverlay.IsChecked = [bool]$Config.Overlay
$ui.chkOverlay.add_Click({ $Config.Overlay = [bool]$ui.chkOverlay.IsChecked; Save-Config })
foreach ($rb in $ui.rbOvSE, $ui.rbOvSD, $ui.rbOvIE, $ui.rbOvID) {
    $rb.IsChecked = $rb.Tag -eq $Config.OverlayCanto
    $rb.add_Checked({ param($s, $e) $Config.OverlayCanto = $s.Tag; Save-Config; if ($Overlay.IsVisible) { Set-PosicaoOverlay } })
}

# ---- Ping em segundo plano: internet (alvo configurável) e roteador ----
$PingSync = [hashtable]::Synchronized(@{ Rodando = $true; Alvo = $Config.PingAlvo; Ultimo = $null; Roteador = '' })
$pingPs = [powershell]::Create()
[void]$pingPs.AddScript({
    param($s)
    $ping = New-Object Net.NetworkInformation.Ping
    $seq = 0; $proxRota = [DateTime]::MinValue
    while ($s.Rodando) {
        $inicio = Get-Date
        if ($inicio -ge $proxRota) {
            try {
                $s.Roteador = (Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop | Sort-Object RouteMetric | Select-Object -First 1).NextHop
            } catch { $s.Roteador = '' }
            $proxRota = $inicio.AddMinutes(1)
        }
        $msNet = $null; $msRot = $null
        try { $r = $ping.Send($s.Alvo, 1000); if ($r.Status -eq 'Success') { $msNet = [double]$r.RoundtripTime } } catch {}
        if ($s.Roteador -and $s.Roteador -ne '0.0.0.0') {
            try { $r = $ping.Send($s.Roteador, 1000); if ($r.Status -eq 'Success') { $msRot = [double]$r.RoundtripTime } } catch {}
        }
        $seq++
        $s.Ultimo = @{ Seq = $seq; Internet = $msNet; Roteador = $msRot }
        $resto = 1000 - ((Get-Date) - $inicio).TotalMilliseconds
        if ($resto -gt 0) { Start-Sleep -Milliseconds ([int]$resto) }
    }
}).AddArgument($PingSync)
$pingHandle = $pingPs.BeginInvoke()

$script:PingSeq = 0
$script:PingNet = New-Object System.Collections.Generic.List[object]   # $null = pacote perdido
$script:PingRot = New-Object System.Collections.Generic.List[object]
$Hist['Ping'] = New-Object System.Collections.Generic.List[double]

function Get-StatsPing($lista) {
    $ok = @($lista | Where-Object { $null -ne $_ })
    $st = @{ Media = $null; Jitter = $null; Perda = 0.0; N = $lista.Count }
    if ($lista.Count) { $st.Perda = ($lista.Count - $ok.Count) / $lista.Count * 100 }
    if ($ok.Count) {
        $st.Media = ($ok | Measure-Object -Average).Average
        if ($ok.Count -gt 1) {
            $dif = 0.0
            for ($i = 1; $i -lt $ok.Count; $i++) { $dif += [Math]::Abs($ok[$i] - $ok[$i - 1]) }
            $st.Jitter = $dif / ($ok.Count - 1)
        }
    }
    $st
}

function Update-Ping {
    $u = $PingSync.Ultimo
    if (-not $u -or $u.Seq -eq $script:PingSeq) { return }
    $script:PingSeq = $u.Seq
    foreach ($par in @(@($script:PingNet, $u.Internet), @($script:PingRot, $u.Roteador))) {
        $par[0].Add($par[1]); if ($par[0].Count -gt 60) { $par[0].RemoveAt(0) }
    }
    if ($script:Sessao) { $script:Sessao.Pings += , $u.Internet }
    $net = Get-StatsPing $script:PingNet
    $rot = Get-StatsPing $script:PingRot
    Add-Historico 'Ping' $(if ($null -ne $u.Internet) { $u.Internet } else { -1 })   # -1 marca perda no gráfico

    if ($ui.tabs.SelectedIndex -ne 8) { return }
    $ui.txtPing.Text = if ($null -ne $u.Internet) { [string][int]$u.Internet } else { 'perdido' }
    $ui.txtJitter.Text = if ($null -ne $net.Jitter) { '{0:N0} ms' -f $net.Jitter } else { '--' }
    $ui.txtPerda.Text = '{0:N0}%' -f $net.Perda
    $ui.txtPerda.Foreground = Get-Pincel $(if ($net.Perda -ge 2) { $Cor.Vermelho } elseif ($net.Perda -gt 0) { $Cor.Ambar } else { '#F2F2F5' })
    $ui.txtPingRot.Text = if ($null -ne $rot.Media) { '{0:N0} ms' -f $rot.Media } elseif ($PingSync.Roteador) { '--' } else { 'N/D' }
    $valores = New-Object System.Collections.Generic.List[double]
    foreach ($v in $Hist.Ping) { $valores.Add([Math]::Max(0, $v)) }
    $escala = [Math]::Max(50, (($valores | Measure-Object -Maximum).Maximum) * 1.2)
    $ui.txtPingEscala.Text = "Últimos 2 min · escala 0–$([int]$escala) ms · pontos vermelhos = pacote perdido · alvo $($PingSync.Alvo)"
    $pts = Update-Serie $ui.cvPing $ui.lnPing $ui.agPing $valores $escala
    $ui.cvPerdas.Children.Clear()
    if ($pts) {
        $desl = $Hist.Ping.Count - @($pts).Count
        for ($i = 0; $i -lt @($pts).Count; $i++) {
            if ($Hist.Ping[$desl + $i] -ge 0) { continue }
            $e = New-Object Windows.Shapes.Ellipse
            $e.Width = 7; $e.Height = 7; $e.Fill = Get-Pincel $Cor.Vermelho
            [Windows.Controls.Canvas]::SetLeft($e, $pts[$i].X - 3.5); [Windows.Controls.Canvas]::SetTop($e, $ui.cvPing.ActualHeight - 8)
            [void]$ui.cvPerdas.Children.Add($e)
        }
    }
    # Onde está o problema: dentro de casa (roteador) ou na internet?
    if ($net.N -lt 5) {
        $diag = @('Medindo...', '', '#F2F2F5')
    } elseif ($rot.Perda -ge 2 -or ($null -ne $rot.Media -and $rot.Media -gt 15)) {
        $diag = @('Instabilidade na rede de casa', "O próprio roteador está respondendo mal ($('{0:N0}' -f $rot.Media) ms, $('{0:N0}' -f $rot.Perda)% perdidos): Wi-Fi fraco ou interferência. Use cabo ou aproxime-se do roteador.", $Cor.Vermelho)
    } elseif ($net.Perda -ge 2) {
        $diag = @('Perda de pacotes na internet', 'O roteador está bem, mas pacotes se perdem no caminho até o servidor — causa "teleporte" nos jogos. Costuma ser o provedor.', $Cor.Vermelho)
    } elseif ($null -ne $net.Jitter -and $net.Jitter -gt 20) {
        $diag = @('Ping instável', 'O tempo de resposta varia muito. Alguém pode estar baixando ou vendo vídeo na mesma rede.', $Cor.Ambar)
    } elseif ($null -ne $net.Media -and $net.Media -gt 100) {
        $diag = @('Ping alto', 'A conexão é estável, mas o servidor está longe ou o provedor é lento. Escolha um servidor do jogo mais próximo.', $Cor.Ambar)
    } else {
        $diag = @('Conexão boa', 'Sem perdas e com ping estável. Se o jogo travar, a causa é o PC, não a internet.', $Cor.Lima)
    }
    $ui.txtPingDiag.Text = $diag[0]; $ui.txtPingDiag.Foreground = Get-Pincel $diag[2]; $ui.txtPingDiagDesc.Text = $diag[1]
}

$ui.txtPingAlvo.Text = $Config.PingAlvo
$ui.txtPingAlvo.add_LostFocus({
    $alvo = $ui.txtPingAlvo.Text.Trim()
    if (-not $alvo) { $alvo = '1.1.1.1'; $ui.txtPingAlvo.Text = $alvo }
    if ($alvo -ne $Config.PingAlvo) {
        $Config.PingAlvo = $alvo; Save-Config
        $PingSync.Alvo = $alvo
        $script:PingNet.Clear(); $Hist.Ping.Clear()
    }
})

# ---- Modo Jogo automático ----
$script:ModoAplicado = $null

function Enable-ModoJogo($jogo) {
    if ($script:ModoAplicado) { return }
    $m = @{ Jogo = $jogo; Desde = Get-Date; Plano = $null; Indexacao = $false; Fechados = @(); Feito = @() }
    if ($Config.ModoJogoPlano) {
        try {
            $atual = [regex]::Match("$(& powercfg /getactivescheme)", '[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}').Value
            if ($atual) { & powercfg /setactive SCHEME_MIN | Out-Null; $m.Plano = $atual; $m.Feito += 'energia em alto desempenho' }
        } catch {}
    }
    if ($Config.ModoJogoIndexacao -and $EhAdmin) {
        try {
            if ((Get-Service WSearch -ErrorAction Stop).Status -eq 'Running') {
                Start-Process sc.exe -ArgumentList 'stop WSearch' -WindowStyle Hidden
                $m.Indexacao = $true; $m.Feito += 'indexação pausada'
            }
        } catch {}
    }
    foreach ($nome in @($Config.ModoJogoFechar)) {
        if (-not $nome -or $nome -eq $jogo) { continue }
        $procs = @(Get-Process -Name $nome -ErrorAction SilentlyContinue)
        if (-not $procs.Count) { continue }
        $caminho = ($procs | Where-Object { $_.Path } | Select-Object -First 1).Path
        if (-not $caminho) { $caminho = $script:Caminhos[$nome] }
        foreach ($p in $procs) { try { [void]$p.CloseMainWindow() } catch {} }
        $m.Fechados += @{ Nome = $nome; Caminho = $caminho }
    }
    if ($m.Fechados.Count) {
        $m.Feito += 'fechou ' + (($m.Fechados | ForEach-Object { $_.Nome }) -join ', ')
        $timerFecharForca.Stop(); $timerFecharForca.Start()
    }
    $script:ModoAplicado = $m
    if ($script:Sessao) { $script:Sessao.ModoJogo = $true }
    Update-ModoStatus
}

# Quem não fechou pelo "X" em 5 s é encerrado
$timerFecharForca = New-Object Windows.Threading.DispatcherTimer
$timerFecharForca.Interval = [TimeSpan]::FromSeconds(5)
$timerFecharForca.add_Tick({
    $timerFecharForca.Stop()
    if (-not $script:ModoAplicado) { return }
    foreach ($f in $script:ModoAplicado.Fechados) { Get-Process -Name $f.Nome -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue }
})

function Disable-ModoJogo {
    $m = $script:ModoAplicado
    if (-not $m) { return }
    $script:ModoAplicado = $null
    $timerFecharForca.Stop()
    if ($m.Plano) { try { & powercfg /setactive $m.Plano | Out-Null } catch {} }
    if ($m.Indexacao) { try { Start-Process sc.exe -ArgumentList 'start WSearch' -WindowStyle Hidden } catch {} }
    if ($Config.ModoJogoReabrir) {
        foreach ($f in $m.Fechados) {
            if ($f.Caminho -and (Test-Path $f.Caminho) -and -not (Get-Process -Name $f.Nome -ErrorAction SilentlyContinue)) {
                try { Start-Process $f.Caminho -WindowStyle Minimized } catch {}
            }
        }
    }
    Update-ModoStatus "Modo Jogo desfeito às $((Get-Date).ToString('HH:mm')) ($($m.Jogo) fechado): tudo voltou ao normal."
}

function Update-ModoStatus([string]$texto = '') {
    if ($script:ModoAplicado) {
        $m = $script:ModoAplicado
        $ui.txtModoStatus.Text = "● Ativo para $($m.Jogo) desde $($m.Desde.ToString('HH:mm'))" + $(if ($m.Feito.Count) { ': ' + ($m.Feito -join ' · ') + '.' } else { '.' })
        $ui.txtModoStatus.Foreground = Get-Pincel $Cor.Lima
        $ui.bdModoStatus.Background = Get-Pincel ('#' + $Paleta['1A2614'])
    } else {
        $ui.txtModoStatus.Text = if ($texto) { $texto } elseif ($Config.ModoJogo) { 'Aguardando um jogo abrir.' } else { 'Desligado. Ative para o Monitor preparar o PC sozinho quando um jogo abrir.' }
        $ui.txtModoStatus.Foreground = Get-Pincel '#C8C8CE'
        $ui.bdModoStatus.Background = Get-Pincel '#141416'
    }
}

# Lista de programas que podem ser fechados: os abertos agora (fora do Windows) + os já escolhidos
function Update-ListaFechar {
    $ui.wpModoFechar.Children.Clear()
    if (((Get-Date) - $script:CaminhosHora).TotalSeconds -gt 10) { Update-Caminhos }
    $nomes = New-Object System.Collections.Generic.List[string]
    foreach ($n in @($Config.ModoJogoFechar)) { if ($n -and -not $nomes.Contains($n)) { $nomes.Add($n) } }
    $ignorar = @($NaoJogos | Where-Object { $_ -in 'dwm', 'explorer', 'powershell', 'ApplicationFrameHost', 'LockApp', 'SearchHost',
        'StartMenuExperienceHost', 'TextInputHost', 'ShellExperienceHost', 'Monitor PC' }) + @('svchost', 'conhost', 'RuntimeBroker', 'dllhost', 'sihost', 'ctfmon', 'fake_presentmon', 'PresentMon')
    # Só programas que a pessoa usa: os que têm janela aberta ou os de bandeja mais comuns (não serviços de fundo)
    $comJanela = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } | ForEach-Object { $_.ProcessName })
    $deBandeja = 'Discord', 'Spotify', 'steam', 'EpicGamesLauncher', 'Teams', 'ms-teams', 'Telegram', 'Slack', 'Zoom', 'OneDrive',
        'Dropbox', 'GoogleDriveFS', 'WhatsApp', 'WhatsApp.Root', 'Battle.net', 'EADesktop', 'EALauncher', 'RiotClientServices',
        'UbisoftConnect', 'upc', 'GalaxyClient', 'qbittorrent', 'utorrent', 'obs64', 'chrome', 'msedge', 'firefox', 'opera', 'brave'
    foreach ($p in $script:Processos | Sort-Object RamMB -Descending) {
        if ($nomes.Count -ge 18) { break }
        if ($p.Nome -notin $comJanela -and $p.Nome -notin $deBandeja) { continue }
        $cam = $script:Caminhos[$p.Nome]
        if ($cam -and $cam.StartsWith($env:windir, [StringComparison]::OrdinalIgnoreCase)) { continue }
        if ($p.Nome -in $ignorar -or $p.Nome -eq $script:JogoAtual -or $nomes.Contains($p.Nome)) { continue }
        $nomes.Add($p.Nome)
    }
    foreach ($n in $nomes) {
        $tb = New-Object Windows.Controls.Primitives.ToggleButton
        $tb.Style = $window.FindResource('ChipSel')
        $tb.Tag = $n
        $tb.IsChecked = @($Config.ModoJogoFechar) -contains $n
        $sp = New-Object Windows.Controls.StackPanel
        $sp.Orientation = 'Horizontal'
        $ico = Get-IconeProcesso $n
        if ($ico) { $img = New-Object Windows.Controls.Image; $img.Source = $ico; $img.Width = 16; $img.Height = 16; $img.Margin = '0,0,6,0'; [void]$sp.Children.Add($img) }
        $t = New-Object Windows.Controls.TextBlock; $t.Text = $n
        [void]$sp.Children.Add($t)
        $tb.Content = $sp
        $tb.add_Click({
            param($s, $e)
            $lista = @(@($Config.ModoJogoFechar) | Where-Object { $_ -and $_ -ne $s.Tag })
            if ($s.IsChecked) { $lista += $s.Tag }
            $Config.ModoJogoFechar = $lista
            Save-Config
        })
        [void]$ui.wpModoFechar.Children.Add($tb)
    }
}

$ui.chkModoJogo.IsChecked = [bool]$Config.ModoJogo
$ui.chkModoPlano.IsChecked = [bool]$Config.ModoJogoPlano
$ui.chkModoIndex.IsChecked = [bool]$Config.ModoJogoIndexacao
$ui.chkModoReabrir.IsChecked = [bool]$Config.ModoJogoReabrir
$ui.chkModoIndex.IsEnabled = $EhAdmin
if (-not $EhAdmin) { $ui.chkModoIndex.Content = 'Pausar a indexação de arquivos (precisa abrir como admin)' }
$ui.chkModoJogo.add_Click({
    $Config.ModoJogo = [bool]$ui.chkModoJogo.IsChecked; Save-Config
    if (-not $Config.ModoJogo) { Disable-ModoJogo }
    elseif ($script:JogoAtual) { Enable-ModoJogo $script:JogoAtual }
    Update-ModoStatus
})
$ui.chkModoPlano.add_Click({ $Config.ModoJogoPlano = [bool]$ui.chkModoPlano.IsChecked; Save-Config })
$ui.chkModoIndex.add_Click({ $Config.ModoJogoIndexacao = [bool]$ui.chkModoIndex.IsChecked; Save-Config })
$ui.chkModoReabrir.add_Click({ $Config.ModoJogoReabrir = [bool]$ui.chkModoReabrir.IsChecked; Save-Config })
$ui.btnModoListar.add_Click({ Update-ListaFechar })
Update-ModoStatus

# ---- Sessões de jogo ----
$SessoesPath = Join-Path $ConfigDir 'sessoes.jsonl'
$script:Sessao = $null
$script:JogoAtual = $null
$script:SemJogoSeg = 0

function Start-SessaoJogo($jogo) {
    $script:Sessao = @{
        Jogo = $jogo; Inicio = Get-Date; Fps = New-Object System.Collections.Generic.List[double]
        Low = New-Object System.Collections.Generic.List[double]; Quedas = 0; Causas = @(); TempMax = $null; GpuSoma = 0.0
        CpuMax = 0; Pings = @(); ModoJogo = [bool]$script:ModoAplicado
    }
}

function Stop-SessaoJogo {
    $s = $script:Sessao
    $script:Sessao = $null
    if (-not $s -or $s.Fps.Count -lt 60) { return }   # menos de 1 minuto não vale
    $ping = Get-StatsPing $s.Pings
    $principal = $s.Causas | Group-Object | Sort-Object Count -Descending | Select-Object -First 1
    $reg = [ordered]@{
        Jogo = $s.Jogo; Inicio = $s.Inicio.ToString('s'); Minutos = [Math]::Round(((Get-Date) - $s.Inicio).TotalMinutes, 1)
        FpsMedio = [Math]::Round(($s.Fps | Measure-Object -Average).Average, 1)
        Low1 = [Math]::Round(($s.Low | Measure-Object -Average).Average, 1)
        Quedas = $s.Quedas; CausaPrincipal = $(if ($principal) { $principal.Name } else { '' })
        GpuTempMax = $s.TempMax; GpuUsoMedio = [Math]::Round($s.GpuSoma / $s.Fps.Count); CpuMax = $s.CpuMax
        PingMedio = $(if ($null -ne $ping.Media) { [Math]::Round($ping.Media) } else { $null }); PerdaPct = [Math]::Round($ping.Perda, 1)
        ModoJogo = $s.ModoJogo
    }
    try { Add-Content -Path $SessoesPath -Value ($reg | ConvertTo-Json -Compress) -Encoding UTF8 } catch {}
    Show-Sessoes
}

function Update-SessaoJogo($temJogo, $nomeJogo, $a, $si) {
    if ($temJogo) {
        $script:SemJogoSeg = 0
        if ($script:JogoAtual -ne $nomeJogo) {
            if ($script:JogoAtual) { Stop-SessaoJogo }
            $script:JogoAtual = $nomeJogo
            Start-SessaoJogo $nomeJogo
        }
        $s = $script:Sessao
        $s.Fps.Add($a.Fps); $s.Low.Add($a.Low1)
        if ($null -ne $si.GpuT -and ($null -eq $s.TempMax -or $si.GpuT -gt $s.TempMax)) { $s.TempMax = $si.GpuT }
        $s.GpuSoma += $si.Gpu
        if ($si.Cpu -gt $s.CpuMax) { $s.CpuMax = $si.Cpu }
        # O Modo Jogo entra depois de 5 s com o jogo aberto (evita reagir a telas de carregamento rápidas)
        if ($Config.ModoJogo -and -not $script:ModoAplicado -and ((Get-Date) - $s.Inicio).TotalSeconds -ge 5) { Enable-ModoJogo $nomeJogo }
    } elseif ($script:JogoAtual) {
        $script:SemJogoSeg++
        if ($script:SemJogoSeg -ge 15) {   # 15 s sem o jogo = jogo fechado
            Stop-SessaoJogo
            Disable-ModoJogo
            $script:JogoAtual = $null
        }
    }
}

function Show-Sessoes {
    $linhas = @()
    if (Test-Path $SessoesPath) {
        $linhas = @(Get-Content $SessoesPath -Encoding UTF8 -ErrorAction SilentlyContinue | Where-Object { $_ } | ForEach-Object { try { $_ | ConvertFrom-Json } catch {} })
    }
    $itens = New-Object System.Collections.Generic.List[object]
    for ($i = $linhas.Count - 1; $i -ge 0 -and $itens.Count -lt 15; $i--) {
        $r = $linhas[$i]
        # Compara com a sessão anterior do mesmo jogo
        $ant = $null
        for ($j = $i - 1; $j -ge 0; $j--) { if ($linhas[$j].Jogo -eq $r.Jogo) { $ant = $linhas[$j]; break } }
        $delta = ''; $deltaCor = '#8E8E96'
        if ($ant -and $ant.FpsMedio -gt 0) {
            $pct = ($r.FpsMedio - $ant.FpsMedio) / $ant.FpsMedio * 100
            $delta = '  {0}{1:N0}%' -f $(if ($pct -ge 0) { '+' } else { '' }), $pct
            $deltaCor = if ($pct -ge 3) { $Cor.Lima } elseif ($pct -le -3) { $Cor.Vermelho } else { '#8E8E96' }
        }
        $quando = [datetime]::Parse($r.Inicio)
        $itens.Add([pscustomobject]@{
            Jogo = $r.Jogo + $(if ($r.ModoJogo) { '  ·  Modo Jogo' } else { '' })
            Quando = ('{0:dd/MM HH:mm} · {1:N0} min' -f $quando, $r.Minutos)
            FpsTxt = '{0:N0}' -f $r.FpsMedio; Delta = $delta; DeltaCor = $deltaCor
            LowTxt = '{0:N0}' -f $r.Low1
            QuedasTxt = "$($r.Quedas)"
            CausaTxt = if ($r.Quedas -and $r.CausaPrincipal) { $r.CausaPrincipal } else { '' }
            TempTxt = if ($null -ne $r.GpuTempMax) { "$($r.GpuTempMax) °C" } else { 'N/D' }
            PingTxt = if ($null -ne $r.PingMedio) { "$($r.PingMedio) ms" + $(if ($r.PerdaPct -gt 0) { " · $($r.PerdaPct)% perda" } else { '' }) } else { '--' }
        })
    }
    $ui.icSessoes.ItemsSource = $itens
    $ui.txtSemSessoes.Visibility = if ($itens.Count) { 'Collapsed' } else { 'Visible' }
}
$ui.btnLimparSessoes.add_Click({
    $r = [Windows.MessageBox]::Show('Apagar o histórico de sessões de jogo?', $AppNome, 'YesNo', 'Question')
    if ($r -eq 'Yes') { Remove-Item $SessoesPath -ErrorAction SilentlyContinue; Show-Sessoes }
})
Show-Sessoes
#endregion

#region Configurações na tela
$AtalhoInicio = Join-Path ([Environment]::GetFolderPath('Startup')) 'Monitor PC.lnk'

function Show-Config {
    $ui.chkAlertas.IsChecked = [bool]$Config.AlertasAtivos
    $ui.chkSom.IsChecked = [bool]$Config.Som
    $ui.chkIniciar.IsChecked = Test-Path $AtalhoInicio
    $ui.slCpu.Value = $Config.CpuLimite
    $ui.slRam.Value = $Config.RamLimite
    $ui.slDsc.Value = $Config.DiscoLimite
    $ui.slEsp.Value = $Config.EspacoMinimo
    $ui.slSus.Value = $Config.Sustentar
    $ui.slRep.Value = $Config.IntervaloAlertaMin
    $ui.slGpuT.Value = $Config.GpuTempLimite
    $ui.txtAno.Text = $Config.AnoCompra
}

# Atalho que abre o Monitor escondendo a janela do PowerShell, com o ícone do app
function New-Atalho([string]$caminho, [string]$extra = '') {
    $sh = New-Object -ComObject WScript.Shell
    $lnk = $sh.CreateShortcut($caminho)
    if ($script:TemLancador -or (Update-Lancador)) {
        $lnk.TargetPath = $LancadorPath
        $lnk.Arguments = $extra
    } else {
        $lnk.TargetPath = "$env:windir\System32\WindowsPowerShell\v1.0\powershell.exe"
        $lnk.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`" $extra".Trim()
        $lnk.WindowStyle = 7   # minimizado: evita piscar a janela preta
    }
    $lnk.WorkingDirectory = $PSScriptRoot
    $lnk.Description = 'Monitor PC - desempenho, alertas e segurança'
    if (Test-Path $IconePath) { $lnk.IconLocation = "$IconePath,0" }
    $lnk.Save()
}

# Atalhos criados por versões antigas apontavam para o powershell.exe: troca pelo lançador
if ($script:AtualizarAtalhosAntigos) {
    $shAntigo = New-Object -ComObject WScript.Shell
    foreach ($par in @(@($AtalhoInicio, '-Minimizado'), @($AtalhoDesktop, ''), @($AtalhoMenu, ''))) {
        if ((Test-Path $par[0]) -and $shAntigo.CreateShortcut($par[0]).TargetPath -like '*powershell.exe') {
            try { New-Atalho $par[0] $par[1] } catch {}
        }
    }
}

function Set-IniciarComWindows([bool]$ligar) {
    if ($ligar) {
        New-Atalho $AtalhoInicio '-Minimizado'
    } elseif (Test-Path $AtalhoInicio) {
        Remove-Item $AtalhoInicio -Force
    }
}

# Alertas: tudo é salvo na hora em que muda (não depende de clicar em Salvar)
$script:CarregandoConfig = $true
function Save-ConfigAlertas {
    if ($script:CarregandoConfig) { return }
    $Config.AlertasAtivos = [bool]$ui.chkAlertas.IsChecked
    $Config.Som = [bool]$ui.chkSom.IsChecked
    $Config.CpuLimite = [int]$ui.slCpu.Value
    $Config.RamLimite = [int]$ui.slRam.Value
    $Config.DiscoLimite = [int]$ui.slDsc.Value
    $Config.EspacoMinimo = [int]$ui.slEsp.Value
    $Config.Sustentar = [int]$ui.slSus.Value
    $Config.IntervaloAlertaMin = [int]$ui.slRep.Value
    $Config.GpuTempLimite = [int]$ui.slGpuT.Value
    $script:UltimoAlerta = @{}
    Save-Config
    $ui.txtSalvo.Text = '✓ Salvo automaticamente às {0:HH:mm:ss}' -f (Get-Date)
}
# Sliders disparam muitas mudanças ao arrastar: salva só quando param por um instante
$timerSalvar = New-Object Windows.Threading.DispatcherTimer
$timerSalvar.Interval = [TimeSpan]::FromMilliseconds(600)
$timerSalvar.add_Tick({ $timerSalvar.Stop(); Save-ConfigAlertas })
foreach ($sl in $ui.slCpu, $ui.slRam, $ui.slDsc, $ui.slEsp, $ui.slSus, $ui.slRep, $ui.slGpuT) {
    $sl.add_ValueChanged({ if (-not $script:CarregandoConfig) { $timerSalvar.Stop(); $timerSalvar.Start() } })
}
foreach ($chk in $ui.chkAlertas, $ui.chkSom) { $chk.add_Click({ Save-ConfigAlertas }) }
$ui.chkIniciar.add_Click({
    try { Set-IniciarComWindows ([bool]$ui.chkIniciar.IsChecked) } catch {}
    Save-ConfigAlertas
})
$ui.btnSalvar.add_Click({
    try { Set-IniciarComWindows ([bool]$ui.chkIniciar.IsChecked) } catch {}
    Save-ConfigAlertas
})

$ui.btnTestar.add_Click({ Send-Alerta 'Teste de alerta' 'Se você está vendo isto, os alertas estão funcionando.' })
$ui.btnLimparHist.add_Click({ $ui.lbHist.Items.Clear() })
$ui.txtAno.add_LostFocus({
    $Config.AnoCompra = $ui.txtAno.Text.Trim()
    Save-Config
})
$ui.btnAnalisar.add_Click({
    $Config.AnoCompra = $ui.txtAno.Text.Trim()
    Save-Config
    Start-Analise
})

# Menu lateral: cada item troca a aba (índice na ordem do menu)
$Menu = @($ui.mnPainel, $ui.mnProc, $ui.mnAlertas, $ui.mnVida, $ui.mnOtim, $ui.mnSeg, $ui.mnApps, $ui.mnJogos)
# Índice da aba (ordem no XAML) de cada item do menu; 7 = Aparência, que não tem item no menu
$AbaDoMenu = @{ mnPainel = 0; mnProc = 1; mnAlertas = 2; mnVida = 3; mnOtim = 4; mnSeg = 5; mnApps = 6; mnJogos = 8 }
function Select-Aba([int]$i) {
    foreach ($m in $Menu) { if ($AbaDoMenu[$m.Name] -eq $i) { $m.IsChecked = $true } }
}
for ($i = 0; $i -lt $Menu.Count; $i++) {
    $Menu[$i].add_Checked({
        param($s, $e)
        $idx = $AbaDoMenu[$s.Name]
        $ui.tabs.SelectedIndex = $idx
        if ($idx -eq 8) { Update-EstadoFps; if (-not $ui.wpModoFechar.Children.Count) { Update-ListaFechar } }
        if ($idx -eq 1) { Update-GradeProcessos }
        if ($idx -eq 2) { $script:NaoLidos = 0; Update-Badge }
        if ($idx -eq 3 -and -not $ui.icVida.ItemsSource) { Start-Analise }
        if ($idx -eq 6 -and -not $script:Apps.Count) { Start-Apps }
    })
}
$ui.btnVerTodos.add_Click({ Select-Aba 1 })
$ui.btnDica.add_Click({ if ($script:DicaGuia) { Show-Guia $script:DicaGuia } else { Select-Aba 4 } })
$ui.btnSino.add_Click({ Select-Aba 2 })
$ui.btnSegMini.add_Click({ Select-Aba 5 })

$ui.txtBusca.add_TextChanged({
    $ui.phBusca.Visibility = if ($ui.txtBusca.Text) { 'Collapsed' } else { 'Visible' }
    if ($ui.txtBusca.Text -and $ui.tabs.SelectedIndex -ne 1) { Select-Aba 1 }
    Update-GradeProcessos
})
foreach ($rb in $ui.rbRam, $ui.rbCpu, $ui.rbNome) { $rb.add_Checked({ if ($ui.dgProc) { Update-GradeProcessos } }) }

$Protegidos = 'System', 'Registry', 'smss', 'csrss', 'wininit', 'winlogon', 'services', 'lsass', 'svchost', 'dwm', 'fontdrvhost', 'Memory Compression', 'MsMpEng'
$ui.btnFinalizar.add_Click({
    $sel = $ui.dgProc.SelectedItem
    if (-not $sel) { [Windows.MessageBox]::Show('Selecione um processo na lista.', $AppNome) | Out-Null; return }
    if ($Protegidos -contains $sel.Nome) {
        [Windows.MessageBox]::Show("'$($sel.Nome)' é um processo essencial do Windows e não deve ser finalizado.", $AppNome, 'OK', 'Warning') | Out-Null
        return
    }
    $resp = [Windows.MessageBox]::Show("Finalizar '$($sel.Nome)' ($($sel.Qtd) instância(s))?`n`nTrabalho não salvo nesse programa será perdido.", $AppNome, 'YesNo', 'Warning')
    if ($resp -ne 'Yes') { return }
    $falhas = 0
    foreach ($p in Get-Process -Name $sel.Nome -ErrorAction SilentlyContinue) {
        try { $p.Kill() } catch { $falhas++ }
    }
    if ($falhas) {
        [Windows.MessageBox]::Show("$falhas instância(s) não puderam ser finalizadas. Tente abrir o Monitor como administrador.", $AppNome) | Out-Null
    }
})

# Selo de administrador
if ($EhAdmin) {
    $ui.txtAdmin.Text = '● Administrador'
    $ui.txtAdmin.Foreground = Get-Pincel $Cor.Lima
    $ui.btnAdmin.Visibility = 'Collapsed'
} else {
    $ui.txtAdmin.Text = 'Modo comum'
    $ui.txtAdmin.Foreground = Get-Pincel $Cor.Suave
}
$ui.btnAdmin.add_Click({
    try {
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`" -Reiniciando"
        Exit-App
    } catch {}
})
#endregion

#region Janela, bandeja e inicialização
$script:Saindo = $false
$script:AvisouBandeja = $false

function Show-Janela {
    $window.Show()
    if ($window.WindowState -eq 'Minimized') { $window.WindowState = 'Normal' }
    $window.Activate() | Out-Null
}

function Remove-IconeBandeja {
    try { $tray.Visible = $false; $tray.Dispose() } catch {}
}

function Exit-App {
    $script:Saindo = $true
    $timer.Stop(); $timerAnim.Stop(); $timerSegAgenda.Stop(); $timerMostrar.Stop()
    $GpuSync.Rodando = $false
    $PingSync.Rodando = $false
    try { Stop-SessaoJogo; Disable-ModoJogo } catch {}   # salva a sessão e devolve o PC ao normal
    try { $Overlay.Close() } catch {}
    Stop-Fps
    Remove-IconeBandeja
    try { $Trava.ReleaseMutex() } catch {}
    $window.Close()
    $app.Shutdown()
}

# Outra cópia tentou abrir: mostra esta janela
$timerMostrar = New-Object Windows.Threading.DispatcherTimer
$timerMostrar.Interval = [TimeSpan]::FromMilliseconds(400)
$timerMostrar.add_Tick({ if ($EventoMostrar.WaitOne(0)) { Show-Janela } })

$ui.btnMin.add_Click({ $window.WindowState = 'Minimized' })
$ui.btnMax.add_Click({ $window.WindowState = if ($window.WindowState -eq 'Maximized') { 'Normal' } else { 'Maximized' } })
$ui.btnFechar.add_Click({ $window.Close() })
$ui.btnSair.add_Click({ Exit-App })
# Sem borda nativa, a janela maximizada passa um pouco da tela: compensa com margem
$window.add_StateChanged({
    $ui.raiz.Margin = if ($window.WindowState -eq 'Maximized') { '7' } else { '0' }
    $ui.btnMax.Content = if ($window.WindowState -eq 'Maximized') { [string][char]0xE923 } else { [string][char]0xE922 }
    Update-Animacoes
})

# As animações do Painel (ponto "ao vivo", brilho do chip, anéis) só rodam com o Painel na tela.
# Fora dele elas só gastariam CPU/GPU — e atrasavam as atualizações das outras abas.
function Update-Animacoes {
    $visivel = $ui.tabs.SelectedIndex -eq 0 -and $window.IsVisible -and $window.WindowState -ne 'Minimized'
    if ($visivel) {
        $sbPonto.Resume($ui.pontoVivo); $sbChip.Resume($ui.brilhoChip); $timerAnim.Start()
    } else {
        $sbPonto.Pause($ui.pontoVivo); $sbChip.Pause($ui.brilhoChip); $timerAnim.Stop()
    }
}
$ui.tabs.add_SelectionChanged({ param($s, $e) if ($e.OriginalSource -eq $ui.tabs) { Update-Animacoes } })
$window.add_IsVisibleChanged({ Update-Animacoes })

$window.add_Closing({
    param($s, $e)
    if ($script:Saindo) { return }
    $e.Cancel = $true
    $window.Hide()
    if (-not $script:AvisouBandeja) {
        $tray.ShowBalloonTip(4000, $AppNome, 'Continuo monitorando aqui na bandeja. Clique com o botão direito no ícone para sair.', 'Info')
        $script:AvisouBandeja = $true
    }
})

Show-Config
$script:CarregandoConfig = $false
# Prioridade Normal: a leitura nunca fica esperando na fila atrás do desenho da tela
$timer = New-Object Windows.Threading.DispatcherTimer([Windows.Threading.DispatcherPriority]::Normal)
$timer.Interval = [TimeSpan]::FromSeconds($Intervalo)
$timer.add_Tick({ Update-Tela })
$timer.Start()
Update-Animacoes
$timerSegAgenda.Start()
$timerMostrar.Start()
$window.add_ContentRendered({ Update-Tela })

$app = New-Object Windows.Application
$app.ShutdownMode = 'OnExplicitShutdown'
# Um erro inesperado não derruba o Monitor (e não deixa ícone "fantasma" na bandeja)
$app.add_DispatcherUnhandledException({
    param($s, $e)
    $e.Handled = $true
    try { Add-Content -Path (Join-Path $ConfigDir 'erros.log') -Value ('{0:dd/MM HH:mm:ss}  {1}' -f (Get-Date), $e.Exception.Message) -Encoding UTF8 } catch {}
})
if ($Minimizado) {
    $script:AvisouBandeja = $true
    Update-Tela
    try { [void]$app.Run() } finally { Remove-IconeBandeja }
} else {
    try { [void]$app.Run($window) } finally { Remove-IconeBandeja }
}
#endregion
