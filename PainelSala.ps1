#Requires -Version 5.1
<#
  PainelSala.ps1 — Painel de Sala de Reuniões para Windows
  ---------------------------------------------------------
  Mostra em tela cheia (estilo Logitech Tap Scheduler): relógio, nome da sala,
  status (Disponível / Começa em X min / Em reunião), próximas reuniões e um
  botão "Entrar no Teams" que abre a reunião direto no aplicativo do Teams.

  Fonte da agenda: Outlook clássico (automação COM) — sem senha, sem Azure.
  Requisitos: Windows 10/11, Outlook clássico com a conta logada, Teams instalado.

  Uso:  powershell -STA -ExecutionPolicy Bypass -File PainelSala.ps1 [-Demo] [-Janela] [-Config caminho\config.json]
  Teclas: Esc/F11 alterna tela cheia · F5 atualiza · Ctrl+Q fecha
#>
[CmdletBinding()]
param(
  [switch]$Demo,
  [switch]$Janela,
  [string]$Config = ''
)

$ErrorActionPreference = 'Stop'
$script:Versao = '1.1'
$script:AppDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $Config) { $Config = Join-Path $script:AppDir 'config.json' }

# ============================================================================
#  WIN32 / WPF
# ============================================================================
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing
try {
  Add-Type -Namespace Native -Name Win32 -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")]   public static extern bool SetProcessDPIAware();
[DllImport("user32.dll")]   public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
[DllImport("kernel32.dll")] public static extern uint SetThreadExecutionState(uint esFlags);
'@
  $h = [Native.Win32]::GetConsoleWindow()
  if ($h -ne [IntPtr]::Zero) { [void][Native.Win32]::ShowWindow($h, 0) }   # esconde o console
  [void][Native.Win32]::SetProcessDPIAware()                                 # texto nítido em telas com escala 125%/150%
} catch {}

# ============================================================================
#  CONFIGURAÇÃO
# ============================================================================
$Defaults = [ordered]@{
  NomeSala                = 'Sala de Reuniões'
  Subtitulo               = 'ABRAJEEP'
  CalendarioCompartilhado = ''         # e-mail de uma sala ou caixa compartilhada; vazio = agenda pessoal
  HorasAFrente            = 30         # até quantas horas à frente listar reuniões
  MaxItens                = 6          # máximo de linhas na lista
  IntervaloAtualizacaoSeg = 60         # de quanto em quanto tempo reler o Outlook
  MinutosAntesParaEntrar  = 15         # libera o botão "Entrar" X minutos antes do início
  IntervaloLivreMin       = 30         # mostra "Livre" entre reuniões quando o intervalo for >= X min (0 = não mostrar)
  MostrarOrganizador      = $true
  TelaCheia               = $true
  SempreNoTopo            = $false     # deixe false para a janela do Teams aparecer por cima ao entrar
  Monitor                 = 0          # índice do monitor (0 = principal, 1 = segundo monitor/TV...)
  ManterTelaLigada        = $true
  AbrirVia                = 'msteams'  # 'msteams' abre direto no app do Teams; 'https' abre pelo navegador
  PermitirReserva         = $false     # mostra botão "Reservar" (cria compromisso rápido na agenda)
  DuracaoReservaMin       = 30
  Fonte                   = 'Segoe UI'
  ImagemFundo             = 'fundo.jpg'
  Logo                    = 'logo.png'
  Demo                    = $false
}
$Cfg = @{}
foreach ($k in $Defaults.Keys) { $Cfg[$k] = $Defaults[$k] }
if (Test-Path -LiteralPath $Config) {
  try {
    $json = Get-Content -LiteralPath $Config -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($k in $Defaults.Keys) {
      $p = $json.PSObject.Properties[$k]
      if ($null -ne $p -and $null -ne $p.Value) { $Cfg[$k] = $p.Value }
    }
  } catch {
    [void][System.Windows.Forms.MessageBox]::Show("Erro ao ler config.json:`n$($_.Exception.Message)`n`nUsando valores padrão.", 'Painel de Sala')
  }
}
$script:Demo = ($Demo.IsPresent -or [bool]$Cfg.Demo)
$script:FullscreenWanted = ([bool]$Cfg.TelaCheia -and -not $Janela.IsPresent)

# ============================================================================
#  LÓGICA (independente da interface)
# ============================================================================
$script:TeamsUrlProp = 'http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/SkypeTeamsMeetingUrl'
$script:PtBR = [System.Globalization.CultureInfo]::GetCultureInfo('pt-BR')

function Get-TeamsLink {
  # Extrai o link de reunião do Teams de um texto (corpo, local ou propriedade do compromisso).
  param([string]$Text, [int]$Depth = 0)
  if (-not $Text -or $Depth -gt 2) { return $null }
  # Links reescritos pelo Defender (safelinks) — decodifica o destino real
  foreach ($m in [regex]::Matches($Text, 'safelinks\.protection\.outlook\.com/[^\s"<>]*?[?&]url=([^&\s"<>]+)', 'IgnoreCase')) {
    $inner = Get-TeamsLink -Text ([uri]::UnescapeDataString($m.Groups[1].Value)) -Depth ($Depth + 1)
    if ($inner) { return $inner }
  }
  $rx = 'https://(?:teams\.microsoft\.com|teams\.live\.com|teams\.microsoft\.us|gov\.teams\.microsoft\.us|dod\.teams\.microsoft\.us|teams\.cloud\.microsoft)/(?:l/meetup-join|meet)/[^\s"<>\)\]]+'
  $m = [regex]::Match($Text, $rx, 'IgnoreCase')
  if ($m.Success) { return $m.Value.TrimEnd('.', ',', ';', '>', ')', ']') }
  return $null
}

function ConvertTo-TeamsProtocol {
  # https://teams.microsoft.com/... -> msteams://teams.microsoft.com/...  (abre direto no app, sem passar pelo navegador)
  param([string]$Url)
  if ($Url -match '^https://teams\.microsoft\.com/') { return ($Url -replace '^https://', 'msteams://') }
  return $Url
}

function Get-DiaRotulo {
  param([datetime]$Data, [datetime]$Agora)
  $d = ($Data.Date - $Agora.Date).Days
  if ($d -eq 0) { return 'Hoje' }
  if ($d -eq 1) { return 'Amanhã' }
  $t = $Data.ToString('dddd, dd/MM', $script:PtBR)
  return $t.Substring(0, 1).ToUpper() + $t.Substring(1)
}

function Get-Estado {
  # Calcula o estado do painel a partir da lista de reuniões e do horário atual.
  param([object[]]$Reunioes, [datetime]$Agora, [int]$MinutosAntes = 15)
  $lista = @($Reunioes | Where-Object { $_ -ne $null } | Sort-Object Inicio)
  $atual = $lista | Where-Object { $_.Ocupado -and $_.Inicio -le $Agora -and $_.Fim -gt $Agora } | Select-Object -First 1
  $proximas = @($lista | Where-Object { $_.Inicio -gt $Agora })
  $proxima = $proximas | Select-Object -First 1
  $tipo = 'livre'; $emBreve = $null
  if ($atual) { $tipo = 'ocupado' }
  elseif ($proxima -and $proxima.Ocupado -and (($proxima.Inicio - $Agora).TotalMinutes -le $MinutosAntes)) { $tipo = 'breve'; $emBreve = $proxima }
  $restantes = if ($tipo -eq 'breve') { @($proximas | Select-Object -Skip 1) } else { $proximas }
  [pscustomobject]@{
    Tipo     = $tipo
    Atual    = $atual
    Proxima  = $proxima
    EmBreve  = $emBreve
    Lista    = $restantes
  }
}

function Test-PodeEntrar {
  param($Reuniao, [datetime]$Agora, [int]$MinutosAntes = 15)
  if (-not $Reuniao -or -not $Reuniao.Link) { return $false }
  return ($Agora -ge $Reuniao.Inicio.AddMinutes(-$MinutosAntes) -and $Agora -lt $Reuniao.Fim)
}

function Get-DemoMeetings {
  $n = Get-Date
  $base = Get-Date -Hour $n.Hour -Minute $n.Minute -Second 0
  $link = 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_demo%40thread.v2/0?context=%7b%22Tid%22%3a%22demo%22%7d'
  $amanha = (Get-Date -Hour 9 -Minute 0 -Second 0).AddDays(1)
  @(
    [pscustomobject]@{ Assunto = 'Comissão Comercial — Stellantis x Rede';   Inicio = $base.AddMinutes(10);  Fim = $base.AddMinutes(70);  Organizador = 'ABRAJEEP';          Local = 'Microsoft Teams'; Link = $link; Ocupado = $true }
    [pscustomobject]@{ Assunto = 'Alinhamento Programa ABRAJEEP IA';         Inicio = $base.AddMinutes(100); Fim = $base.AddMinutes(130); Organizador = 'Diretoria Executiva'; Local = 'Microsoft Teams'; Link = $link; Ocupado = $true }
    [pscustomobject]@{ Assunto = 'Fechamento Nota & Anota — Setembro';       Inicio = $base.AddMinutes(200); Fim = $base.AddMinutes(245); Organizador = 'Comunicação';        Local = 'Sala 2';          Link = $null; Ocupado = $true }
    [pscustomobject]@{ Assunto = 'Reunião de Diretoria Vendas JEEP/RAM';     Inicio = $amanha;               Fim = $amanha.AddHours(1.5);  Organizador = 'Stellantis';         Local = 'Microsoft Teams'; Link = $link; Ocupado = $true }
    [pscustomobject]@{ Assunto = 'Comissão de Pós-Vendas';                   Inicio = $amanha.AddHours(5);   Fim = $amanha.AddHours(6);    Organizador = 'ABRAJEEP';          Local = 'Microsoft Teams'; Link = $link; Ocupado = $true }
  )
}

# ============================================================================
#  OUTLOOK (COM)
# ============================================================================
function Connect-Outlook {
  if ($script:OutlookNs) {
    try { $null = $script:OutlookNs.CurrentProfileName; return } catch { $script:OutlookNs = $null }
  }
  try {
    $app = New-Object -ComObject Outlook.Application
  } catch {
    throw 'Outlook clássico não encontrado (o "novo Outlook" não é suportado).'
  }
  $script:OutlookNs = $app.GetNamespace('MAPI')
}

function Get-CalendarFolder {
  $ns = $script:OutlookNs
  if ($Cfg.CalendarioCompartilhado) {
    $r = $ns.CreateRecipient([string]$Cfg.CalendarioCompartilhado)
    [void]$r.Resolve()
    if (-not $r.Resolved) { throw "Agenda '$($Cfg.CalendarioCompartilhado)' não encontrada." }
    return $ns.GetSharedDefaultFolder($r, 9)   # olFolderCalendar
  }
  return $ns.GetDefaultFolder(9)
}

function Read-Meetings {
  Connect-Outlook
  $folder = Get-CalendarFolder
  $items = $folder.Items
  $items.IncludeRecurrences = $true
  $items.Sort('[Start]')
  $agora = Get-Date
  $ini = $agora.AddMinutes(-1)
  $fim = $agora.AddHours([double]$Cfg.HorasAFrente)
  $sel = $null
  foreach ($fmt in @('g', 'MM/dd/yyyy HH:mm')) {
    try {
      $filtro = "[End] >= '" + $ini.ToString($fmt) + "' AND [Start] <= '" + $fim.ToString($fmt) + "'"
      $sel = $items.Restrict($filtro)
      break
    } catch { $sel = $null }
  }
  if (-not $sel) { throw 'Não foi possível filtrar a agenda.' }

  $lista = New-Object System.Collections.Generic.List[object]
  $n = 0
  foreach ($it in $sel) {
    if ($null -eq $it) { continue }
    $n++; if ($n -gt 80) { break }
    try {
      if ([bool]$it.AllDayEvent) { continue }
      $ms = [int]$it.MeetingStatus
      if ($ms -eq 5 -or $ms -eq 7) { continue }          # cancelada
      if ([int]$it.ResponseStatus -eq 4) { continue }    # recusada
      $link = $null
      try { $link = Get-TeamsLink -Text ([string]$it.PropertyAccessor.GetProperty($script:TeamsUrlProp)) } catch {}
      if (-not $link) { try { $link = Get-TeamsLink -Text ([string]$it.Location) } catch {} }
      if (-not $link) { try { $link = Get-TeamsLink -Text ([string]$it.Body) } catch {} }
      $org = ''
      if ([bool]$Cfg.MostrarOrganizador) { try { $org = [string]$it.Organizer } catch {} }
      $assunto = [string]$it.Subject
      if (-not $assunto) { $assunto = '(sem assunto)' }
      $lista.Add([pscustomobject]@{
        Assunto     = $assunto
        Inicio      = [datetime]$it.Start
        Fim         = [datetime]$it.End
        Local       = [string]$it.Location
        Organizador = $org
        Link        = $link
        Ocupado     = ([int]$it.BusyStatus -ne 0)         # olFree = 0
      })
    } catch {}
  }
  return @($lista | Sort-Object Inicio)
}

function New-Reserva {
  Connect-Outlook
  $folder = Get-CalendarFolder
  $appt = $folder.Items.Add(1)   # olAppointmentItem
  $agora = Get-Date
  $inicio = $agora.AddSeconds(-$agora.Second)
  $appt.Subject = 'Reserva rápida — ' + $Cfg.NomeSala
  $appt.Start = $inicio
  $appt.Duration = [int]$Cfg.DuracaoReservaMin
  $appt.BusyStatus = 2           # olBusy
  $appt.ReminderSet = $false
  $appt.Location = [string]$Cfg.NomeSala
  $appt.Save()
}

function Open-Teams {
  param([string]$Url)
  if (-not $Url) { return }
  $alvo = $Url
  if (([string]$Cfg.AbrirVia).ToLower() -eq 'msteams') {
    $temProtocolo = $false
    try { $temProtocolo = Test-Path 'Registry::HKEY_CLASSES_ROOT\msteams' } catch {}
    if ($temProtocolo) { $alvo = ConvertTo-TeamsProtocol $Url }
  }
  try { Start-Process $alvo } catch { try { Start-Process $Url } catch {} }
}

# ============================================================================
#  INTERFACE (WPF)
# ============================================================================
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Painel de Sala" Width="1280" Height="720" Background="#04201B"
        WindowStartupLocation="Manual" Foreground="White" TextOptions.TextFormattingMode="Ideal">
  <Window.Resources>
    <Style x:Key="Btn" TargetType="Button">
      <Setter Property="Foreground" Value="#0B2F27"/>
      <Setter Property="Background" Value="White"/>
      <Setter Property="FontSize" Value="24"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="MinHeight" Value="64"/>
      <Setter Property="Padding" Value="28,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="14" Padding="{TemplateBinding Padding}" MinHeight="{TemplateBinding MinHeight}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.88"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.7"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="bd" Property="Opacity" Value="0.35"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BtnSmall" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="FontSize" Value="20"/>
      <Setter Property="MinHeight" Value="54"/>
      <Setter Property="Padding" Value="22,0"/>
    </Style>
    <Style x:Key="BtnGhost" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="Background" Value="#2EFFFFFF"/>
      <Setter Property="Foreground" Value="White"/>
    </Style>
    <Style x:Key="Icon" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="Segoe MDL2 Assets"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style x:Key="BtnMini" TargetType="Button" BasedOn="{StaticResource BtnGhost}">
      <Setter Property="FontSize" Value="17"/>
      <Setter Property="MinHeight" Value="46"/>
      <Setter Property="Padding" Value="16,0"/>
      <Setter Property="Margin" Value="8,0,0,0"/>
    </Style>
    <Style x:Key="Rotulo" TargetType="TextBlock">
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="Opacity" Value="0.7"/>
      <Setter Property="Margin" Value="2,0,0,0"/>
    </Style>
    <Style x:Key="Campo" TargetType="TextBox">
      <Setter Property="Background" Value="#1AFFFFFF"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="CaretBrush" Value="White"/>
      <Setter Property="SelectionBrush" Value="#22C55E"/>
      <Setter Property="BorderBrush" Value="#55FFFFFF"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize" Value="20"/>
      <Setter Property="Padding" Value="12,9"/>
      <Setter Property="Margin" Value="0,6,0,16"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="10" Padding="{TemplateBinding Padding}">
              <ScrollViewer x:Name="PART_ContentHost" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter Property="BorderBrush" Value="#22C55E"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid x:Name="Root">
    <Grid x:Name="BgLayer"/>
    <Rectangle>
      <Rectangle.Fill>
        <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
          <GradientStop Color="#99000000" Offset="0"/>
          <GradientStop Color="#33000000" Offset="1"/>
        </LinearGradientBrush>
      </Rectangle.Fill>
    </Rectangle>

    <!-- "LED" laterais, como no painel físico -->
    <Rectangle x:Name="LedLeft"  Width="12" HorizontalAlignment="Left"  Fill="#22C55E"/>
    <Rectangle x:Name="LedRight" Width="12" HorizontalAlignment="Right" Fill="#22C55E"/>

    <Viewbox Stretch="Uniform" Margin="20,0">
      <Grid Width="1600" Height="900">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="800"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- ESQUERDA: relógio, data, sala, logo -->
        <Grid Grid.Column="0" Margin="64,52,36,52">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <TextBlock Grid.Row="0" x:Name="TxtHora" Text="--:--" FontSize="140" FontWeight="Light" Margin="-6,-20,0,0"/>
          <TextBlock Grid.Row="1" x:Name="TxtData" Text="" FontSize="28" Opacity="0.85" Margin="4,-8,0,0"/>
          <StackPanel Grid.Row="2" VerticalAlignment="Center" Margin="4,0,0,0">
            <TextBlock x:Name="TxtSala" Text="Sala de Reuniões" FontSize="66" FontWeight="SemiBold" TextWrapping="Wrap"/>
            <TextBlock x:Name="TxtSubtitulo" Text="" FontSize="26" Opacity="0.8" Margin="2,10,0,0"/>
          </StackPanel>
          <Image Grid.Row="3" x:Name="ImgLogo" Height="56" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="0,0,0,4"/>
        </Grid>

        <!-- DIREITA: status + lista -->
        <Border Grid.Column="1" Margin="0,44,52,44" Background="#8F07201B" CornerRadius="24" Padding="22">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <Border Grid.Row="0" x:Name="CardStatus" Background="#1FA84A" CornerRadius="18" Padding="28,22,28,24" MinHeight="236">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <TextBlock Grid.Row="0" x:Name="TxtStatus" Text="Disponível" FontSize="52" FontWeight="Bold" Margin="0,-6,0,0"/>
                <TextBlock Grid.Row="1" x:Name="TxtStatusSub" Text="" FontSize="24" Opacity="0.92"/>
                <TextBlock Grid.Row="2" x:Name="TxtStatusAssunto" Text="" FontSize="30" FontWeight="SemiBold" TextWrapping="Wrap" TextTrimming="CharacterEllipsis" MaxHeight="82" Margin="0,14,0,0"/>
                <TextBlock Grid.Row="3" x:Name="TxtStatusOrg" Text="" FontSize="20" Opacity="0.85" Margin="0,4,0,0"/>
                <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,18,0,0">
                  <Button x:Name="BtnReservar" Style="{StaticResource BtnGhost}" Visibility="Collapsed" Margin="0,0,14,0">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Style="{StaticResource Icon}" Text="&#xE787;" FontSize="24" Margin="0,0,12,0"/>
                      <TextBlock Text="Reservar"/>
                    </StackPanel>
                  </Button>
                  <Button x:Name="BtnEntrar" Style="{StaticResource Btn}">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Style="{StaticResource Icon}" Text="&#xE714;" FontSize="26" Margin="0,0,12,0"/>
                      <TextBlock Text="Entrar no Teams"/>
                    </StackPanel>
                  </Button>
                </StackPanel>
              </Grid>
            </Border>

            <TextBlock Grid.Row="1" Text="PRÓXIMAS REUNIÕES" FontSize="17" FontWeight="SemiBold" Opacity="0.65" Margin="8,24,0,10"/>

            <ScrollViewer Grid.Row="2" VerticalScrollBarVisibility="Hidden" HorizontalScrollBarVisibility="Disabled" PanningMode="VerticalOnly">
              <StackPanel x:Name="Lista"/>
            </ScrollViewer>

            <Grid Grid.Row="3" Margin="8,12,0,0">
              <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                <Ellipse x:Name="DotConexao" Width="10" Height="10" Fill="#9CA3AF" Margin="0,1,10,0"/>
                <TextBlock x:Name="TxtRodape" Text="Conectando ao Outlook…" FontSize="16" Opacity="0.7" VerticalAlignment="Center"/>
              </StackPanel>
              <Button x:Name="BtnConfig" HorizontalAlignment="Right" Style="{StaticResource BtnGhost}" Background="Transparent" MinHeight="44" Padding="12,0">
                <TextBlock Style="{StaticResource Icon}" Text="&#xE713;" FontSize="22" Opacity="0.8"/>
              </Button>
            </Grid>
          </Grid>
        </Border>
      </Grid>
    </Viewbox>

    <!-- Painel de opções -->
    <Grid x:Name="Overlay" Background="#D0000000" Visibility="Collapsed">
      <Border Background="#0F2F28" CornerRadius="20" Padding="36" Width="520" HorizontalAlignment="Center" VerticalAlignment="Center">
        <StackPanel>
          <TextBlock x:Name="TxtOverlayTitulo" Text="Painel de Sala" FontSize="30" FontWeight="Bold"/>
          <TextBlock x:Name="TxtInfo" Text="" FontSize="15" Opacity="0.8" TextWrapping="Wrap" Margin="0,8,0,22"/>
          <Button x:Name="BtnPersonalizar" Style="{StaticResource BtnGhost}" Margin="0,0,0,12">
            <StackPanel Orientation="Horizontal">
              <TextBlock Style="{StaticResource Icon}" Text="&#xE790;" FontSize="22" Margin="0,0,12,0"/>
              <TextBlock Text="Personalizar (nome, fundo, logo)"/>
            </StackPanel>
          </Button>
          <Button x:Name="BtnAtualizar" Style="{StaticResource BtnGhost}" Content="Atualizar agora" Margin="0,0,0,12"/>
          <Button x:Name="BtnTela" Style="{StaticResource BtnGhost}" Content="Sair da tela cheia" Margin="0,0,0,12"/>
          <Button x:Name="BtnAbrirConfig" Style="{StaticResource BtnGhost}" Content="Abrir config.json" Margin="0,0,0,12"/>
          <Button x:Name="BtnSair" Style="{StaticResource BtnGhost}" Content="Fechar o painel" Margin="0,0,0,12"/>
          <Button x:Name="BtnVoltar" Style="{StaticResource Btn}" Content="Voltar"/>
        </StackPanel>
      </Border>
    </Grid>

    <!-- Personalização -->
    <Grid x:Name="OverlayPers" Background="#D0000000" Visibility="Collapsed">
      <Border Background="#0F2F28" CornerRadius="20" Padding="36" Width="660" HorizontalAlignment="Center" VerticalAlignment="Center">
        <StackPanel>
          <TextBlock Text="Personalizar" FontSize="30" FontWeight="Bold"/>
          <TextBlock Text="As alterações aparecem na hora e ficam salvas no config.json." FontSize="15" Opacity="0.7" Margin="0,6,0,22"/>

          <TextBlock Style="{StaticResource Rotulo}" Text="Nome da sala"/>
          <TextBox x:Name="TxtPersNome" Style="{StaticResource Campo}"/>

          <TextBlock Style="{StaticResource Rotulo}" Text="Subtítulo (opcional — deixe vazio para ocultar)"/>
          <TextBox x:Name="TxtPersSub" Style="{StaticResource Campo}"/>

          <TextBlock Style="{StaticResource Rotulo}" Text="Imagem de fundo"/>
          <Grid Margin="0,6,0,16">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" x:Name="TxtPersFundo" Text="" FontSize="17" Opacity="0.85" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Margin="2,0,8,0"/>
            <Button Grid.Column="1" x:Name="BtnPersFundo" Style="{StaticResource BtnMini}" Content="Escolher…"/>
            <Button Grid.Column="2" x:Name="BtnPersFundoPadrao" Style="{StaticResource BtnMini}" Content="Padrão"/>
            <Button Grid.Column="3" x:Name="BtnPersFundoNenhum" Style="{StaticResource BtnMini}" Content="Nenhuma"/>
          </Grid>

          <TextBlock Style="{StaticResource Rotulo}" Text="Logo (PNG com fundo transparente fica melhor)"/>
          <Grid Margin="0,6,0,26">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" x:Name="TxtPersLogo" Text="" FontSize="17" Opacity="0.85" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Margin="2,0,8,0"/>
            <Button Grid.Column="1" x:Name="BtnPersLogo" Style="{StaticResource BtnMini}" Content="Escolher…"/>
            <Button Grid.Column="2" x:Name="BtnPersLogoPadrao" Style="{StaticResource BtnMini}" Content="Padrão"/>
            <Button Grid.Column="3" x:Name="BtnPersLogoNenhum" Style="{StaticResource BtnMini}" Content="Sem logo"/>
          </Grid>

          <TextBlock x:Name="TxtPersAviso" Text="" FontSize="15" Foreground="#FCA5A5" TextWrapping="Wrap" Visibility="Collapsed" Margin="0,0,0,14"/>

          <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="BtnPersCancelar" Style="{StaticResource BtnGhost}" Content="Cancelar" Margin="0,0,14,0"/>
            <Button x:Name="BtnPersSalvar" Style="{StaticResource Btn}" Content="Salvar"/>
          </StackPanel>
        </StackPanel>
      </Border>
    </Grid>
  </Grid>
</Window>
'@

$window = [System.Windows.Markup.XamlReader]::Parse($xaml)
$UI = @{}
foreach ($n in @('Root','BgLayer','LedLeft','LedRight','TxtHora','TxtData','TxtSala','TxtSubtitulo','ImgLogo','CardStatus',
                 'TxtStatus','TxtStatusSub','TxtStatusAssunto','TxtStatusOrg','BtnReservar','BtnEntrar','Lista',
                 'DotConexao','TxtRodape','BtnConfig','Overlay','TxtInfo','BtnPersonalizar','BtnAtualizar','BtnTela','BtnAbrirConfig','BtnSair','BtnVoltar',
                 'OverlayPers','TxtPersNome','TxtPersSub','TxtPersFundo','BtnPersFundo','BtnPersFundoPadrao','BtnPersFundoNenhum',
                 'TxtPersLogo','BtnPersLogo','BtnPersLogoPadrao','BtnPersLogoNenhum','TxtPersAviso','BtnPersCancelar','BtnPersSalvar')) {
  $UI[$n] = $window.FindName($n)
}

# --- helpers visuais ---------------------------------------------------------
$script:Cores = @{
  livre   = @{ Card = '#1FA84A'; Led = '#22C55E' }
  breve   = @{ Card = '#D97706'; Led = '#F59E0B' }
  ocupado = @{ Card = '#C62828'; Led = '#EF4444' }
}
function Get-Brush { param([string]$Hex) return [System.Windows.Media.BrushConverter]::new().ConvertFromString($Hex) }
function New-Thickness { param([double]$L, [double]$T, [double]$R, [double]$B) return New-Object System.Windows.Thickness($L, $T, $R, $B) }
function New-Text {
  param([string]$Text, [double]$Size = 20, [string]$Weight = 'Normal', [double]$Opacity = 1, [string]$Color = 'White', [switch]$Trim, [switch]$Wrap)
  $t = New-Object System.Windows.Controls.TextBlock
  $t.Text = $Text; $t.FontSize = $Size; $t.Opacity = $Opacity
  $t.FontWeight = [System.Windows.FontWeights]::$Weight
  $t.Foreground = Get-Brush $Color
  $t.VerticalAlignment = 'Center'
  if ($Trim) { $t.TextTrimming = 'CharacterEllipsis' }
  if ($Wrap) { $t.TextWrapping = 'Wrap' }
  return $t
}
function New-Icon { param([string]$Glyph, [double]$Size = 22, [double]$Opacity = 1, [string]$Color = 'White')
  $t = New-Text -Text $Glyph -Size $Size -Opacity $Opacity -Color $Color
  $t.FontFamily = New-Object System.Windows.Media.FontFamily('Segoe MDL2 Assets')
  return $t
}
function New-RowGrid {
  $g = New-Object System.Windows.Controls.Grid
  $conv = New-Object System.Windows.GridLengthConverter
  foreach ($w in @('150', '*', 'Auto')) {
    $cd = New-Object System.Windows.Controls.ColumnDefinition
    $cd.Width = $conv.ConvertFromString($w)
    $g.ColumnDefinitions.Add($cd) | Out-Null
  }
  return $g
}
function New-RowBorder { param([string]$Bg)
  $b = New-Object System.Windows.Controls.Border
  $b.Background = Get-Brush $Bg
  $b.CornerRadius = New-Object System.Windows.CornerRadius(14)
  $b.Padding = New-Thickness 18 12 18 12
  $b.Margin = New-Thickness 0 0 0 10
  $b.MinHeight = 76
  return $b
}
function Add-Child { param($Panel, $Child, [int]$Col = -1)
  if ($Col -ge 0) { [System.Windows.Controls.Grid]::SetColumn($Child, $Col) }
  $Panel.Children.Add($Child) | Out-Null
}

function New-DayHeader { param([string]$Texto)
  $t = New-Text -Text $Texto.ToUpper() -Size 15 -Weight 'SemiBold' -Opacity 0.55
  $t.Margin = New-Thickness 8 8 0 8
  return $t
}
function New-FreeRow { param([datetime]$De, [datetime]$Ate)
  $b = New-RowBorder '#14FFFFFF'
  $g = New-RowGrid
  Add-Child $g (New-Text -Text ($De.ToString('HH:mm') + ' – ' + $Ate.ToString('HH:mm')) -Size 20 -Opacity 0.6) 0
  Add-Child $g (New-Text -Text 'Livre' -Size 22 -Opacity 0.6) 1
  $b.Child = $g
  return $b
}
function New-MeetingRow { param($M, [datetime]$Agora)
  $b = New-RowBorder '#26FFFFFF'
  $g = New-RowGrid

  $hora = New-Object System.Windows.Controls.StackPanel
  $hora.VerticalAlignment = 'Center'
  Add-Child $hora (New-Text -Text $M.Inicio.ToString('HH:mm') -Size 27 -Weight 'SemiBold')
  Add-Child $hora (New-Text -Text $M.Fim.ToString('HH:mm') -Size 18 -Opacity 0.65)
  Add-Child $g $hora 0

  $info = New-Object System.Windows.Controls.StackPanel
  $info.VerticalAlignment = 'Center'
  $info.Margin = New-Thickness 0 0 16 0
  Add-Child $info (New-Text -Text $M.Assunto -Size 24 -Weight 'SemiBold' -Trim)
  $linha2 = @()
  if ($M.Organizador) { $linha2 += $M.Organizador }
  if ($M.Local -and $M.Local -notmatch 'teams' -and $M.Local -notmatch '^https?://') { $linha2 += $M.Local }
  if ($linha2.Count -gt 0) { Add-Child $info (New-Text -Text ($linha2 -join ' · ') -Size 17 -Opacity 0.7 -Trim) }
  Add-Child $g $info 1

  if (Test-PodeEntrar -Reuniao $M -Agora $Agora -MinutosAntes ([int]$Cfg.MinutosAntesParaEntrar)) {
    $btn = New-Object System.Windows.Controls.Button
    $btn.Style = $window.FindResource('BtnSmall')
    $sp = New-Object System.Windows.Controls.StackPanel; $sp.Orientation = 'Horizontal'
    $ic = New-Icon -Glyph ([char]0xE714) -Size 22 -Color '#0B2F27'; $ic.Margin = New-Thickness 0 0 10 0
    Add-Child $sp $ic
    Add-Child $sp (New-Text -Text 'Entrar' -Size 20 -Weight 'SemiBold' -Color '#0B2F27')
    $btn.Content = $sp
    $btn.Tag = [string]$M.Link
    $btn.Add_Click({ param($s, $e) Open-Teams -Url ([string]$s.Tag) })
    Add-Child $g $btn 2
  } elseif ($M.Link) {
    $ic = New-Icon -Glyph ([char]0xE714) -Size 24 -Opacity 0.45
    $ic.Margin = New-Thickness 8 0 8 0
    $ic.ToolTip = 'Reunião do Teams'
    Add-Child $g $ic 2
  }
  $b.Child = $g
  return $b
}

function Set-Footer { param([string]$Texto, [string]$Nivel = 'ok')
  $UI.TxtRodape.Text = $Texto
  $cor = switch ($Nivel) { 'ok' { '#22C55E' } 'warn' { '#F59E0B' } 'err' { '#EF4444' } default { '#9CA3AF' } }
  $UI.DotConexao.Fill = Get-Brush $cor
}

# --- montagem da tela --------------------------------------------------------
function Update-View {
  $agora = Get-Date
  $UI.TxtHora.Text = $agora.ToString('HH:mm')
  $d = $agora.ToString("dddd, d 'de' MMMM 'de' yyyy", $script:PtBR)
  $UI.TxtData.Text = $d.Substring(0, 1).ToUpper() + $d.Substring(1)

  $minAntes = [int]$Cfg.MinutosAntesParaEntrar
  $estado = Get-Estado -Reunioes $script:Meetings -Agora $agora -MinutosAntes $minAntes
  $cores = $script:Cores[$estado.Tipo]
  $UI.CardStatus.Background = Get-Brush $cores.Card
  $UI.LedLeft.Fill = Get-Brush $cores.Led
  $UI.LedRight.Fill = Get-Brush $cores.Led

  $destaque = $null
  switch ($estado.Tipo) {
    'ocupado' {
      $destaque = $estado.Atual
      $UI.TxtStatus.Text = 'Em reunião'
      $UI.TxtStatusSub.Text = $destaque.Inicio.ToString('HH:mm') + ' – ' + $destaque.Fim.ToString('HH:mm')
    }
    'breve' {
      $destaque = $estado.EmBreve
      $min = [math]::Max(1, [math]::Ceiling(($destaque.Inicio - $agora).TotalMinutes))
      $UI.TxtStatus.Text = if ($min -eq 1) { 'Começa em 1 min' } else { "Começa em $min min" }
      $UI.TxtStatusSub.Text = $destaque.Inicio.ToString('HH:mm') + ' – ' + $destaque.Fim.ToString('HH:mm')
    }
    default {
      $UI.TxtStatus.Text = 'Disponível'
      $p = $estado.Proxima
      if (-not $p) { $UI.TxtStatusSub.Text = 'Sem reuniões nas próximas ' + [int]$Cfg.HorasAFrente + ' horas' }
      elseif ($p.Inicio.Date -eq $agora.Date) { $UI.TxtStatusSub.Text = 'até ' + $p.Inicio.ToString('HH:mm') }
      else { $UI.TxtStatusSub.Text = 'pelo resto do dia' }
    }
  }
  if ($destaque) {
    $UI.TxtStatusAssunto.Text = $destaque.Assunto
    $UI.TxtStatusAssunto.Visibility = 'Visible'
    $UI.TxtStatusOrg.Text = $destaque.Organizador
    $UI.TxtStatusOrg.Visibility = if ($destaque.Organizador) { 'Visible' } else { 'Collapsed' }
  } else {
    $UI.TxtStatusAssunto.Visibility = 'Collapsed'
    $UI.TxtStatusOrg.Visibility = 'Collapsed'
  }
  $script:LinkDestaque = if ($destaque) { $destaque.Link } else { $null }
  $UI.BtnEntrar.Visibility = if (Test-PodeEntrar -Reuniao $destaque -Agora $agora -MinutosAntes $minAntes) { 'Visible' } else { 'Collapsed' }
  $UI.BtnReservar.Visibility = if ([bool]$Cfg.PermitirReserva -and $estado.Tipo -eq 'livre' -and -not $script:Demo) { 'Visible' } else { 'Collapsed' }

  # lista
  $UI.Lista.Children.Clear()
  $itens = @($estado.Lista | Select-Object -First ([int]$Cfg.MaxItens))
  if ($itens.Count -eq 0) {
    $t = New-Text -Text 'Nenhuma reunião agendada.' -Size 22 -Opacity 0.55
    $t.Margin = New-Thickness 8 10 0 0
    Add-Child $UI.Lista $t
  }
  $gapMin = [int]$Cfg.IntervaloLivreMin
  $diaAnterior = $null
  $fimAnterior = $null
  if ($estado.Atual) { $fimAnterior = $estado.Atual.Fim } elseif ($estado.EmBreve) { $fimAnterior = $estado.EmBreve.Fim }
  foreach ($m in $itens) {
    $rot = Get-DiaRotulo -Data $m.Inicio -Agora $agora
    if ($rot -ne 'Hoje' -and $rot -ne $diaAnterior) { Add-Child $UI.Lista (New-DayHeader $rot); $fimAnterior = $null }
    if ($gapMin -gt 0 -and $fimAnterior -and $m.Inicio.Date -eq $agora.Date -and ($m.Inicio - $fimAnterior).TotalMinutes -ge $gapMin) {
      Add-Child $UI.Lista (New-FreeRow -De $fimAnterior -Ate $m.Inicio)
    }
    Add-Child $UI.Lista (New-MeetingRow -M $m -Agora $agora)
    $diaAnterior = $rot
    if ($m.Fim -gt $fimAnterior) { $fimAnterior = $m.Fim }
  }
}

function Refresh-Data {
  $script:LastRefresh = Get-Date
  if ($script:Demo) {
    $script:Meetings = @(Get-DemoMeetings)
    Set-Footer 'Modo demonstração — dados fictícios' 'warn'
    Update-View
    return
  }
  try {
    $script:Meetings = @(Read-Meetings)
    $fonte = if ($Cfg.CalendarioCompartilhado) { [string]$Cfg.CalendarioCompartilhado } else { 'Outlook' }
    Set-Footer ('Atualizado às ' + (Get-Date).ToString('HH:mm') + ' · ' + $fonte) 'ok'
  } catch {
    $script:OutlookNs = $null
    $msg = $_.Exception.Message -replace '\s+', ' '
    if ($msg.Length -gt 90) { $msg = $msg.Substring(0, 90) + '…' }
    Set-Footer ('Sem conexão com o Outlook — ' + $msg) 'err'
  }
  Update-View
}

function Set-Fullscreen { param([bool]$On)
  $script:Fullscreen = $On
  if ($On) {
    $window.WindowStyle = 'None'; $window.ResizeMode = 'NoResize'
    $window.WindowState = 'Maximized'; $window.Topmost = [bool]$Cfg.SempreNoTopo
    $UI.BtnTela.Content = 'Sair da tela cheia'
  } else {
    $window.Topmost = $false; $window.WindowState = 'Normal'
    $window.WindowStyle = 'SingleBorderWindow'; $window.ResizeMode = 'CanResize'
    $UI.BtnTela.Content = 'Tela cheia'
  }
}

function Show-Overlay { param([bool]$On)
  if ($On) {
    $info = @(
      "Versão $script:Versao · " + $(if ($script:Demo) { 'modo demonstração' } elseif ($Cfg.CalendarioCompartilhado) { 'agenda: ' + $Cfg.CalendarioCompartilhado } else { 'agenda pessoal do Outlook' })
      'Atalhos: Esc/F11 tela cheia · F5 atualizar · Ctrl+Q fechar'
      "Config: $Config"
    ) -join "`n"
    $UI.TxtInfo.Text = $info
  }
  $UI.Overlay.Visibility = if ($On) { 'Visible' } else { 'Collapsed' }
}

# --- personalização (nome, subtítulo, fundo, logo) ---------------------------
function Resolve-AppPath { param([string]$P) if (-not $P) { return $null }; if ([System.IO.Path]::IsPathRooted($P)) { return $P }; return (Join-Path $script:AppDir $P) }

function New-Bitmap { param([string]$Path)
  $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
  $bmp.BeginInit()
  $bmp.UriSource = New-Object System.Uri($Path)
  $bmp.CacheOption = 'OnLoad'                       # não trava o arquivo — pode ser substituído depois
  try { if ((Get-Item -LiteralPath $Path).Length -gt 1.5MB) { $bmp.DecodePixelWidth = 2560 } } catch {}   # fotos grandes: decodifica menor
  $bmp.EndInit()
  return $bmp
}

function Set-Background { param([string]$Caminho)
  # Caminho relativo à pasta do app ou absoluto; vazio/inexistente = fundo liso (verde escuro)
  try {
    $bg = Resolve-AppPath $Caminho
    if ($bg -and (Test-Path -LiteralPath $bg)) {
      $brush = New-Object System.Windows.Media.ImageBrush((New-Bitmap $bg)); $brush.Stretch = 'UniformToFill'
      $UI.BgLayer.Background = $brush
      return $true
    }
  } catch {}
  $UI.BgLayer.Background = $null
  return $false
}

function Set-Logo { param([string]$Caminho)
  try {
    $lg = Resolve-AppPath $Caminho
    if ($lg -and (Test-Path -LiteralPath $lg)) {
      $UI.ImgLogo.Source = New-Bitmap $lg
      $UI.ImgLogo.Visibility = 'Visible'
      return $true
    }
  } catch {}
  $UI.ImgLogo.Source = $null
  $UI.ImgLogo.Visibility = 'Collapsed'
  return $false
}

function Set-NomeSala { param([string]$Nome, [string]$Subtitulo)
  $UI.TxtSala.Text = $Nome
  $UI.TxtSubtitulo.Text = $Subtitulo
  $UI.TxtSubtitulo.Visibility = if ($Subtitulo) { 'Visible' } else { 'Collapsed' }
}

function Apply-Personalizacao {
  # Reaplica o que está salvo em $Cfg (usado na abertura e ao cancelar a edição)
  Set-NomeSala -Nome ([string]$Cfg.NomeSala) -Subtitulo ([string]$Cfg.Subtitulo)
  [void](Set-Background ([string]$Cfg.ImagemFundo))
  [void](Set-Logo ([string]$Cfg.Logo))
}

function Save-Config {
  # Grava TODAS as chaves (na ordem dos padrões), preservando o que já estava no config.json
  $o = [ordered]@{}
  foreach ($k in $Defaults.Keys) { $o[$k] = $Cfg[$k] }
  ($o | ConvertTo-Json -Depth 3) | Set-Content -LiteralPath $Config -Encoding UTF8
}

function Import-Asset {
  # Copia a imagem escolhida para a pasta do app (assim o painel não depende de arquivos em Downloads/pendrive)
  # e devolve o valor a gravar no config: nome relativo (preferido) ou caminho absoluto se não der para copiar.
  param([string]$Caminho, [string]$NomeBase)
  if (-not $Caminho) { return '' }
  if (-not [System.IO.Path]::IsPathRooted($Caminho)) { return $Caminho }                 # já é relativo (padrão)
  $dir = [System.IO.Path]::GetDirectoryName($Caminho).TrimEnd('\')
  if ($dir -ieq $script:AppDir.TrimEnd('\')) { return [System.IO.Path]::GetFileName($Caminho) }   # já está na pasta do app
  try {
    $ext = [System.IO.Path]::GetExtension($Caminho).ToLower()
    $destNome = $NomeBase + $ext
    $dest = Join-Path $script:AppDir $destNome
    Copy-Item -LiteralPath $Caminho -Destination $dest -Force -ErrorAction Stop
    return $destNome
  } catch {
    return $Caminho
  }
}

function Select-ImageFile { param([string]$Titulo)
  $dlg = New-Object Microsoft.Win32.OpenFileDialog
  $dlg.Title = $Titulo
  $dlg.Filter = 'Imagens (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|Todos os arquivos (*.*)|*.*'
  try { $dlg.InitialDirectory = [Environment]::GetFolderPath('MyPictures') } catch {}
  $ok = $dlg.ShowDialog($window)
  if ($ok) { return $dlg.FileName }
  return $null
}

function Get-NomeExibicao { param([string]$Valor, [string]$Padrao, [string]$Nenhum)
  if (-not $Valor) { return $Nenhum }
  if ($Valor -eq $Padrao) { return 'Padrão (' + $Padrao + ')' }
  return ($Valor.Split([char[]]@('\', '/')))[-1]
}

function Update-PersLabels {
  $UI.TxtPersFundo.Text = Get-NomeExibicao -Valor $script:Pers.Fundo -Padrao ([string]$Defaults.ImagemFundo) -Nenhum 'Nenhuma (fundo liso)'
  $UI.TxtPersLogo.Text  = Get-NomeExibicao -Valor $script:Pers.Logo  -Padrao ([string]$Defaults.Logo)        -Nenhum 'Sem logo'
}

function Show-OverlayPers { param([bool]$On)
  if ($On) {
    Show-Overlay $false
    $script:Pers = @{ Fundo = [string]$Cfg.ImagemFundo; Logo = [string]$Cfg.Logo }
    $UI.TxtPersNome.Text = [string]$Cfg.NomeSala
    $UI.TxtPersSub.Text  = [string]$Cfg.Subtitulo
    $UI.TxtPersAviso.Visibility = 'Collapsed'
    Update-PersLabels
    $UI.OverlayPers.Visibility = 'Visible'
    $UI.TxtPersNome.Focus() | Out-Null
    $UI.TxtPersNome.SelectAll()
  } else {
    $UI.OverlayPers.Visibility = 'Collapsed'
  }
}

function Cancel-Personalizacao {
  Show-OverlayPers $false
  Apply-Personalizacao          # desfaz a pré-visualização
}

function Save-Personalizacao {
  $nome = $UI.TxtPersNome.Text.Trim()
  if (-not $nome) { $nome = [string]$Defaults.NomeSala }
  $Cfg.NomeSala    = $nome
  $Cfg.Subtitulo   = $UI.TxtPersSub.Text.Trim()
  $Cfg.ImagemFundo = Import-Asset -Caminho $script:Pers.Fundo -NomeBase 'fundo_personalizado'
  $Cfg.Logo        = Import-Asset -Caminho $script:Pers.Logo  -NomeBase 'logo_personalizado'
  try {
    Save-Config
  } catch {
    $UI.TxtPersAviso.Text = 'Não foi possível gravar o config.json: ' + $_.Exception.Message
    $UI.TxtPersAviso.Visibility = 'Visible'
    return
  }
  Apply-Personalizacao
  Show-OverlayPers $false
  Set-Footer 'Personalização salva' 'ok'
}

# --- aplica configuração à janela --------------------------------------------
try { $window.FontFamily = New-Object System.Windows.Media.FontFamily([string]$Cfg.Fonte + ', Segoe UI, Arial') } catch {}
Apply-Personalizacao
try {
  $ico = Join-Path $script:AppDir 'PainelSala.ico'
  if (Test-Path -LiteralPath $ico) { $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create((New-Object System.Uri($ico))) }
} catch {}

# --- eventos -----------------------------------------------------------------
$UI.BtnEntrar.Add_Click({ Open-Teams -Url $script:LinkDestaque })
$UI.BtnReservar.Add_Click({
  try { New-Reserva; Refresh-Data } catch { Set-Footer ('Não foi possível reservar — ' + $_.Exception.Message) 'err' }
})
$UI.BtnConfig.Add_Click({ Show-Overlay $true })
$UI.BtnVoltar.Add_Click({ Show-Overlay $false })
$UI.BtnAtualizar.Add_Click({ Show-Overlay $false; Refresh-Data })
$UI.BtnTela.Add_Click({ Show-Overlay $false; Set-Fullscreen (-not $script:Fullscreen) })
$UI.BtnAbrirConfig.Add_Click({
  Show-Overlay $false
  if (-not (Test-Path -LiteralPath $Config)) {
    ($Defaults | ConvertTo-Json -Depth 3) | Set-Content -LiteralPath $Config -Encoding UTF8
  }
  try { Start-Process notepad.exe -ArgumentList ('"' + $Config + '"') } catch {}
})
$UI.BtnSair.Add_Click({ $window.Close() })
$UI.Overlay.Add_MouseDown({ param($s, $e) if ($e.OriginalSource -eq $s) { Show-Overlay $false } })

# personalização
$UI.BtnPersonalizar.Add_Click({ Show-OverlayPers $true })
$UI.BtnPersCancelar.Add_Click({ Cancel-Personalizacao })
$UI.BtnPersSalvar.Add_Click({ Save-Personalizacao })
$UI.OverlayPers.Add_MouseDown({ param($s, $e) if ($e.OriginalSource -eq $s) { Cancel-Personalizacao } })
$UI.TxtPersNome.Add_TextChanged({ Set-NomeSala -Nome $UI.TxtPersNome.Text -Subtitulo $UI.TxtPersSub.Text })   # pré-visualização ao digitar
$UI.TxtPersSub.Add_TextChanged({ Set-NomeSala -Nome $UI.TxtPersNome.Text -Subtitulo $UI.TxtPersSub.Text })
$UI.BtnPersFundo.Add_Click({
  $f = Select-ImageFile -Titulo 'Escolher imagem de fundo'
  if ($f) {
    if (Set-Background $f) { $script:Pers.Fundo = $f; $UI.TxtPersAviso.Visibility = 'Collapsed' }
    else { $UI.TxtPersAviso.Text = 'Não foi possível abrir essa imagem.'; $UI.TxtPersAviso.Visibility = 'Visible'; [void](Set-Background $script:Pers.Fundo) }
    Update-PersLabels
  }
})
$UI.BtnPersFundoPadrao.Add_Click({ $script:Pers.Fundo = [string]$Defaults.ImagemFundo; [void](Set-Background $script:Pers.Fundo); Update-PersLabels })
$UI.BtnPersFundoNenhum.Add_Click({ $script:Pers.Fundo = ''; [void](Set-Background ''); Update-PersLabels })
$UI.BtnPersLogo.Add_Click({
  $f = Select-ImageFile -Titulo 'Escolher logo'
  if ($f) {
    if (Set-Logo $f) { $script:Pers.Logo = $f; $UI.TxtPersAviso.Visibility = 'Collapsed' }
    else { $UI.TxtPersAviso.Text = 'Não foi possível abrir essa imagem.'; $UI.TxtPersAviso.Visibility = 'Visible'; [void](Set-Logo $script:Pers.Logo) }
    Update-PersLabels
  }
})
$UI.BtnPersLogoPadrao.Add_Click({ $script:Pers.Logo = [string]$Defaults.Logo; [void](Set-Logo $script:Pers.Logo); Update-PersLabels })
$UI.BtnPersLogoNenhum.Add_Click({ $script:Pers.Logo = ''; [void](Set-Logo ''); Update-PersLabels })

$window.Add_KeyDown({
  param($s, $e)
  switch ($e.Key) {
    'Escape' {
      if ($UI.OverlayPers.Visibility -eq 'Visible') { Cancel-Personalizacao }
      elseif ($UI.Overlay.Visibility -eq 'Visible') { Show-Overlay $false }
      else { Set-Fullscreen (-not $script:Fullscreen) }
    }
    'F11'    { Set-Fullscreen (-not $script:Fullscreen) }
    'F5'     { Refresh-Data }
    'Q'      { if ($e.KeyboardDevice.Modifiers -eq [System.Windows.Input.ModifierKeys]::Control) { $window.Close() } }
  }
})

$window.Add_SourceInitialized({
  try {
    $telas = [System.Windows.Forms.Screen]::AllScreens
    $i = [int]$Cfg.Monitor
    if ($i -gt 0 -and $i -lt $telas.Count) {
      $b = $telas[$i].Bounds
      $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($window)).Handle
      [void][Native.Win32]::SetWindowPos($hwnd, [IntPtr]::Zero, $b.X, $b.Y, $b.Width, $b.Height, 0x0004 -bor 0x0010)
    }
  } catch {}
  Set-Fullscreen $script:FullscreenWanted
})

$script:Meetings = @()
$script:LastRefresh = [datetime]::MinValue
$script:LastMinute = -1

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [timespan]::FromSeconds(1)
$timer.Add_Tick({
  $agora = Get-Date
  $UI.TxtHora.Text = $agora.ToString('HH:mm')
  if (($agora - $script:LastRefresh).TotalSeconds -ge [double]$Cfg.IntervaloAtualizacaoSeg) { Refresh-Data; $script:LastMinute = $agora.Minute }
  elseif ($agora.Minute -ne $script:LastMinute) { $script:LastMinute = $agora.Minute; Update-View }
})

$window.Add_ContentRendered({
  if ([bool]$Cfg.ManterTelaLigada) { try { [void][Native.Win32]::SetThreadExecutionState(0x80000003) } catch {} }
  Update-View
  $timer.Start()
  $window.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{ Refresh-Data }) | Out-Null
})
$window.Add_Closed({
  $timer.Stop()
  try { [void][Native.Win32]::SetThreadExecutionState(0x80000000) } catch {}
})

[void]$window.ShowDialog()
