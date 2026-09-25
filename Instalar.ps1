#Requires -Version 5.1
<#
  Instalar.ps1 — cria os atalhos do Painel de Sala (área de trabalho + iniciar com o Windows)
  Uso: Instalar.cmd            -> instala (atalho na área de trabalho e inicialização automática)
       Instalar.cmd -SemInicio -> instala só o atalho da área de trabalho
       Desinstalar.cmd         -> remove os atalhos (não apaga a pasta)
#>
param(
  [switch]$SemInicio,
  [switch]$Remover
)
$ErrorActionPreference = 'Stop'
$app  = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ps1  = Join-Path $app 'PainelSala.ps1'
$ico  = Join-Path $app 'PainelSala.ico'
$psExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
$nome = 'Painel de Sala'

$desktop = [Environment]::GetFolderPath('Desktop')
$startup = [Environment]::GetFolderPath('Startup')
$lnkDesktop = Join-Path $desktop  "$nome.lnk"
$lnkDemo    = Join-Path $desktop  "$nome (Demo).lnk"
$lnkStartup = Join-Path $startup  "$nome.lnk"

if ($Remover) {
  foreach ($f in @($lnkDesktop, $lnkDemo, $lnkStartup)) {
    if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force; Write-Host "Removido: $f" }
  }
  Write-Host 'Atalhos removidos. A pasta do painel foi mantida.'
  return
}

if (-not (Test-Path -LiteralPath $ps1)) { throw "Arquivo nao encontrado: $ps1" }

# Remove a marca "arquivo baixado da internet" de todos os arquivos da pasta
Get-ChildItem -LiteralPath $app -File | ForEach-Object { try { Unblock-File -LiteralPath $_.FullName } catch {} }

$ws = New-Object -ComObject WScript.Shell
function New-Atalho {
  param([string]$Caminho, [string]$ArgsExtra = '', [string]$Descricao = 'Painel de Sala de Reunioes')
  $s = $ws.CreateShortcut($Caminho)
  $s.TargetPath = $psExe
  $s.Arguments = ('-NoLogo -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "{0}" {1}' -f $ps1, $ArgsExtra).Trim()
  $s.WorkingDirectory = $app
  if (Test-Path -LiteralPath $ico) { $s.IconLocation = "$ico,0" }
  $s.WindowStyle = 7   # minimizado (o console some em seguida)
  $s.Description = $Descricao
  $s.Save()
  Write-Host "Criado: $Caminho"
}

New-Atalho -Caminho $lnkDesktop
New-Atalho -Caminho $lnkDemo -ArgsExtra '-Demo' -Descricao 'Painel de Sala (dados de demonstracao)'
if (-not $SemInicio) { New-Atalho -Caminho $lnkStartup } elseif (Test-Path -LiteralPath $lnkStartup) { Remove-Item -LiteralPath $lnkStartup -Force }

# config.json inicial (se ainda nao existir)
$cfg = Join-Path $app 'config.json'
if (-not (Test-Path -LiteralPath $cfg)) {
  @'
{
  "NomeSala": "Sala de Reuniões",
  "Subtitulo": "",
  "CalendarioCompartilhado": "",
  "HorasAFrente": 30,
  "MaxItens": 6,
  "IntervaloAtualizacaoSeg": 60,
  "MinutosAntesParaEntrar": 15,
  "IntervaloLivreMin": 30,
  "MostrarOrganizador": true,
  "TelaCheia": true,
  "SempreNoTopo": false,
  "Monitor": 0,
  "ManterTelaLigada": true,
  "AbrirVia": "msteams",
  "PermitirReserva": false,
  "DuracaoReservaMin": 30,
  "Fonte": "Segoe UI",
  "ImagemFundo": "fundo.jpg",
  "Logo": "logo.png",
  "Demo": false
}
'@ | Set-Content -LiteralPath $cfg -Encoding UTF8
  Write-Host "Criado: $cfg"
}

Write-Host ''
Write-Host 'Instalacao concluida.'
Write-Host " - Atalho na area de trabalho: $nome"
if (-not $SemInicio) { Write-Host ' - O painel abre automaticamente ao entrar no Windows.' }
Write-Host ' - Para testar sem Outlook, use o atalho "Painel de Sala (Demo)".'
