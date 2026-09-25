#Requires -Version 5.1
<#
  Empacotar-MSIX.ps1 — gera o pacote .msix do Painel de Sala para envio à Microsoft Store.

  O que faz:
    1. Lê store.json (identidade do app copiada do Partner Center)
    2. Compila o launcher PainelSala.exe com o csc.exe que já vem no Windows (.NET Framework 4.x)
    3. Monta a pasta do pacote (script, imagens, config padrão, ícones, manifesto preenchido)
    4. Gera dist\PainelSala_<versão>.msix com o MakeAppx.exe do Windows SDK
    5. (opcional, -TestarLocal) assina com certificado autoassinado e instala neste PC para testar

  Requisito único: "Windows SDK Signing Tools for Desktop Apps" (MakeAppx.exe / SignTool.exe).
    Instale pelo instalador do Windows SDK marcando SÓ esse componente (alguns MB):
    https://developer.microsoft.com/windows/downloads/windows-sdk/
    (ou: winget install Microsoft.WindowsSDK.10.0.26100 — instala o SDK completo)

  Uso:  Empacotar-MSIX.cmd            -> gera o .msix para enviar à Store
        Empacotar-MSIX.cmd -TestarLocal -> gera, assina e instala aqui (pede admin uma vez p/ confiar no certificado)
#>
[CmdletBinding()]
param(
  [switch]$TestarLocal
)
$ErrorActionPreference = 'Stop'
$root   = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }   # pasta store\
$appDir = Split-Path -Parent $root                                                                       # pasta do app (PainelSala.ps1)
$build  = Join-Path $root 'build'
$pkg    = Join-Path $build 'pkg'
$dist   = Join-Path $root 'dist'

function Passo([string]$t) { Write-Host ''; Write-Host "==> $t" -ForegroundColor Cyan }
function Falha([string]$t) { Write-Host ''; Write-Host "ERRO: $t" -ForegroundColor Red; exit 1 }

# ---------------------------------------------------------------- 1. store.json
Passo 'Lendo store.json'
$storeJson = Join-Path $root 'store.json'
if (-not (Test-Path -LiteralPath $storeJson)) { Falha "store.json não encontrado em $root" }
$st = Get-Content -LiteralPath $storeJson -Raw -Encoding UTF8 | ConvertFrom-Json
foreach ($k in 'IdentityName', 'Publisher', 'PublisherDisplayName', 'DisplayName', 'Version') {
  $v = [string]$st.$k
  if (-not $v -or $v -like '*PREENCHA*') { Falha "Preencha o campo '$k' no store.json com o valor do Partner Center (Identidade do produto)." }
}
if ($st.Version -notmatch '^\d+\.\d+\.\d+\.\d+$') { Falha "Version deve ter 4 números, ex.: 1.3.0.0 (está '$($st.Version)')" }
if ($st.Publisher -notmatch '^CN=') { Falha "Publisher deve começar com 'CN=' (copie exatamente do Partner Center)" }
Write-Host "  $($st.DisplayName)  v$($st.Version)  |  $($st.IdentityName)  |  $($st.Publisher)"

# ---------------------------------------------------------------- 2. ferramentas
Passo 'Localizando ferramentas'
$csc = Get-ChildItem "$env:WINDIR\Microsoft.NET\Framework64\v4*\csc.exe", "$env:WINDIR\Microsoft.NET\Framework\v4*\csc.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $csc) { Falha 'csc.exe (.NET Framework 4.x) não encontrado — ele vem com o Windows; verifique se o .NET Framework 4.8 está instalado.' }
function Find-SdkTool([string]$exe) {
  $cands = @()
  foreach ($base in @("${env:ProgramFiles(x86)}\Windows Kits\10\bin", "$env:ProgramFiles\Windows Kits\10\bin")) {
    if (Test-Path $base) { $cands += Get-ChildItem -Path $base -Recurse -Filter $exe -ErrorAction SilentlyContinue | Where-Object { $_.FullName -match '\\(x64|x86)\\' } }
  }
  $cands | Sort-Object FullName -Descending | Where-Object { $_.FullName -match '\\x64\\' } | Select-Object -First 1
}
$makeappx = Find-SdkTool 'makeappx.exe'
$signtool = Find-SdkTool 'signtool.exe'
if (-not $makeappx) {
  Falha @'
MakeAppx.exe não encontrado. Instale o componente "Windows SDK Signing Tools for Desktop Apps":
  1. Baixe o instalador do Windows SDK: https://developer.microsoft.com/windows/downloads/windows-sdk/
  2. Na tela de componentes, desmarque tudo e marque apenas "Windows SDK Signing Tools for Desktop Apps"
  3. Conclua e rode este script de novo.
'@
}
Write-Host "  csc:      $($csc.FullName)"
Write-Host "  makeappx: $($makeappx.FullName)"
if ($signtool) { Write-Host "  signtool: $($signtool.FullName)" }

# ---------------------------------------------------------------- 3. pasta do pacote
Passo 'Montando a pasta do pacote'
if (Test-Path $build) { Remove-Item -Recurse -Force $build }
New-Item -ItemType Directory -Path $pkg, (Join-Path $pkg 'Assets'), $dist -Force | Out-Null

$arquivosApp = @('PainelSala.ps1', 'config.json', 'fundo.jpg', 'logo.png', 'LEIA-ME.txt')
foreach ($f in $arquivosApp) {
  $src = Join-Path $appDir $f
  if (-not (Test-Path -LiteralPath $src)) { Falha "Arquivo do app não encontrado: $src" }
  Copy-Item -LiteralPath $src -Destination $pkg
}
Copy-Item -Path (Join-Path $root 'Assets\*.png') -Destination (Join-Path $pkg 'Assets')

# ---------------------------------------------------------------- 4. launcher
Passo 'Compilando PainelSala.exe (launcher)'
$exe = Join-Path $pkg 'PainelSala.exe'
$ico = Join-Path $root 'PainelSala.ico'
$cscArgs = @('/nologo', '/target:winexe', '/optimize+', "/out:`"$exe`"", '/reference:System.Windows.Forms.dll', "/win32icon:`"$ico`"", "`"$(Join-Path $root 'Launcher.cs')`"")
$p = Start-Process -FilePath $csc.FullName -ArgumentList $cscArgs -NoNewWindow -Wait -PassThru
if ($p.ExitCode -ne 0 -or -not (Test-Path $exe)) { Falha 'Falha ao compilar o launcher.' }

# ---------------------------------------------------------------- 5. manifesto
Passo 'Gerando AppxManifest.xml'
$tpl = Get-Content -LiteralPath (Join-Path $root 'AppxManifest.template.xml') -Raw -Encoding UTF8
$man = $tpl.Replace('{{IdentityName}}', [string]$st.IdentityName).
            Replace('{{Publisher}}', [string]$st.Publisher).
            Replace('{{PublisherDisplayName}}', [System.Security.SecurityElement]::Escape([string]$st.PublisherDisplayName)).
            Replace('{{DisplayName}}', [System.Security.SecurityElement]::Escape([string]$st.DisplayName)).
            Replace('{{Version}}', [string]$st.Version)
[IO.File]::WriteAllText((Join-Path $pkg 'AppxManifest.xml'), $man, (New-Object System.Text.UTF8Encoding($false)))

# ---------------------------------------------------------------- 6. makeappx
Passo 'Gerando o pacote .msix'
$msix = Join-Path $dist ("PainelSala_" + $st.Version + ".msix")
if (Test-Path $msix) { Remove-Item -Force $msix }
$p = Start-Process -FilePath $makeappx.FullName -ArgumentList @('pack', '/d', "`"$pkg`"", '/p', "`"$msix`"", '/o') -NoNewWindow -Wait -PassThru
if ($p.ExitCode -ne 0 -or -not (Test-Path $msix)) { Falha 'MakeAppx falhou (veja as mensagens acima).' }
Write-Host ''
Write-Host "PACOTE PRONTO: $msix" -ForegroundColor Green
Write-Host '  -> Envie este arquivo no Partner Center, em Pacotes. A Store assina o pacote por você; não precisa assinar para enviar.'

# ---------------------------------------------------------------- 7. teste local (opcional)
if ($TestarLocal) {
  Passo 'Teste local: assinando com certificado autoassinado e instalando'
  if (-not $signtool) { Falha 'SignTool.exe não encontrado (vem junto com o MakeAppx no mesmo componente do SDK).' }
  $cn = [string]$st.Publisher
  $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq $cn } | Select-Object -First 1
  if (-not $cert) {
    $cert = New-SelfSignedCertificate -Type Custom -Subject $cn -KeyUsage DigitalSignature -FriendlyName 'Painel de Sala (teste local)' `
              -CertStoreLocation 'Cert:\CurrentUser\My' -TextExtension @('2.5.29.37={text}1.3.6.1.5.5.7.3.3', '2.5.29.19={text}')
  }
  $pfx = Join-Path $build 'teste-local.pfx'
  $senha = ConvertTo-SecureString -String 'painel' -Force -AsPlainText
  Export-PfxCertificate -Cert $cert -FilePath $pfx -Password $senha | Out-Null
  $p = Start-Process -FilePath $signtool.FullName -ArgumentList @('sign', '/fd', 'SHA256', '/a', '/f', "`"$pfx`"", '/p', 'painel', "`"$msix`"") -NoNewWindow -Wait -PassThru
  if ($p.ExitCode -ne 0) { Falha 'SignTool falhou.' }
  # o Windows só instala pacotes de certificados confiáveis: importa em "Pessoas confiáveis" da máquina (pede admin)
  $cer = Join-Path $build 'teste-local.cer'
  Export-Certificate -Cert $cert -FilePath $cer | Out-Null
  Write-Host '  Confirmando o certificado de teste como confiável (janela de administrador)...'
  Start-Process powershell.exe -Verb RunAs -Wait -ArgumentList @('-NoProfile', '-Command', "Import-Certificate -FilePath '$cer' -CertStoreLocation Cert:\LocalMachine\TrustedPeople | Out-Null")
  Add-AppxPackage -Path $msix
  Write-Host ''
  Write-Host 'INSTALADO. Procure "Painel de Sala" no menu Iniciar. Para remover: Configurações > Aplicativos.' -ForegroundColor Green
  Write-Host '  (o certificado de teste serve só para este PC; o pacote enviado à Store é assinado pela Microsoft)'
}
