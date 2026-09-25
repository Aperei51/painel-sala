@echo off
rem Abre o painel sem instalar nada. Aceita os mesmos parametros do script: -Demo  -Janela
start "" powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0PainelSala.ps1" %*
