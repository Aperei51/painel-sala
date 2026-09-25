@echo off
rem Abre o painel com dados ficticios (nao precisa de Outlook)
start "" powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0PainelSala.ps1" -Demo
