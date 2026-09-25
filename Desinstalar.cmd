@echo off
title Painel de Sala - Remover atalhos
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Instalar.ps1" -Remover
echo.
pause
