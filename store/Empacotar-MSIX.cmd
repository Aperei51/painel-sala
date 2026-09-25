@echo off
title Painel de Sala - gerar pacote MSIX para a Microsoft Store
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Empacotar-MSIX.ps1" %*
echo.
pause
