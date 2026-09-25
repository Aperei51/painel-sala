@echo off
title Painel de Sala - Instalacao
echo Instalando o Painel de Sala (atalhos na area de trabalho e inicializacao automatica)...
echo.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Instalar.ps1" %*
echo.
pause
