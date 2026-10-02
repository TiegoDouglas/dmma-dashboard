@echo off
setlocal EnableExtensions
title Sincronizacao de vagas com Azure DevOps

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0executar_sincronizacao.ps1"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
echo Pressione qualquer tecla para fechar.
pause >nul

endlocal & exit /b %EXIT_CODE%
