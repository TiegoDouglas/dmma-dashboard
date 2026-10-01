@echo off
setlocal
call "%~dp0automacao_azure_devops\executar_sincronizacao.bat"
set "EXIT_CODE=%ERRORLEVEL%"
endlocal & exit /b %EXIT_CODE%
