@echo off
setlocal EnableExtensions
title Sincronizacao de vaga com Azure DevOps

set "REPO=TiegoDouglas/dmma-dashboard"
set "WORKFLOW=sync-azure-work-item.yml"
set "BRANCH=main"
set "EXIT_CODE=0"

echo.
echo Sincronizacao de vaga com Azure DevOps
echo Repositorio: %REPO%
echo.

where gh >nul 2>nul
if errorlevel 1 (
  echo ERRO: GitHub CLI ^(gh^) nao foi encontrado no PATH.
  echo Instale-o em https://cli.github.com/ e tente novamente.
  set "EXIT_CODE=1"
  goto finish
)

echo Verificando autenticacao no GitHub...
call gh auth status
if errorlevel 1 (
  echo.
  echo ERRO: GitHub CLI nao esta autenticado.
  echo Execute "gh auth login" e tente novamente.
  set "EXIT_CODE=1"
  goto finish
)

for /f "usebackq delims=" %%I in (`powershell -NoProfile -Command "[DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')"`) do set "DISPATCH_TIME=%%I"
if not defined DISPATCH_TIME (
  echo ERRO: Nao foi possivel obter o horario do disparo.
  set "EXIT_CODE=1"
  goto finish
)

echo.
echo Disparando o workflow na branch %BRANCH%...
call gh workflow run "%WORKFLOW%" --repo "%REPO%" --ref "%BRANCH%"
if errorlevel 1 (
  echo ERRO: Nao foi possivel disparar o workflow.
  set "EXIT_CODE=1"
  goto finish
)

echo Localizando a execucao criada...
set /a ATTEMPT=0

:find_run
set /a ATTEMPT+=1
set "RUN_ID="
for /f "usebackq delims=" %%I in (`call gh run list --repo "%REPO%" --workflow "%WORKFLOW%" --branch "%BRANCH%" --event workflow_dispatch --created "^>=%DISPATCH_TIME%" --limit 20 --json databaseId,createdAt --jq "(sort_by(.createdAt) ^| last ^| .databaseId) // empty" 2^>nul`) do set "RUN_ID=%%I"

if defined RUN_ID goto watch_run
if %ATTEMPT% GEQ 30 (
  echo ERRO: A execucao nao foi localizada apos 60 segundos.
  set "EXIT_CODE=1"
  goto finish
)

timeout /t 2 /nobreak >nul
goto find_run

:watch_run
set "RUN_URL="
for /f "usebackq delims=" %%I in (`call gh run view "%RUN_ID%" --repo "%REPO%" --json url --jq ".url" 2^>nul`) do set "RUN_URL=%%I"

echo Execucao encontrada: %RUN_ID%
if defined RUN_URL echo URL: %RUN_URL%
echo Aguardando a conclusao...
echo.

call gh run watch "%RUN_ID%" --repo "%REPO%" --exit-status
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if "%EXIT_CODE%"=="0" (
  echo SUCESSO: sincronizacao concluida.
) else (
  echo ERRO: a execucao terminou com falha ^(codigo %EXIT_CODE%^).
  if defined RUN_URL echo Consulte os detalhes em: %RUN_URL%
)

:finish
echo.
echo Pressione qualquer tecla para fechar.
pause >nul
endlocal & exit /b %EXIT_CODE%
