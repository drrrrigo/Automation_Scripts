@echo off
setlocal EnableExtensions

net session >nul 2>&1
if not "%errorlevel%"=="0" (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

set "TASK_NAME=Hapvida - Limpeza PixViewerLogs Arya"
set "BASE_DIR=%ProgramData%\Hapvida\AryaCleanup"
set "CLEANUP_PS1=%BASE_DIR%\Arya_Cleanup_PixViewerLogs.ps1"
set "RUNNER_CMD=%BASE_DIR%\Executar_Limpeza_Arya_Logs.cmd"
set "INSTALL_LOG=%BASE_DIR%\instalacao.log"
set "RUN_LOG=%BASE_DIR%\execucao.log"
set "START_TIME="

if not exist "%BASE_DIR%" mkdir "%BASE_DIR%"

> "%INSTALL_LOG%" echo [%date% %time%] Iniciando instalacao da tarefa.

for /f %%I in ('powershell -NoProfile -ExecutionPolicy Bypass -Command "(Get-Date).AddMinutes(1).ToString('HH:mm')"') do set "START_TIME=%%I"
if not defined START_TIME set "START_TIME=00:00"
>> "%INSTALL_LOG%" echo [%date% %time%] Horario inicial configurado para %START_TIME%.
>> "%INSTALL_LOG%" echo [%date% %time%] Executor da tarefa: %RUNNER_CMD%
>> "%INSTALL_LOG%" echo [%date% %time%] Contexto da tarefa: SYSTEM - valida para todos os usuarios da maquina.

(
echo $ErrorActionPreference = 'SilentlyContinue'
echo $usersRoot = 'C:\Users'
echo $runLog = 'C:\ProgramData\Hapvida\AryaCleanup\execucao.log'
echo $excludedProfiles = @^(
echo     'All Users',
echo     'Default',
echo     'Default User',
echo     'Public',
echo     'defaultuser0',
echo     'WDAGUtilityAccount'
echo ^)
echo function Write-RunLog {
echo     param^([string]$Message^)
echo     $timestamp = Get-Date -Format 'dd/MM/yyyy HH:mm:ss'
echo     Add-Content -LiteralPath $runLog -Value "[$timestamp] $Message"
echo }
echo Write-RunLog 'Inicio da limpeza.'
echo if ^(-not ^(Test-Path -LiteralPath $usersRoot^)^) {
echo     Write-RunLog 'C:\Users nao encontrado.'
echo     exit 0
echo }
echo $profiles = Get-ChildItem -LiteralPath $usersRoot -Directory -Force -ErrorAction SilentlyContinue ^| Where-Object {
echo     $excludedProfiles -notcontains $_.Name
echo }
echo foreach ^($profile in $profiles^) {
echo     try {
echo         $folders = New-Object System.Collections.Generic.List[System.String]
echo         $directFolder = Join-Path -Path $profile.FullName -ChildPath 'pixviewerlogs'
echo         if ^(Test-Path -LiteralPath $directFolder^) {
echo             [void]$folders.Add^($directFolder^)
echo         }
echo         $foundFolders = Get-ChildItem -LiteralPath $profile.FullName -Directory -Recurse -Force -ErrorAction SilentlyContinue -Filter 'pixviewerlogs'
echo         foreach ^($found in $foundFolders^) {
echo             if ^(-not $folders.Contains^($found.FullName^)^) {
echo                 [void]$folders.Add^($found.FullName^)
echo             }
echo         }
echo         foreach ^($folderPath in $folders^) {
echo             Write-RunLog ^("Limpando " + $folderPath^)
echo             $items = Get-ChildItem -LiteralPath $folderPath -Force -ErrorAction SilentlyContinue
echo             foreach ^($item in $items^) {
echo                 try {
echo                     Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
echo                 }
echo                 catch {
echo                     Write-RunLog ^("Ignorado em uso/bloqueado: " + $item.FullName^)
echo                 }
echo             }
echo         }
echo     }
echo     catch {
echo         Write-RunLog ^("Falha ao processar perfil: " + $profile.FullName^)
echo     }
echo }
echo Write-RunLog 'Fim da limpeza.'
echo exit 0
)> "%CLEANUP_PS1%"

if errorlevel 1 (
    >> "%INSTALL_LOG%" echo [%date% %time%] Falha ao criar o script PowerShell.
    echo Falha ao criar o script de limpeza.
    exit /b 1
)

(
echo @echo off
echo powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%CLEANUP_PS1%"
echo exit /b 0
)> "%RUNNER_CMD%"

if errorlevel 1 (
    >> "%INSTALL_LOG%" echo [%date% %time%] Falha ao criar o executor CMD.
    echo Falha ao criar o arquivo executor.
    exit /b 1
)

if not exist "%CLEANUP_PS1%" (
    >> "%INSTALL_LOG%" echo [%date% %time%] Script PowerShell nao localizado apos criacao.
    echo Script PowerShell nao localizado apos criacao.
    exit /b 1
)

if not exist "%RUNNER_CMD%" (
    >> "%INSTALL_LOG%" echo [%date% %time%] Executor CMD nao localizado apos criacao.
    echo Executor CMD nao localizado apos criacao.
    exit /b 1
)

>> "%INSTALL_LOG%" echo [%date% %time%] Script criado em "%CLEANUP_PS1%".
>> "%INSTALL_LOG%" echo [%date% %time%] Executor criado em "%RUNNER_CMD%".

schtasks /Delete /TN "%TASK_NAME%" /F >nul 2>&1

schtasks /Create ^
 /TN "%TASK_NAME%" ^
 /SC MINUTE ^
 /MO 15 ^
 /ST %START_TIME% ^
 /RU SYSTEM ^
 /RL HIGHEST ^
 /TR "%RUNNER_CMD%" ^
 /F >nul 2>&1

if errorlevel 1 (
    >> "%INSTALL_LOG%" echo [%date% %time%] Falha ao registrar a tarefa agendada.
    echo Falha ao registrar a tarefa agendada.
    exit /b 1
)

schtasks /Query /TN "%TASK_NAME%" >nul 2>&1
if errorlevel 1 (
    >> "%INSTALL_LOG%" echo [%date% %time%] A tarefa nao foi localizada apos a criacao.
    echo A tarefa nao foi localizada apos a criacao.
    exit /b 1
)

schtasks /Run /TN "%TASK_NAME%" >nul 2>&1
>> "%INSTALL_LOG%" echo [%date% %time%] Tarefa criada com sucesso: %TASK_NAME%
>> "%INSTALL_LOG%" echo [%date% %time%] Log de execucao: "%RUN_LOG%"
>> "%INSTALL_LOG%" echo [%date% %time%] A limpeza sera executada para todos os perfis encontrados em C:\Users.

echo Tarefa criada com sucesso: %TASK_NAME%
exit /b 0
