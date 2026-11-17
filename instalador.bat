@echo off
TITLE Instalador do App de Timesheet
CLS

echo ========================================================
echo   Configurando o App de Timesheet...
echo ========================================================

:: 1. Cria a pasta de destino fixa no computador
if not exist "C:\Produtividade" mkdir "C:\Produtividade"

:: 2. Copia o executável correto para a pasta C:\Produtividade
echo Copiando arquivo executavel...
copy /Y "%~dp0app_timesheet.exe" "C:\Produtividade\app_timesheet.exe"

:: 3. Cria automaticamente o atalho na pasta de Inicializacao do Windows (Startup) para rodar com o sistema
echo Configurando inicializacao automatica com o Windows...
powershell -Command "$Wscript = New-Object -ComObject WScript.Shell; $Shortcut = $Wscript.CreateShortcut('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\AppTimesheet.lnk'); $Shortcut.TargetPath = 'C:\Produtividade\app_timesheet.exe'; $Shortcut.Save()"

:: 4. Inicializa o aplicativo imediatamente apos a instalacao
echo Iniciando o aplicativo de timesheet...
start "" "C:\Produtividade\app_timesheet.exe"

echo ========================================================
echo   Instalacao concluida com sucesso e app rodando!
echo ========================================================
pause