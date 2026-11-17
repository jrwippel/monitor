@echo off
TITLE Reset e Limpeza do App de Timesheet
CLS

echo ========================================================
echo   Limpando dados e encerrando o App de Timesheet...
echo ========================================================

:: 1. Encerra qualquer processo do timesheet que esteja rodando em segundo plano
taskkill /F /IM app_timesheet.exe >nul 2>&1

:: 2. Remove o executável da pasta de destino fixa
if exist "C:\Produtividade\app_timesheet.exe" (
    del /Q "C:\Produtividade\app_timesheet.exe"
    echo [OK] Executavel removido de C:\Produtividade.
)

:: 3. Remove o atalho da pasta de Inicializacao do Windows (Startup)
if exist "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\AppTimesheet.lnk" (
    del /Q "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\AppTimesheet.lnk"
    echo [OK] Atalho de inicializacao automatica removido.
)

:: 4. Remove o arquivo de configuracao (e-mail salvo nos Documentos)
if exist "%USERPROFILE%\Documents\config_timesheet_escritorio.json" (
    del /Q "%USERPROFILE%\Documents\config_timesheet_escritorio.json"
    echo [OK] Arquivo de configuracao e e-mail salvo removidos.
)

echo ========================================================
echo   Reset concluido com sucesso! O ambiente esta limpo.
echo ========================================================
pause