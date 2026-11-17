; ============================================================
;  Instalador do Monitor de Produtividade - TimeSheet
;  Requer: Inno Setup 6.x  (https://jrsoftware.org/isdl.php)
;
;  Antes de compilar, gere o executavel:
;     python -m PyInstaller app_timesheet_integrado.spec --noconfirm --clean
;  Isso cria dist\MonitorTimeSheet.exe (referenciado abaixo).
;
;  Para compilar o instalador: abra este arquivo no Inno Setup 6 e clique
;  em "Compile" (Run > Compile). O Setup.exe sai na pasta "Output".
; ============================================================

#define MeuApp "Monitor de Produtividade"
#define MeuAppVersao "1.0.2"
#define MeuAppPublisher "JRW Solucoes em TI"
#define MeuAppExe "MonitorTimeSheet.exe"
#define MeuAppURL "https://ecadvogados.azurewebsites.net"

[Setup]
AppId={{A7E3F1C2-8B4D-4E9A-9C21-TIMESHEET0001}
AppName={#MeuApp}
AppVersion={#MeuAppVersao}
AppPublisher={#MeuAppPublisher}
AppPublisherURL={#MeuAppURL}
DefaultDirName={autopf}\MonitorTimeSheet
DefaultGroupName=Monitor de Produtividade
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=MonitorTimeSheet-Setup-{#MeuAppVersao}
SetupIconFile=timesheet.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
; Nao exige admin: instala no perfil do usuario e configura inicio automatico por usuario
PrivilegesRequired=lowest
; Mesmo mutex da instância única do app: o instalador detecta o Monitor aberto e pede para fechá-lo
AppMutex=JRW_MonitorTimeSheet_SingleInstance
UninstallDisplayIcon={app}\{#MeuAppExe}
UninstallDisplayName={#MeuApp}

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "iniciarcomwindows"; Description: "Iniciar o monitor automaticamente com o Windows (recomendado)"; GroupDescription: "Inicializacao:"
Name: "desktopicon"; Description: "Criar atalho na area de trabalho"; GroupDescription: "Atalhos:"; Flags: unchecked

[Files]
Source: "dist\{#MeuAppExe}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Atalho no menu iniciar
Name: "{autoprograms}\{#MeuApp}"; Filename: "{app}\{#MeuAppExe}"
; Atalho na area de trabalho (opcional, via Task)
Name: "{autodesktop}\{#MeuApp}"; Filename: "{app}\{#MeuAppExe}"; Tasks: desktopicon

[Registry]
; Inicio automatico com o Windows (por usuario) - so se a task estiver marcada
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; \
    ValueType: string; ValueName: "MonitorTimeSheet"; \
    ValueData: """{app}\{#MeuAppExe}"""; \
    Flags: uninsdeletevalue; Tasks: iniciarcomwindows

[Run]
; Abre o app logo apos instalar
Filename: "{app}\{#MeuAppExe}"; Description: "Iniciar o Monitor agora"; Flags: nowait postinstall skipifsilent
; Atualizacao automatica (instalacao silenciosa disparada pelo proprio app): reabre o Monitor ao terminar
Filename: "{app}\{#MeuAppExe}"; Flags: nowait; Check: WizardSilent

[UninstallRun]
; Encerra o app ao desinstalar (evita arquivo em uso)
Filename: "{cmd}"; Parameters: "/C taskkill /F /IM {#MeuAppExe}"; Flags: runhidden; RunOnceId: "KillMonitor"
