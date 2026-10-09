; Busca Databook - instalador por usuario (sem administrador)
; Compilar:  ISCC.exe /DAppVersion=3.0.0 installer\BuscaDatabook.iss
; Espera encontrar:
;   ..\dist\BuscaDatabook\      saida do PyInstaller (onedir)
;   ..\vendor\tesseract\        Tesseract OCR com tessdata\por.traineddata

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#define AppName "Busca Databook"
#define AppExe "BuscaDatabook.exe"
#define AppUrl "https://github.com/AlexandreTavares75/busca-databook"

[Setup]
; Nunca altere o AppId: e ele que liga as atualizacoes a instalacao existente.
AppId={{8F3C2A71-5B4E-4D2A-9C61-3E7A1B9D0F42}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher=Alexandre Tavares
AppPublisherURL={#AppUrl}
AppSupportURL={#AppUrl}/issues
AppUpdatesURL={#AppUrl}/releases
VersionInfoVersion={#AppVersion}

; Instalacao por usuario: %LOCALAPPDATA%\Programs\BuscaDatabook, sem UAC
PrivilegesRequired=lowest
DefaultDirName={autopf}\BuscaDatabook
DisableDirPage=auto
DisableProgramGroupPage=yes
UsePreviousAppDir=yes

ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

OutputDir=..\output
OutputBaseFilename=BuscaDatabook-Setup-{#AppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
SetupIconFile=..\assets\icone.ico
UninstallDisplayIcon={app}\{#AppExe}
UninstallDisplayName={#AppName}
SetupLogging=yes
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na Área de Trabalho"; GroupDescription: "Atalhos:"

[InstallDelete]
; Remove bibliotecas da versao anterior antes de copiar as novas (evita DLLs misturadas)
Type: filesandordirs; Name: "{app}\_internal"
Type: filesandordirs; Name: "{app}\tesseract"

[Files]
Source: "..\dist\BuscaDatabook\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\vendor\tesseract\*"; DestDir: "{app}\tesseract"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\LEIA-ME.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
; Instalacao normal: oferece abrir o programa no final
Filename: "{app}\{#AppExe}"; Description: "Abrir o {#AppName}"; Flags: nowait postinstall skipifsilent
; Atualizacao automatica (/SILENT /RELAUNCH): reabre o programa sozinho
Filename: "{app}\{#AppExe}"; Flags: nowait; Check: DeveReabrir

; Os dados do usuario (indices, usuarios.db, config, erro.log) ficam em
; %LOCALAPPDATA%\BuscaDatabook e NAO sao apagados na desinstalacao.

[Code]
const
  MUTEX_APP = 'BuscaDatabook_Instancia';

{ Espera ate ~10 s o programa fechar (ele fecha sozinho ao iniciar a atualizacao). }
function ProgramaFechou(): Boolean;
var
  I: Integer;
begin
  for I := 1 to 40 do
  begin
    if not CheckForMutexes(MUTEX_APP) then
    begin
      Result := True;
      Exit;
    end;
    Sleep(250);
  end;
  Result := not CheckForMutexes(MUTEX_APP);
end;

function ConfirmarFechado(): Boolean;
begin
  Result := True;
  while not ProgramaFechou() do
  begin
    if SuppressibleMsgBox('O Busca Databook está aberto.' + #13#10 +
         'Feche o programa e clique em Repetir.', mbError, MB_RETRYCANCEL, IDCANCEL) = IDCANCEL then
    begin
      Result := False;
      Exit;
    end;
  end;
end;

function InitializeSetup(): Boolean;
begin
  Result := ConfirmarFechado();
end;

function InitializeUninstall(): Boolean;
begin
  Result := ConfirmarFechado();
end;

function DeveReabrir(): Boolean;
begin
  Result := WizardSilent() and (Pos('/RELAUNCH', Uppercase(GetCmdTail())) > 0);
end;