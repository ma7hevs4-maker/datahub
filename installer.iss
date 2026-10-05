[Setup]
AppName=DataHub
AppVersion=2.0.0
AppPublisher=DataHub
DefaultDirName={localappdata}\DataHub
DefaultGroupName=DataHub
OutputDir=C:\Users\BR0163806927\Downloads\Agente I.A\datahub_v2\dist
OutputBaseFilename=DataHub-2.0.0-Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
CloseApplications=yes
CloseApplicationsFilter=DataHub.exe
RestartApplications=no
UninstallDisplayIcon={app}\DataHub.exe
ArchitecturesInstallIn64BitMode=x64
ArchitecturesAllowed=x64
WizardStyle=modern
DisableProgramGroupPage=auto
; O app atualiza sozinho pelo GitHub, entao instalamos numa pasta gravavel
; sem admin (localappdata) para o updater funcionar sem pedir UAC.

[Messages]
; textos em portugues (fallback english se faltar)
WelcomeLabel1= Bem-vindo ao Assistente de Instalação do DataHub

[InstallDelete]
; limpa flag obsoleta de falha de update para nao avisar a toa apos reinstalar
Type: files; Name: "{localappdata}\DataHub\update_failed.flag"

[Files]
Source: C:\Users\BR0163806927\Downloads\Agente I.A\datahub_v2\dist\DataHub\*; DestDir: {app}; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: {group}\DataHub; Filename: {app}\DataHub.exe
Name: {autodesktop}\DataHub; Filename: {app}\DataHub.exe; Tasks: desktopicon

[Tasks]
Name: desktopicon; Description: "Criar atalho na Área de Trabalho"; GroupDescription: "Atalhos:"; Flags: unchecked
