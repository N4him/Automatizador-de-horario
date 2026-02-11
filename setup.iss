; Script de Inno Setup para Sistema de Asignación de Monitores
; Actualizado para la nueva estructura del proyecto

#define MyAppName "Sistema de Asignación de Monitores"
#define MyAppVersion "1.0"
#define MyAppPublisher "Tu Nombre o Empresa"
#define MyAppURL "https://tusitioweb.com"
#define MyAppExeName "Asignacion_Monitores.exe"

[Setup]
; IMPORTANTE: Genera un GUID único en https://www.guidgenerator.com/
AppId={{A7B8C9D0-E1F2-4A5B-8C9D-0E1F2A3B4C5D}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
; Licencia (opcional - descomenta si tienes un archivo de licencia)
; LicenseFile=LICENSE.txt
; Carpeta donde se guardará el instalador generado
OutputDir=installer_output
OutputBaseFilename=Setup_Sistema_Asignacion_Monitores_v{#MyAppVersion}
; Icono del instalador (ajustado a la nueva ubicación)
SetupIconFile=icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; Privilegios de instalación
PrivilegesRequired=lowest
; Arquitectura
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Archivo ejecutable principal desde la carpeta dist
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Todos los archivos de la carpeta dist (incluye dependencias de PyInstaller)
Source: "dist\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Iconos (ajustados a la nueva estructura)
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "icon.png"; DestDir: "{app}"; Flags: ignoreversion
; Archivos adicionales opcionales
; Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion isreadme
; Source: ".gitignore"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Acceso directo en el menú inicio
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\icon.ico"
; Acceso directo en el escritorio (opcional)
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon
; Desinstalador en el menú inicio
Name: "{autoprograms}\Desinstalar {#MyAppName}"; Filename: "{uninstallexe}"

[Run]
; Ejecutar la aplicación después de instalar (opcional)
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Limpiar archivos generados por la aplicación
Type: filesandordirs; Name: "{app}"
; Limpiar caché de Python si existe
Type: filesandordirs; Name: "{app}\__pycache__"
Type: filesandordirs; Name: "{app}\_internal"

[Code]
// Función para verificar si la aplicación está en ejecución
function InitializeSetup(): Boolean;
begin
  Result := True;
end;

function InitializeUninstall(): Boolean;
var
  ErrorCode: Integer;
begin
  Result := True;
  // Intentar cerrar la aplicación antes de desinstalar
  if CheckForMutexes('{#MyAppName}') then
  begin
    if MsgBox('La aplicación está en ejecución. ¿Desea cerrarla y continuar con la desinstalación?', 
              mbConfirmation, MB_YESNO) = IDYES then
    begin
      Result := True;
    end
    else
      Result := False;
  end;
end;