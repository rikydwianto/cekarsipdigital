; Inno Setup Script untuk Arsip Digital OwnCloud Application
; Generated for cx_Freeze build output
; Download Inno Setup: https://jrsoftware.org/isdl.php

#define MyAppName "Arsip Digital OwnCloud"
#define MyAppVersion "1.1.6"
#define MyAppPublisher "Your Organization"
#define MyAppURL "https://github.com/rikydwianto/cekarsipdigital"
#define MyAppExeName "main.exe"

[Setup]
AppId={{A8F9B3C4-D5E6-4F7A-8B9C-0D1E2F3A4B5C}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
LicenseFile=frozen_application_license.txt
OutputDir=installer_output
OutputBaseFilename=ArsipDigital_Setup_v{#MyAppVersion}
SetupIconFile=
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "main.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "python310.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "frozen_application_license.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "app_config.json"; DestDir: "{app}"; Flags: ignoreversion

Source: "lib\*"; DestDir: "{app}\lib"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "share\*"; DestDir: "{app}\share"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "src_web\*"; DestDir: "{app}\src_web"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "poppler-25.07.0\*"; DestDir: "{app}\poppler-25.07.0"; Flags: ignoreversion recursesubdirs createallsubdirs

; Sertakan installer Visual C++ Redistributable (x64)
Source: "VC_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Jalankan installer Visual C++ hanya jika belum terinstal
Filename: "{tmp}\VC_redist.x64.exe"; Parameters: "/install /quiet /norestart"; \
    Check: not VCInstalled64; Flags: waituntilterminated

; Jalankan aplikasi setelah selesai install
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
// Fungsi cek registry untuk VC++ Redistributable 64-bit
function VCInstalled64: Boolean;
var
  Installed: Cardinal;
begin
  Result := RegQueryDWordValue(
    HKLM, 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64', 'Installed', Installed)
    and (Installed = 1);
end;

function InitializeSetup(): Boolean;
begin
  if VCInstalled64 then
    MsgBox('Microsoft Visual C++ Redistributable x64 sudah terpasang di sistem Anda.', mbInformation, MB_OK)
  else
    MsgBox('Microsoft Visual C++ Redistributable x64 belum terdeteksi. Installer akan menginstalnya secara otomatis.', mbInformation, MB_OK);
  Result := True;
end;

[UninstallDelete]
Type: filesandordirs; Name: "{app}\app_config.json"
Type: filesandordirs; Name: "{app}\*.log"
Type: filesandordirs; Name: "{app}\database.xlsx"
