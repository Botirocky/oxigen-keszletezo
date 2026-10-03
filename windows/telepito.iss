#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif

#define AppName "Oxigén készletező"
#define AppExe "Oxigen_keszletezo.exe"
#define DataDir "{userdocs}\Oxigén készletező"

[Setup]
AppId={{6E4BBAA0-14B3-4865-935C-9685C5B5BBD1}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher=Botirocky
AppPublisherURL=https://github.com/Botirocky/oxigen-keszletezo
AppSupportURL=https://github.com/Botirocky/oxigen-keszletezo/blob/main/docs/TELEPITES_WINDOWS.md
AppUpdatesURL=https://github.com/Botirocky/oxigen-keszletezo/releases
AppCopyright=© Botirocky
VersionInfoCompany=Botirocky
VersionInfoProductName={#AppName}
VersionInfoDescription={#AppName} telepítő
VersionInfoVersion={#AppVersion}
DefaultDirName={localappdata}\Programs\Oxigen keszletezo
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=Oxigen_keszletezo_telepito_v{#AppVersion}
SetupIconFile=..\assets\app_icon.ico
UninstallDisplayIcon={app}\{#AppExe}
UninstallDisplayName={#AppName}
WizardStyle=modern
Compression=lzma2
SolidCompression=yes
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "hungarian"; MessagesFile: "compiler:Languages\Hungarian.isl"

[Tasks]
Name: "asztal"; Description: "Ikon az asztalra"; GroupDescription: "További ikonok:"

[Dirs]
Name: "{#DataDir}"; Flags: uninsneveruninstall

[Files]
Source: "..\dist\{#AppExe}"; DestDir: "{app}"; Flags: ignoreversion
Source: "telepitett.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autoprograms}\{#AppName} – adatok (Excel)"; Filename: "{#DataDir}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: asztal

[Run]
Filename: "{app}\{#AppExe}"; Description: "A program indítása"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: files; Name: "{app}\{#AppExe}.new"
Type: files; Name: "{app}\frissites_csere.bat"

[Messages]
hungarian.FinishedLabel=A telepítés kész.%n%nA bolti adatok (keszlet.xlsx, beállítások, mentések) helye:%nDokumentumok\Oxigén készletező%n%nAz eltávolítás ezeket NEM törli.
