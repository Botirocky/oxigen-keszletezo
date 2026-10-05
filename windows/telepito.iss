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
WizardSizePercent=110
WizardImageFile=kepek\nagy_100.bmp,kepek\nagy_125.bmp,kepek\nagy_150.bmp
WizardSmallImageFile=kepek\kicsi_100.bmp,kepek\kicsi_125.bmp,kepek\kicsi_150.bmp,kepek\kicsi_175.bmp,kepek\kicsi_200.bmp,kepek\kicsi_225.bmp,kepek\kicsi_250.bmp
DisableWelcomePage=no
DisableReadyPage=yes
Compression=lzma2
SolidCompression=yes
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "hungarian"; MessagesFile: "compiler:Languages\Hungarian.isl"

[Tasks]
Name: "asztal"; Description: "Ikon az asztalra"; GroupDescription: "Parancsikonok:"

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
hungarian.WelcomeLabel1=Üdv az Oxigén készletezőben!
hungarian.WelcomeLabel2=Ez a varázsló telepíti a(z) [name/ver] programot erre a gépre.%n%nA telepítéshez nem kell rendszergazdai jog, és pár másodperc az egész.
hungarian.WizardSelectTasks=Parancsikonok
hungarian.SelectTasksDesc=Hol legyen ikonja a programnak?
hungarian.SelectTasksLabel2=Kérsz ikont az asztalra is? Utána a Telepítés gombbal indul.
hungarian.FinishedHeadingLabel=Kész, a program telepítve!
hungarian.FinishedLabel=A telepítés kész.%n%nA bolti adatok (keszlet.xlsx, beállítások, mentések) helye:%nDokumentumok\Oxigén készletező%n%nAz eltávolítás ezeket NEM törli.

[Code]
const
  Zold = $002B3516;
  Krem = $00E8F1F4;
  Hatter = $00F1F4F1;
  Felirat = $00CDE3BF;
  Szoveg = $001E2416;

procedure InitializeWizard;
begin
  WizardForm.Color := Hatter;
  WizardForm.MainPanel.Color := Zold;
  WizardForm.PageNameLabel.Font.Color := Krem;
  WizardForm.PageDescriptionLabel.Font.Color := Felirat;
  WizardForm.WelcomePage.Color := Krem;
  WizardForm.WelcomeLabel1.Font.Color := Zold;
  WizardForm.WelcomeLabel2.Font.Color := Szoveg;
  WizardForm.FinishedPage.Color := Krem;
  WizardForm.FinishedHeadingLabel.Font.Color := Zold;
  WizardForm.FinishedLabel.Font.Color := Szoveg;
  WizardForm.RunList.Color := Krem;
  WizardForm.RunList.Font.Color := Szoveg;
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if CurPageID = wpSelectTasks then
    WizardForm.NextButton.Caption := SetupMessage(msgButtonInstall);
end;
