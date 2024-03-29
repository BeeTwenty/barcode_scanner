; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "Barcode Scanner"
#define MyAppVersion "1.1.0"
#define MyAppPublisher "Sindre Berge"
#define MyAppURL "https://github.com/BeeTwenty/barcode_scanner"
#define MyAppExeName "barcode_scanner.exe"
#define MyAppAssocName MyAppName + " File"
#define MyAppAssocExt ".myp"
#define MyAppAssocKey StringChange(MyAppAssocName, " ", "") + MyAppAssocExt

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{DED19133-B841-4F0C-8F54-AE8E377D874D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
ChangesAssociations=yes
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
LicenseFile=C:\Users\SBerge1\Documents\barcode_scanner\LICENSE
; Remove the following line to run in administrative install mode (install for all users.)
PrivilegesRequired=lowest
OutputDir=C:\Users\SBerge1\Documents\barcode_scanner\setup
OutputBaseFilename=BarcodeSetup
SetupIconFile=C:\Users\SBerge1\Documents\barcode_scanner\barcode.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; other setup options...
; Enable logging
SetupLogging=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "C:\Users\SBerge1\Documents\barcode_scanner\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\SBerge1\Documents\barcode_scanner\barcode.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\SBerge1\Documents\barcode_scanner\changelog.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\SBerge1\Documents\barcode_scanner\preferences.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\SBerge1\Documents\barcode_scanner\README.txt"; DestDir: "{app}"; Flags: isreadme
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Registry]
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocExt}\OpenWithProgids"; ValueType: string; ValueName: "{#MyAppAssocKey}"; ValueData: ""; Flags: uninsdeletevalue
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}"; ValueType: string; ValueName: ""; ValueData: "{#MyAppAssocName}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#MyAppExeName},0"
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".myp"; ValueData: ""

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon


[Code]
function NextButtonClick(CurPage: Integer): Boolean;
var
  ResultCode: Integer;
begin
  if CurPage = wpSelectTasks then
  begin
    // Terminate the barcode program
    Exec('taskkill', '/F /IM Barcode_Scanner.exe', '', SW_HIDE, 
      ewNoWait, ResultCode);
  end;
  Result := True;
end;

[Run]

Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

