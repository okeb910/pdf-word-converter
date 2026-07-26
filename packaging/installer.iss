#define MyAppName "PDF Word PPT Converter"
#define MyAppVersion "0.4.0"
#define MyAppExeName "PDFWordConverter.exe"

[Setup]
AppId={{EE1F5FCB-3E00-4A0C-BD4A-AFAFD380B055}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=okeb910
AppPublisherURL=https://github.com/okeb910/pdf-word-converter
DefaultDirName={localappdata}\Programs\PDFWordConverter
DefaultGroupName=PDF Word PPT Converter
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\dist
OutputBaseFilename=PDFWordConverter-v0.4.0-Setup-x64
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
CloseApplications=yes
SetupLogging=yes
VersionInfoVersion=0.4.0.0
VersionInfoCompany=okeb910
VersionInfoDescription=PDF Word PPT Converter Setup
VersionInfoProductName=PDF Word PPT Converter
VersionInfoProductVersion=0.4.0

[Languages]
Name: "chinesesimplified"; MessagesFile: "ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "..\dist\PDFWordConverter-v0.4.0\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\PDF Word PPT Converter"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\PDF Word PPT Converter"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch PDF Word PPT Converter"; Flags: nowait postinstall skipifsilent
