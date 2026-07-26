#define MyAppName "PDF Word Converter"
#define MyAppVersion "0.3.0"
#define MyAppExeName "PDFWordConverter.exe"

[Setup]
AppId={{EE1F5FCB-3E00-4A0C-BD4A-AFAFD380B055}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=okeb910
AppPublisherURL=https://github.com/okeb910/pdf-word-converter
DefaultDirName={localappdata}\Programs\PDFWordConverter
DefaultGroupName=PDF Word Converter
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\dist
OutputBaseFilename=PDFWordConverter-v0.3.0-Setup-x64
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
CloseApplications=yes
SetupLogging=yes
VersionInfoVersion=0.3.0.0
VersionInfoCompany=okeb910
VersionInfoDescription=PDF Word Converter Setup
VersionInfoProductName=PDF Word Converter
VersionInfoProductVersion=0.3.0

[Languages]
Name: "chinesesimplified"; MessagesFile: "ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加快捷方式："; Flags: unchecked

[Files]
Source: "..\dist\PDFWordConverter-v0.3.0\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\PDF Word Converter"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\PDF Word Converter"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "启动 PDF Word Converter"; Flags: nowait postinstall skipifsilent
