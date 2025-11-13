; excelSheetSorterInstaller.iss
; Inno Setup script for Excel Sheet Sorter
[Setup]
AppName=Excel Sheet Sorter
AppVersion=v_1.0.0
AppPublisher=NKau$hal
AppPublisherURL=https://www.linkedin.com/in/narender-kaushal-b72b075/
DefaultDirName={pf}\Excel Sheet Sorter
DefaultGroupName=Excel Sheet Sorter
DisableDirPage=no
DisableProgramGroupPage=no
UninstallDisplayIcon={app}\ExcelSheetSorter.exe
OutputDir=.
OutputBaseFilename=ExcelSheetSorterInstaller
SetupIconFile=C:\Users\NKaushal\excel_Sorter\excelWorkBook_Sorter\package\appIcon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Main application EXE from PyInstaller
Source: "dist\ExcelSheetSorter.exe"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion
Source: "appIcon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "titleAppIcon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "Use_cases_ExcelWorkbook_Sorter.pdf"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu shortcut
Name: "{commondesktop}\Excel Sheet Sorter"; Filename: "{app}\ExcelSheetSorter.exe"; IconFilename: "{app}\appIcon.ico"
Name: "{group}\Excel Sheet Sorter"; Filename: "{app}\ExcelSheetSorter.exe"; IconFilename: "{app}\appIcon.ico"
Name: "{group}\Excel Sheet Sorter - User Guide"; Filename: "{app}\Use_cases_ExcelWorkbook_Sorter.pdf"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Run]
; Offer to launch app after install
Filename: "{app}\ExcelSheetSorter.exe"; Description: "Launch Excel Sheet Sorter"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Optional: clean up any extra files the app might create in its folder
Type: filesandordirs; Name: "{app}\logs"