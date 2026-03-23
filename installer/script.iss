#define Repository      "..\."
#define MyAppName       "Schedules Excel Export Revit Add-in"
#define MyAppVersion    "1.1.0"
#define MyAppPublisher  "ancodia"
#define MyAppURL        "https://github.com/ancodia/schedules-excel-export-revit-addin"
#define MyAppExeName    "SchedulesExcelExport.Revit.Addin.Installer.exe"


#define RevitAppName      "SchedulesExcelExport"
#define RevitAddinFolder  "{sd}\ProgramData\Autodesk\Revit\Addins"
#define RevitFolder20     RevitAddinFolder+"\2020\"+RevitAppName
#define RevitAddin20      RevitAddinFolder+"\2020\"
#define RevitFolder21     RevitAddinFolder+"\2021\"+RevitAppName
#define RevitAddin21      RevitAddinFolder+"\2021\"
#define RevitFolder22     RevitAddinFolder+"\2022\"+RevitAppName
#define RevitAddin22      RevitAddinFolder+"\2022\"
#define RevitFolder23     RevitAddinFolder+"\2023\"+RevitAppName
#define RevitAddin23      RevitAddinFolder+"\2023\"
#define RevitFolder24     RevitAddinFolder+"\2024\"+RevitAppName
#define RevitAddin24      RevitAddinFolder+"\2024\"
#define RevitFolder25     RevitAddinFolder+"\2025\"+RevitAppName
#define RevitAddin25      RevitAddinFolder+"\2025\"

[Setup]
AppId=DAF7B9CA-9E11-41BA-9263-93FBCF68D74C
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={commonpf64}\{#MyAppName}
DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
DisableWelcomePage=no
OutputDir={#Repository}\output
OutputBaseFilename=SchedulesExcelExport-{#MyAppVersion}-setup
Compression=lzma
SolidCompression=yes
ChangesAssociations=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Components]
Name: revit20; Description: Addin for Autodesk Revit 2020;  Types: full
Name: revit21; Description: Addin for Autodesk Revit 2021;  Types: full
Name: revit22; Description: Addin for Autodesk Revit 2022;  Types: full
Name: revit23; Description: Addin for Autodesk Revit 2023;  Types: full
Name: revit24; Description: Addin for Autodesk Revit 2024;  Types: full
Name: revit25; Description: Addin for Autodesk Revit 2025;  Types: full

[Files]
; ========== Revit 2020 ==========
Source: "{#Repository}\src\output\Release-2020\*.dll"; Excludes: "RevitAPI.dll,RevitAPIUI.dll,AdWindows.dll,UIFramework.dll"; DestDir: "{#RevitFolder20}"; Flags: ignoreversion; Components: revit20
Source: "{#Repository}\src\*.addin"; DestDir: "{#RevitAddin20}"; Flags: ignoreversion; Components: revit20

; ========== Revit 2021 ==========
Source: "{#Repository}\src\output\Release-2021\*.dll"; Excludes: "RevitAPI.dll,RevitAPIUI.dll,AdWindows.dll,UIFramework.dll"; DestDir: "{#RevitFolder21}"; Flags: ignoreversion; Components: revit21
Source: "{#Repository}\src\*.addin"; DestDir: "{#RevitAddin21}"; Flags: ignoreversion; Components: revit21

; ========== Revit 2022 ==========
Source: "{#Repository}\src\output\Release-2022\*.dll"; Excludes: "RevitAPI.dll,RevitAPIUI.dll,AdWindows.dll,UIFramework.dll"; DestDir: "{#RevitFolder22}"; Flags: ignoreversion; Components: revit22
Source: "{#Repository}\src\*.addin"; DestDir: "{#RevitAddin22}"; Flags: ignoreversion; Components: revit22

; ========== Revit 2023 ==========
Source: "{#Repository}\src\output\Release-2023\*.dll"; Excludes: "RevitAPI.dll,RevitAPIUI.dll,AdWindows.dll,UIFramework.dll"; DestDir: "{#RevitFolder23}"; Flags: ignoreversion; Components: revit23
Source: "{#Repository}\src\*.addin"; DestDir: "{#RevitAddin23}"; Flags: ignoreversion; Components: revit23

; ========== Revit 2024 ==========
Source: "{#Repository}\src\output\Release-2024\*.dll"; Excludes: "RevitAPI.dll,RevitAPIUI.dll,AdWindows.dll,UIFramework.dll"; DestDir: "{#RevitFolder24}"; Flags: ignoreversion; Components: revit24
Source: "{#Repository}\src\*.addin"; DestDir: "{#RevitAddin24}"; Flags: ignoreversion; Components: revit24

; ========== Revit 2025 ==========
Source: "{#Repository}\src\output\Release-2025\*.dll"; Excludes: "RevitAPI.dll,RevitAPIUI.dll,AdWindows.dll,UIFramework.dll"; DestDir: "{#RevitFolder25}"; Flags: ignoreversion; Components: revit25
Source: "{#Repository}\src\*.addin"; DestDir: "{#RevitAddin25}"; Flags: ignoreversion; Components: revit25