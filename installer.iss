[Setup]
AppName=Inventarisierung
#ifdef AppVersion
AppVersion={#AppVersion}
#else
AppVersion=1.0.0.0
#endif
AppPublisher=Abkoo AG
DefaultDirName={autopf}\Inventarisierung
DefaultGroupName=Inventarisierung
OutputDir=C:\Users\CC-Student\Desktop\Installer
OutputBaseFilename=Inventarisierung_Setup_{#AppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Tasks]
Name: "desktopicon"; Description: "Desktop-Verknüpfung erstellen"; GroupDescription: "Zusätzliche Aufgaben:"

[Files]
Source: "C:\Users\CC-Student\source\repos\WindowsFormsApp1\WindowsFormsApp1\bin\Release\Inventarisierung.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\CC-Student\source\repos\WindowsFormsApp1\WindowsFormsApp1\bin\Release\*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\CC-Student\source\repos\WindowsFormsApp1\WindowsFormsApp1\bin\Release\*.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\CC-Student\source\repos\WindowsFormsApp1\WindowsFormsApp1\bin\Release\x64\*"; DestDir: "{app}\x64"; Flags: ignoreversion
Source: "C:\Users\CC-Student\source\repos\WindowsFormsApp1\WindowsFormsApp1\bin\Release\x86\*"; DestDir: "{app}\x86"; Flags: ignoreversion

[Icons]
Name: "{group}\Inventarisierung"; Filename: "{app}\Inventarisierung.exe"
Name: "{group}\Deinstallieren"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Inventarisierung"; Filename: "{app}\Inventarisierung.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Inventarisierung.exe"; Description: "Inventarisierung starten"; Flags: nowait postinstall skipifsilent
