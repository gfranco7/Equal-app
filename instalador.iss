
[Setup]
AppName=EqualApp
AppVersion=1.1
DefaultDirName={pf}\EqualApp
DefaultGroupName=EqualApp
OutputDir=.
OutputBaseFilename=Instalador_EqualApp
Compression=lzma
SolidCompression=yes
DisableDirPage=no
LicenseFile=Licencia.txt
WizardImageFile=banner.bmp
SetupIconFile=GEAN_F.ico

[Files]
Source: "EqualApp.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "chromedriver.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "Instrucciones.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "Licencia.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\EqualApp"; Filename: "{app}\EqualApp.exe"
Name: "{group}\Desinstalar EqualApp"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\EqualApp.exe"; Description: "Ejecutar Buscador de Placas"; Flags: nowait postinstall skipifsilent
