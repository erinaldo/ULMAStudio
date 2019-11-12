; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define YEAR "2019"
#define MyAppName "ULMA Studio"
#define MyAppFolder "ULMAStudio"
#define MyAppVersion ".0.0.16"
#define MyAppPublisher "ULMA CONSTRUCTION � 2aCAD Graitec Group (Jos� Alberto Torres Jaraute)"
#define MyAppURL "https://www.ulmaconstruction.com"
#define MyWeb "ULMA CONSTRUCTION"
#define APPDATA "{%APPDATA}"
#define INSTALL "G:\ULMA\INSTALL"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{3EA32971-2F1B-47DA-BA6D-4C16D71C8805}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName="{#APPDATA}\Autodesk\Revit\Addins\{#YEAR}\{#MyAppFolder}"
DisableDirPage=yes
DefaultGroupName=ULMA CONSTRUCTION\{#MyAppName} {#YEAR}
DisableProgramGroupPage=yes
OutputDir="{srcexe}\..\SALIDA"
OutputBaseFilename={#MyAppName}_v{#YEAR}{#MyAppVersion}
SetupIconFile="{srcexe}\..\..\IMAGENES\ULMA.ico "
Compression=lzma
SolidCompression=yes
UninstallDisplayIcon={uninstallexe}
VersionInfoVersion={#YEAR}{#MyAppVersion}
VersionInfoCompany=ULMA CONSTRUCTION
VersionInfoDescription=Utilidades REVIT {#YEAR} para Ulma Construction
VersionInfoCopyright=ULMA CONSTRUCTION � 2aCAD Global Group (Jos� Alberto Torres Jaraute)
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#YEAR}{#MyAppVersion}
ArchitecturesAllowed=x64

;[Languages]
;Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
UsePreviousAppDir=False
WizardSmallImageFile="{srcexe}\..\..\IMAGENES\InstallerSmallImage.bmp"
WizardImageFile="{srcexe}\..\..\IMAGENES\InstallerCompleting.bmp"
AllowCancelDuringInstall=False
ShowLanguageDialog=no
UninstallLogMode=new

[Files]
; *** APP principal y fichero configuracion
Source: "{#INSTALL}\*.*"; DestDir: "{app}\..\"; Flags: ignoreversion createallsubdirs recursesubdirs
Source: "{#INSTALL}\..\ULMAUpdaterAddin2018.ini"; DestDir: "{app}"; DestName: "ULMAUpdaterAddin.ini"; Flags: ignoreversion createallsubdirs recursesubdirs

[Icons]
Name: "{group}\{cm:ProgramOnTheWeb,{#MyWeb}}"; Filename: "{#MyAppURL}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
;Name: "{group}\{#MyAppName}.ini"; Filename: "{app}\{#MyAppName}.ini"; WorkingDir: "{app}"; IconFilename: "{app}\{#MyAppName}.ini"
Name: "{group}\{#MyAppName}"; Filename: "{app}"; WorkingDir: "{app}"

[UninstallDelete]
Type: files; Name: "{app}\..\{#MyAppName}.addin"
Type: filesandordirs; Name: "{app}\*.*" 
Type: dirifempty; Name: "{app}"

[ThirdParty]
UseRelativePaths=True