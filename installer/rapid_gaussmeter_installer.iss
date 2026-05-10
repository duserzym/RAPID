; Inno Setup 6 script for RapidPy Gaussmeter Control
; Build after running PyInstaller for both specs:
;   pyinstaller installer\rapid_gaussmeter.spec --noconfirm
;   pyinstaller installer\install_fwbell_drivers.spec --noconfirm
;   iscc installer\rapid_gaussmeter_installer.iss
;
; Download Inno Setup: https://jrsoftware.org/isinfo.php

#define AppName      "RapidPy Gaussmeter"
#define AppVersion   "1.0.0"
#define AppPublisher "RAPID Lab"
#define AppExeName   "RapidPy_Gaussmeter.exe"
#define GaussExe     "..\dist\RapidPy_Gaussmeter.exe"
#define DrvInstExe   "..\dist\install_fwbell_drivers.exe"

[Setup]
AppId={{A3F2C1D0-7B4E-4A9F-8C23-1E5D6F7A8B9C}}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\RapidPy\Gaussmeter
DefaultGroupName=RapidPy
AllowNoIcons=yes
OutputDir=..\dist\installer
OutputBaseFilename=RapidPy_Gaussmeter_Setup_{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
; x86 and x64 – the helper exe is x86 but runs fine on x64 Windows
ArchitecturesAllowed=x86 x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";   Description: "{cm:CreateDesktopIcon}";   GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "installdrivers"; Description: "Install FW Bell USB driver files (usb5100.dll + libusb0.dll)"; GroupDescription: "Driver Setup:"; Flags: checkedonce

[CustomMessages]
DllPageTitle=FW Bell Driver Files (Optional)
DllPageDesc=Locate the FW Bell vendor DLLs so the app can find them automatically.
DllNote=These files come from the FW Bell PC5180 software package:%n  usb5100.dll%n  libusb0.dll%nIf you skip this step you can run the driver installer separately from the Start Menu.

[Files]
; Main application (one-file bundle)
Source: "{#GaussExe}"; DestDir: "{app}"; Flags: ignoreversion

; Standalone driver installer (one-file bundle)
Source: "{#DrvInstExe}"; DestDir: "{app}\driver_installer"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}";                   Filename: "{app}\{#AppExeName}"
Name: "{group}\Install FW Bell Drivers";      Filename: "{app}\driver_installer\install_fwbell_drivers.exe"
Name: "{group}\Uninstall {#AppName}";         Filename: "{uninstallexe}"
Name: "{commondesktop}\{#AppName}";           Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
; Run the driver installer silently if the user opted in during setup (already elevated).
Filename: "{app}\driver_installer\install_fwbell_drivers.exe"; Parameters: "--silent"; \
  Description: "Install FW Bell USB driver files now"; \
  Flags: runascurrentuser nowait; \
  Tasks: installdrivers

; Offer to launch the app after setup completes.
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

; ── Pascal script ─────────────────────────────────────────────────────────────
[Code]

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
end;
