; build/clipgen.iss — per-user Windows installer around the PyInstaller one-dir
; output. Compiled in CI by ISCC (preinstalled on the windows-latest image):
;     ISCC.exe /DAppVer=<tag-or-dev> build\clipgen.iss
; Source/Output paths below are relative to this file's directory (build/).
; The version is injected via /DAppVer — never hardcoded here (build/VERSION is
; the single version source; CI derives the value from the tag, or "dev").

#ifndef AppVer
  #error AppVer must be defined on the ISCC command line, e.g. /DAppVer=v1.2.3
#endif

[Setup]
; Fixed GUID identifying this app to Windows across upgrades. Never change it,
; or upgrades stop replacing the previous install.
AppId={{F5720C20-F697-47A6-9AFA-CEC80348A190}
AppName=clipgen
AppVersion={#AppVer}
AppPublisher=Signal Research
AppPublisherURL=https://github.com/henedl/clipgen
; Per-user install: no admin prompt, no UAC elevation, and the standard
; location for per-user apps ({localappdata}\Programs, same as VS Code).
PrivilegesRequired=lowest
DefaultDirName={localappdata}\Programs\clipgen
DefaultGroupName=clipgen
DisableProgramGroupPage=yes
; Land the setup exe in dist/ beside the zip so the release globs stay
; anchored to dist/ and fail_on_unmatched_files keeps meaning something.
OutputDir=..\dist
OutputBaseFilename=clipgen-{#AppVer}-setup
SetupIconFile=clipgen.ico
UninstallDisplayIcon={app}\clipgen.exe
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; The bundle ships GPL components; surface the notices in the wizard. This is
; an info page, not a click-through EULA — clipgen itself is MIT.
InfoBeforeFile=THIRD-PARTY-LICENSES

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; dist/clipgen already contains INSTALL.txt and THIRD-PARTY-LICENSES by the
; time ISCC runs (the zip step copies them in first), so they install too.
Source: "..\dist\clipgen\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{autoprograms}\clipgen"; Filename: "{app}\clipgen.exe"
Name: "{autodesktop}\clipgen"; Filename: "{app}\clipgen.exe"; Tasks: desktopicon
