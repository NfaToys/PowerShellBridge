[Setup]
AppName=PowerShellBridge
AppVersion=1.0.0.0
DefaultDirName={pf64}\PowerShellBridge
OutputBaseFilename=PowerShellBridgeSetupX64
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
Compression=lzma
SolidCompression=yes

[Files]
Source: "C:\Users\nfato\source\repos\PowerShellBridge\PowerShellBridge\bin\Release\PowerShellBridge.dll"; DestDir: "{app}"; Flags: ignoreversion

[Run]
Filename: "{win}\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"; Parameters: """{app}\PowerShellBridge.dll"" /codebase /tlb"; StatusMsg: "Registering PowerShellBridge COM components..."; Flags: runhidden

[UninstallRun]
Filename: "{win}\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"; Parameters: """{app}\PowerShellBridge.dll"" /unregister /tlb"; StatusMsg: "Unregistering PowerShellBridge COM components..."; Flags: runhidden