[Setup]
AppName=BDOpenOffice
AppVersion=1.0
DefaultDirName={localappdata}\BDOpenOffice
DefaultGroupName=BDOpenOffice
OutputBaseFilename=bdopenoffice-installer
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
DisableDirPage=yes

[Files]
; 修改 Source 为你的发布目录（注意不要带 {} 宏在 Source 路径里）
Source: "publish\win-x64\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\BDOpenOffice"; Filename: "{app}\bdopenoffice.exe"

[Run]
; 安装完成后以用户上下文注册协议（静默或可见由 Flags 控制）
Filename: "{app}\bdopenoffice.exe"; Parameters: "--register"; Flags: waituntilterminated runhidden skipifsilent

[UninstallRun]
; 卸载时注销协议
Filename: "{app}\bdopenoffice.exe"; Parameters: "--unregister"; Flags: runhidden waituntilterminated