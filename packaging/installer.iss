; 审计工具箱 安装程序（Inno Setup 6）
;
; 前置：先在工具箱根目录用 PyInstaller 出便携版
;   .venv/Scripts/python.exe -m PyInstaller --noconfirm packaging/wxsync.spec
; 再编译安装程序（在工具箱根目录执行）：
;   ".tools\Inno Setup 6\ISCC.exe" packaging/installer.iss
; 产物：dist/审计工具箱_v<版本>_安装程序.exe（与便携版文件夹同目录并存）
;
; 升级：AppId 不变，直接运行新安装程序即原地覆盖升级，快捷方式与卸载项保留。
; 目录：默认按用户安装（免管理员权限），可写入 exe 旁的 .wesync_ui.json；
;       向导里可改到其他位置（如 D:\审计工具箱），装到 C:\Program Files 会需管理员且偏好文件写不进去。

#define MyAppName "审计工具箱"
#define MyAppVersion "1.11.1"
#define MySrcDir "..\dist\" + MyAppName + "_v" + MyAppVersion
#define MyExeName MyAppName + ".exe"

[Setup]
AppId={{A62B3C4D-5E6F-4A7B-8C9D-0E1F2A3B4C5D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
DefaultDirName={autopf}\{#MyAppName}
DisableDirPage=no
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\dist
OutputBaseFilename={#MyAppName}_v{#MyAppVersion}_安装程序
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
SetupIconFile=..\excel_word\wesync\resources\audit_toolbox.ico
ShowLanguageDialog=no
UninstallDisplayName={#MyAppName} {#MyAppVersion}
UninstallDisplayIcon={app}\{#MyExeName}
CloseApplications=yes

[Languages]
Name: "chs"; MessagesFile: "languages\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加任务："

[Files]
Source: "{#MySrcDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyExeName}"
Name: "{group}\使用说明"; Filename: "{app}\使用说明.txt"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyExeName}"; Description: "立即运行{#MyAppName}"; Flags: nowait postinstall skipifsilent
