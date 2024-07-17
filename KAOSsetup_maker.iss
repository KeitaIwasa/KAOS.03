; Setup セクション: インストーラーの基本情報を設定します
[Setup]
AppName=KAOS                     
AppVerName=KAOS 03.01            
OutputBaseFilename=KAOSsetup     
VersionInfoDescription=kAOSセットアッププログラム
DefaultDirName={pf}\KAOS
VersionInfoVersion=3.1.0.1
AppCopyright=岩佐デジタル

; Languages セクション: 使用する言語を指定します
[Languages]
Name: japanese; MessagesFile: compiler:Languages\Japanese.isl

; Dirs セクション: インストール先のディレクトリを指定します
[Dirs]
Name: "{app}\_internal";
Name: "{app}\setup"; Permissions: users-modify ;

; Files セクション: インストールするファイルを指定します
[Files]
Source: "dist\KAOS\KAOS.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\KAOS\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "dist\KAOS\setup\*"; DestDir: "{app}\setup"; Flags: ignoreversion recursesubdirs createallsubdirs

; Icons セクション: スタートメニューやデスクトップにショートカットを作成します
[Icons]
Name: "{group}\KAOS"; Filename: "{app}\KAOS.exe"
Name: "{commondesktop}\自動発注システムKAOS"; Filename: "{app}\KAOS.exe"