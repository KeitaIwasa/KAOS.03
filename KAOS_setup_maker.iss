;Inno Setup で使用するファイル

; Setup セクション: インストーラーの基本情報を設定します
[Setup]
AppName=KAOS                     
AppVerName=KAOS 3.7.0
OutputBaseFilename=KAOS_setup.3.7.0.4
VersionInfoDescription=kAOSセットアッププログラム
DefaultDirName={pf}\KAOS
VersionInfoVersion=3.7.0.4
AppCopyright=岩佐デジタル
CloseApplications=yes
RestartApplications=yes

; Languages セクション: 使用する言語を指定します
[Languages]
Name: japanese; MessagesFile: compiler:Languages\Japanese.isl

; Dirs セクション: インストール先のディレクトリを指定します
[Dirs]
Name: "{app}\_internal";
Name: "{app}\error_log"; Permissions: users-modify ;
Name: "{app}\setup"; Permissions: users-modify ;

; Files セクション: インストールするファイルを指定します
[Files]
Source: "dist\KAOS\KAOS.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\KAOS\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs

; Icons セクション: スタートメニューやデスクトップにショートカットを作成します
[Icons]
Name: "{group}\KAOS"; Filename: "{app}\KAOS.exe"
Name: "{commondesktop}\自動発注システムKAOS"; Filename: "{app}\KAOS.exe"

; Run セクション: アプリをインストール後に自動で起動します。
[Run]
Filename: "{app}\KAOS.exe"; Flags: nowait