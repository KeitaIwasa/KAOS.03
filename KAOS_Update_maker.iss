;Inno Setup で使用するファイル

; Setup セクション: インストーラーの基本情報を設定します
[Setup]
AppName=KAOS                     
AppVerName=KAOS 3.6.0            
OutputBaseFilename=KAOS_Update.3.6.0
VersionInfoDescription=kAOSセットアッププログラム
DefaultDirName={pf}\KAOS
VersionInfoVersion=3.6.0.0
AppCopyright=岩佐デジタル
DefaultGroupName=KAOS
DisableStartupPrompt=yes          ; スタートアップ時のプロンプトを無効化
AllowCancelDuringInstall=no       ; インストール中のキャンセルを無効化
AlwaysRestart=no                  ; インストール後に再起動を要求しない
CloseApplications=yes             ; アプリケーションを閉じてからインストール
RestartApplications=yes           ; インストール後にアプリケーションを再起動
SilentInstall=yes                 ; サイレントインストールを有効化
SilentUninstall=yes               ; サイレントアンインストールを有効化

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