echo off
echo 〇〇のデータ更新バッチを適用します。
pause
rem コピー元
Set SourcePath="\\foo\bar\contents\"
rem コピー先
Set CopyToPath="C:\Program Files\baz\contents\"
rem 共有フォルダでは管理者実行できない
rem そのため一度ユーザーフォルダにbatファイルを作成してそれを実行するようにする
cd %USERPROFILE%
echo robocopy /e /xo %SourcePath%  %CopyToPath% > CopyContents.bat
rem  作ったbatファイルで管理者実行
powershell start-process CopyContents.bat -verb runas -wait
rem 作成したbatファイル削除
del CopyContents.bat
echo 〇〇のデータ更新バッチの適用が完了しました。
pause

