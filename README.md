# mkpsh

RHEL / AIX / Windows 自動パラメータシート生成ツール。

OSのパラメータを収集し、HTML形式でパラメータ情報を出力します。

## 対応OS

* RHEL 7 / 8
* CentOS 7 / 8
* AIX 7
* Windows Server 2019

## 使用方法

### RHEL / AIX

適当なディレクトリにコピーし、管理者権限で引数無しで実行するとカレントディレクトリに「ホスト名.html」ファイルが生成されます。
```
# mkpsh
```

### Windows

適当なディレクトリにコピーし、管理者権限で引数無しで実行します。実行結果は標準出力に出力されるのでリダイレクトでファイルに出力してください。
```
PS> mkpsh.ps1 > hogehoge.html
```
