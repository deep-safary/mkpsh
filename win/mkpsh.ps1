#-------------------------------------------------------------------------------
#   mkpsh.ps1
#       Create a parameter-sheet html file.
#       Supported Windows Server 2016 or higher
#
#-------------------------------------------------------------------------------

#-------------------------------------------------------------------------------
#   fnc_htmlheader
#       HTMLのヘッダ要素を出力する
#-------------------------------------------------------------------------------
function fnc_htmlheader {
    echo @"
<!doctype html>
<html lang="ja">
<head>
<title>${HOSTNAME} Parameter Sheet</title>
<meta charset="UTF-8">
<style>
* {
  font-family:      'MS gothic',monospace;
}

html {
  font-size:        0.8em;
  font-family:      'MS gothic',monospace;
}

body {
  font-size:        0.9em;
  font-family:      'MS gothic',monospace;
}

h1 {
  background-color: #006099;
  border-style:     outset;
  color:            white;
  margin-top:       20px;
  margin-bottom:    10px;
  padding:          2px;
}

h2 {
  background-color: #009999;
  border-style:     outset;
  color:            white;
  margin-top:       20px;
  margin-bottom:    10px;
  padding:          2px;
}

h3 {
  background-color: gray;
  border-style:     outset;
  color:            white;
  margin-top:       20px;
  margin-bottom:    10px;
  padding:          2px;
}

table {
  border:           solid 1px;
  border-collapse:  collapse;
  margin:           6px;
  width:            99%;
}

td {
  border:           solid 1px;
  padding:          4px;
  vertical-align:   top;
}

th {
  background-color: #ffff00;
  border:           solid 1px;
  padding:          4px;
  width:            120px;
}

textarea {
  background-color: #f0ffff;
  margin:           6px;
  padding:          6px;
  width:            98%;
}

div#site-block {
  width:            100%;
  height:           100%;
}

nav {
  float:            left;
  height:           100%;
  overflow-y:       scroll;
  padding:          5px;
  position:         fixed;
  width:            15%;
}

div#right-block {
  float:            right;
  margin-bottom:    80px;
  padding-left:     10px;
  width:            83%;
}

div#bottom {
  margin-bottom:    80px;
}

div#date {
  margin-bottom:    80px;
}

</style>
</head>

<body class="container">

"@
}


#-------------------------------------------------------------------------------
#   fnc_htmlfooter
#       HTMLの最終行を出力する
#-------------------------------------------------------------------------------
function fnc_htmlfooter {
    echo @"
<div id="bottom"></div>
</div>
</div>
</body>
"@
}


#-------------------------------------------------------------------------------
#   fnc_index
#       HTML左側のINDEXを出力する
#-------------------------------------------------------------------------------
function fnc_index {
    $buf = Get-Date

    echo @"
<nav>
<h1>$hostname</h1>
<h1>INDEX</h1>
<details>
  <summary>systeminfo</summary>
  <li><a href="#systeminfo">systeminfo</a>
</details>
<br>

<details>
  <summary>システムのプロパティ</summary>
  <li><a href="#オペレーティングシステム">オペレーティングシステム</a>
  <li><a href="#コンピューター名/ドメイン名の変更">コンピューター名/ドメイン名の変更</a>
  <li><a href="#DNSサフィックスとNetBIOSコンピューター名">DNSサフィックスとNetBIOSコンピューター名</a>
  <li><a href="#ハードウェア">ハードウェア</a>
  <li><a href="#パフォーマンスオプション">パフォーマンスオプション</a>
  <li><a href="#起動と回復">起動と回復</a>
  <li><a href="#環境変数">環境変数</a>
  <li><a href="#リモート">リモート</a>
</details>
<br>

<details>
  <summary>時計と地域</summary>
  <li><a href="#日付と時刻">日付と時刻</a>
  <li><a href="#インターネット時刻">インターネット時刻</a>
</details>
<br>

<details>
  <summary>時刻同期サーバーサービス</summary>
  <li><a href="#NTPサーバー設定">NTPサーバー設定</a>
</details>
<br>

<details>
  <summary>イベントログ</summary>
  <li><a href="#Windowsログ">Windowsログ</a>
</details>
<br>

<details>
  <summary>Windows Update</summary>
  <li><a href="#Windows Update">Windows Update</a>
  <li><a href="#WSUS">WSUS</a>
</details>
<br>

<details>
  <summary>タスクスケジューラ</summary>
  <li><a href="#タスクスケジューラ">タスクスケジューラ</a>
</details>
<br>

<details>
  <summary>IEセキュリティ強化の構成</summary>
  <li><a href="#IEセキュリティ強化の構成">IEセキュリティ強化の構成</a>
</details>
<br>

<details>
  <summary>電源プラン</summary>
  <li><a href="#電源プラン">電源プラン</a>
</details>
<br>

<details>
  <summary>ネットワークアダプタ</summary>
  <li><a href="#接続名称とデバイス情報">接続名称とデバイス情報</a>
  <li><a href="#チーミングデバイス情報">チーミングデバイス情報</a>
</details>
<br>

<details>
  <summary>ネットワーク</summary>
  <li><a href="#接続プロパティ">接続プロパティ</a>
  <li><a href="#Internet Protocol Version 4 (TCP/IPv4)のプロパティ">Internet Protocol Version 4 (TCP/IPv4)のプロパティ</a>
  <li><a href="#DNSサーバーアドレス">DNSサーバーアドレス</a>
  <li><a href="#hosts">hosts</a>
  <li><a href="#静的経路">静的経路</a>
  <li><a href="#Scalable Networking Pack設定">Scalable Networking Pack設定</a>
  <li><a href="#SMBマルチチャネル機能">SMBマルチチャネル機能</a>
</details>
<br>

<details>
  <summary>ファイヤーウォール</summary>
  <li><a href="#プロファイル">プロファイル</a>
  <li><a href="#ファイアウォールルール">ファイアウォールルール</a>
</details>
<br>

<details>
  <summary>ディスク</summary>
  <li><a href="#ディスクデバイス">ディスクデバイス</a>
  <li><a href="#ボリューム/パーティション">ボリューム/パーティション</a>
</details>
<br>

<details>
  <summary>アカウント</summary>
  <li><a href="#グループ">グループ</a>
  <li><a href="#ユーザー">ユーザー</a>
  <li><a href="#パスワード/ロックアウトポリシー">パスワード/ロックアウトポリシー</a>
  <li><a href="#ユーザーアカウント制御(UAC)">ユーザーアカウント制御(UAC)</a>
</details>
<br>

<details>
  <summary>プログラムと機能</summary>
  <li><a href="#導入プログラム">導入プログラム</a>
  <li><a href="#インストールされた更新プログラム">インストールされた更新プログラム</a>
  <li><a href="#役割と機能">役割と機能</a>
</details>
<br>

<details>
  <summary>サービス</summary>
  <li><a href="#サービス">サービス</a>
</details>
<br>

<div id="date">
作成日時:<br>
$buf
</div>
</nav>



<div id="right-block">
"@
}


#-------------------------------------------------------------------------------
#   fnc_systeminfo
#       systeminfo情報を出力
#-------------------------------------------------------------------------------
function fnc_systeminfo {
    echo @"
<h2><a name="systeminfo">systeminfo</a></h2>
<table>
  <tr><th>確認コマンド</th> <td>systeminfo.exe</td></tr>
</table>
<textarea  rows="10" wrap="off" readonly>
"@
    systeminfo.exe
    echo @"
</textarea>

"@
}


#-------------------------------------------------------------------------------
#   fnc_operatingsystem
#       コンピュータ名
#       コンピュータの説明
#       Windowsのエディション
#       バージョン
#       アーキテクチャ
#-------------------------------------------------------------------------------
function fnc_operatingsystem {
    $buf = Get-WmiObject Win32_OperatingSystem
    $desctiption = $buf.Description
    $caption = $buf.Caption
    $version = $buf.Version
    $architect = $buf.OSArchitecture

    echo @"
<h2><a name="オペレーティングシステム">オペレーティングシステム</a></h2>
<h3>OS情報</h3>
<table>
  <tr><th>項目</th> <th>値</th>   <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>コンピュータ名</th>     <td>$hostname</td> <td>hostname</td> <td></td> <td></td></tr>
  <tr><th>コンピュータの説明</th> <td>$description</td> <td>Get-WmiObject Win32_OperatingSystem</td> <td>Description</td> <td></td></tr>
  <tr><th>エディション</th>       <td>$caption</td> <td>Get-WmiObject Win32_OperatingSystem</td> <td>Caption</td> <td></td></tr>
  <tr><th>バージョン</th>         <td>$version</td> <td>Get-WmiObject Win32_OperatingSystem</td> <td>Version</td> <td></td></tr>
  <tr><th>アーキテクチャ</th>     <td>$architect</td> <td>Get-WmiObject Win32_OperatingSystem</td> <td>OSArchitecture</td> <td></td></tr>
</table>

"@

    # 言語
    $locale = (Get-WinSystemLocale).DisplayName
    # 時刻と通貨
    $culture = (Get-Culture).DisplayName
    # キーボード
    $keyboard = (Get-WmiObject Win32_Keyboard).Description

    echo @"
<h3>環境情報</h3>
<table>
  <tr><th>項目</th>       <th>値</th>   <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>言語</th>       <td>$locale</td> <td>Get-WinSystemLocale</td> <td>DisplayName</td> <td></td></tr>
  <tr><th>時刻と通貨</th> <td>$culture</td> <td>Get-Culture</td> <td>DisplayName</td> <td></td></tr>
  <tr><th>キーボード</th> <td>$keyboard</td> <td>Get-WmiObject Win32_Keyboard</td> <td>Description</td> <td></td></tr>
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_workgroup
#       ドメイン/ワークグループ
#       所属するグループ
#-------------------------------------------------------------------------------
function fnc_workgroup {
    $obj = get-wmiobject win32_computersystem
    $group = $obj.PartOfDomain
    $name = $obj.Domain

    echo @"
<h2><a name="コンピューター名/ドメイン名の変更">コンピューター名/ドメイン名の変更</a></h2>
<table>
  <tr><th>項目</th> <th>値</th>     <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>所属</th> <td>$group</td> <td>get-wmiobject win32_computersystem</td> <td>PartOfDomain</td> <td>True = DOMAIN<br>False = WORKGROUP</td></tr>
  <tr><th>名前</th> <td>$name</td>  <td>get-wmiobject win32_computersystem</td> <td>Domain</td> <td></td></tr>
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_dnssuffix
#       このコンピュータのプライマリ DNS サフィックス
#       ドメインのメンバシップが変更されるときにプライマリ DNS サフィックスを変更する
#-------------------------------------------------------------------------------
function fnc_dnssuffix {
    $obj = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\services\Tcpip\Parameters'
    $buf1 = $obj.'NV Domain'
    $buf2 = $obj.SyncDomainWithMembership

    echo @"
<h2><a name="DNSサフィックスとNetBIOSコンピューター名">DNSサフィックスとNetBIOSコンピューター名</a></h2>
<table>
  <tr><th>項目</th> <th>値</th>   <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>このコンピュータのプライマリ DNS サフィックス</th> <td>$buf1</td> <td>Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\services\Tcpip\Parameters'</td> <td>NV Domain</td> <td></td></tr>
  <tr><th>ドメインのメンバシップが変更されるときにプライマリ DNS サフィックスを変更する</th> <td>$buf2</td> <td>Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\services\Tcpip\Parameters'</td> <td>SyncDomainWithMembership</td> <td>0 = チェックオフ<br>1 = チェックオン</td></tr>
</table>

"@
}



#-------------------------------------------------------------------------------
#   fnc_hardware
#       ハードウェア情報
#-------------------------------------------------------------------------------
function fnc_hardware {
    $proc = (Get-WmiObject Win32_Processor).Name
    $mem = (Get-WmiObject Win32_OperatingSystem).TotalVisibleMemorySize.ToString("#,0 KB")
    $buf = Get-WmiObject Win32_PnpEntity | Sort-Object -Property Name
    $dev = $buf.Name

    echo @"
<h2><a name="ハードウェア">ハードウェア</a></h2>
<h3><a name="プロセッサ・メモリ">プロセッサ・メモリ</a></h3>
<table>
  <tr><th>項目</th> <th>値</th>   <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>プロセッサ</th> <td>$proc</td> <td>Get-WmiObject Win32_Processor</td> <td>Name</td> <td></td></tr>
  <tr><th>メモリ</th>     <td>$mem</td> <td>Get-WmiObject Win32_OperatingSystem</td> <td>TotalVisibleMemorySize</td> <td></td></tr>
</table>

<h3><a name="デバイス">デバイス</a></h3>
<table>
  <tr><th>コマンドレット</th> <td>Get-WmiObject Win32_PnpEntity</td> <th>プロパティ</th> <td>Name</td></tr>
</table>
<textarea  rows="10" wrap="off" readonly>
"@
$dev
    echo @"
</textarea>

"@
}


#-------------------------------------------------------------------------------
#   fnc_visualeffect
#       視覚効果
#-------------------------------------------------------------------------------
function fnc_visualeffect {
    $obj = (Get-ItemProperty -Path HKLM:System\CurrentControlSet\Control\CrashControl).VisualFXSetting

    echo @"
<h2><a name="パフォーマンスオプション">パフォーマンスオプション</a></h2>
<h3>視覚効果</h3>
<table>
  <tr><th>項目</th> <th>値</th>   <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>視覚効果</th> <td>$obj</td> <td>Get-ItemProperty -Path HKLM:System\CurrentControlSet\Control\CrashControl</td> <td>VisualFXSetting</td> <td>
    null or 0 = コンピュータに応じて最適なものを自動的に選択する【既定】<br>
    1 = デザインを優先する<br>
    2 = パフォーマンスを優先する<br>
    3 = カスタム</td></tr>
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_prioritycontrol
#       プロセッサのスケジュール
#-------------------------------------------------------------------------------
function fnc_prioritycontrol {
    $obj = [string](Get-ItemProperty -Path HKLM:SYSTEM\CurrentControlSet\Control\PriorityControl).Win32PrioritySeparation

    echo @"
<h3>プロセッサのスケジュール</h3>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>プロセッサのスケジュール</th> <td>$obj</td> <td>Get-ItemProperty -Path HKLM:SYSTEM\CurrentControlSet\Control\PriorityControl</td> <td>Win32PrioritySeparation</td> <td>38 = プログラム<br>2 or 24 = バックグラウンドサービス</td></tr>
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_recoveros
#       システムログにイベントを書き込む
#       自動的に再起動する
#       デバッグ情報の書き込み
#       ダンプ先
#       既存のファイルに上書きする
#-------------------------------------------------------------------------------
function fnc_recoveros {

    $obj = Get-WmiObject Win32_OSRecoveryConfiguration
    $WriteToSystemLog = $obj.WriteToSystemLog
    $AutoReboot = $obj.AutoReboot
    $DebugInfoType = $obj.DebugInfoType
    $DebugFilePath = $obj.DebugFilePath
    $OverwriteExistingDebugFile =  $obj.OverwriteExistingDebugFile
    $AlwaysKeepMemoryDump = (Get-ItemProperty -Path HKLM:System\CurrentControlSet\Control\CrashControl).AlwaysKeepMemoryDump

    echo @"
<h2><a name="起動と回復">起動と回復</a></h2>
<h3>システムエラー</h3>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>システムログにイベントを書き込む</th> <td>$WriteToSystemLog</td> <td>Get-WmiObject Win32_OSRecoveryConfiguration</td> <td>WriteToSystemLog</td> <td></td></tr>
  <tr><th>自動的に再起動する</th> <td>$AutoReboot</td> <td>Get-WmiObject Win32_OSRecoveryConfiguration</td> <td>AutoReboot</td> <td></td></tr>
  <tr><th>デバッグ情報の書き込み</th> <td>$DebugInfoType</td> <td>Get-WmiObject Win32_OSRecoveryConfiguration</td> <td>DebugInfoType</td> <td>
    1 = 完全なメモリダンプ<br>
    2 = カーネルメモリダンプ<br>
    3 = 最小メモリダンプ<br>
    7 = 自動メモリダンプ
  </td></tr>
  <tr><th>ダンプ先</th> <td>$DebugFilePath</td> <td>Get-WmiObject Win32_OSRecoveryConfiguration</td> <td>DebugFilePath</td> <td></td></tr>
  <tr><th>既存のファイルに上書きする</th> <td>$OverwriteExistingDebugFile</td> <td>Get-WmiObject Win32_OSRecoveryConfiguration</td> <td>OverwriteExistingDebugFile</td> <td></td></tr>
  <tr><th>ディスク領域が少ないときでもメモリダンプの自動削除を無効にする</th> <td>$AlwaysKeepMemoryDump</td> <td>Get-ItemProperty -Path HKLM:System\CurrentControlSet\Control\CrashControl</td> <td>AlwaysKeepMemoryDump</td> <td>null or 0 = チェックオフ<br>1 = チェックオン</td></tr>
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_env
#       環境変数
#-------------------------------------------------------------------------------
function fnc_env {
    $obj = Get-ChildItem env:

    echo @"
<h2><a name="環境変数">環境変数</a></h2>
<table>
  <tr><th>コマンドレット</th> <td>Get-ChildItem env:</td> <th>プロパティ</th> <td>Name,Value</td></tr>
</table>

<div style="height:300px; width=100%; overflow-y:scroll;">
<table>
  <tr><th>環境変数</th> <th>値</th> <th>備考</th></tr>
"@

    foreach($line in $obj){
        Write-Output "  <tr><td>"$line.Name"</td> <td>"$line.Value"</td> <td></td></tr>"
    }

    echo @"
</table>
</div>

"@
}


#-------------------------------------------------------------------------------
#   fnc_remote_desktop
#       リモートデスクトップ
#-------------------------------------------------------------------------------
function fnc_remote_desktop {
    $fDenyTSConnections = (Get-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Terminal Server").fDenyTSConnections
    $UserAuthentication = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp").UserAuthentication
    $obj = net localgroup "remote desktop users"
    $buf = $obj[4]

    echo @"
<h2><a name="リモート">リモート</a></h2>
<h3>リモートデスクトップ</h3>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>このコンピューターへのリモート接続を許可</th> <td>$fDenyTSConnections</td> <td>Get-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Terminal Server"</td> <td>fDenyTSConnections</td> <td>
    0 = リモートデスクトップを実行しているコンピュータからの接続を許可する<br>
    1 = このコンピュータへの接続を許可しない
  </td></tr>
  <tr><th>ネットワークレベル認証でリモートデスクトップを実行しているコンピュータからのみ接続を許可する</th> <td>$fDenyTSConnections</td> <td>Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp"</td> <td>UserAuthentication</td> <td>0 = 無効【既定】<br>1 = 有効</td></tr>
</table>

"@

    echo @"
<h3>リモートデスクトップユーザー</h3>
<table>
  <tr><th>確認コマンド</th> <td>net localgroup "remote desktop users"</td></tr>
</table>
<textarea  rows="10" wrap="off" readonly>
"@
    net localgroup "remote desktop users"
    echo @"
</textarea>

"@
}


#-------------------------------------------------------------------------------
#   fnc_timezone
#       タイムゾーン
#-------------------------------------------------------------------------------
function fnc_timezone {
    $timezone = (Get-TimeZone).DisplayName
    $server = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters").NtpServer
    $type = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters").Type
    $SpecialPollInterval = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient").SpecialPollInterval
    $MaxAllowedPhaseOffset = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config").MaxAllowedPhaseOffset

    echo @"
<h2><a name="日付と時刻">日付と時刻</a></h2>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>タイムゾーン</th> <td>$timezone</td> <td>Get-Timezone</td> <td>DisplayName</td> <td></td></tr>
</table>

<h2><a name="インターネット時刻">インターネット時刻</a></h2>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>サーバー</th> <td>$server</td> <td>Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters"</td> <td>NtpServer</td> <td>
    0x1 = Symmetric Activeモードで同期／Windowsで実装した一定間隔での同期<br>
    0x2 = Symmetric Activeモードで同期／フォールバック時に利用するNTPサーバを指定<br>
    0x4 = Symmetric Activeモードで同期／RFC1305に準拠した間隔での同期<br>
    0x8 = Clientモードで同期／RFC1305に準拠した間隔での同期<br>
    0x9 = 同期しない
  </td></tr>
  <tr><th>同期のタイプ</th> <td>$type</td> <td>Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters"</td> <td>Type</td> <td></td></tr>
  <tr><th>時刻同期間隔</th> <td>$SpecialPollInterval</td> <td>Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient"</td> <td>SpecialPollInterval</td> <td>GUI設定不可</td></tr>
  <tr><th>slewモード有効範囲時間</th> <td>$MaxAllowedPhaseOffset</td> <td>Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config"</td> <td>MaxAllowedPhaseOffset</td> <td>GUI設定不可</td></tr>
</table>

"@
}



#-------------------------------------------------------------------------------
#   fnc_ntpserver
#       NTPサーバ指定(WindowsにGUIによる表示・設定方法なし)
#-------------------------------------------------------------------------------
function fnc_ntpserver {

    $ntpenabled = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpServer").Enabled
    $AnnounceFlags = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\services\W32Time\Config").AnnounceFlags

    echo @"
<h2><a name="NTPサーバー設定">NTPサーバー設定</a></h2>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>NTPサーバー機能の有効化</th> <td>$ntpenabled</td> <td>Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\w32time\TimeProviders\NtpServer"</td> <td>Enabled</td> <td>0 = 無効<br>1 = 有効</td></tr>
  <tr><th>タイムサーバの動作</th> <td>$AnnounceFlags</td> <td>Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\services\W32Time\Config"</td> <td>AnnounceFlags</td> <td>
    0x00 = タイム サーバーではありません<br>
    0x01 = 常にタイム サーバーです<br>
    0x02 = 自動タイム サーバー<br>
    0x04 = 常に信頼できるタイム サーバーです<br>
    0x08 = 信頼できる自動タイム サーバーです
  </td></tr>
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_systemlog
#       Windowsログ
#-------------------------------------------------------------------------------
function fnc_systemlog {

    # アプリケーション
    $obj = Get-WinEvent -listlog Application
    $LogFilePath = $obj.LogFilePath
    $MaximumSizeInBytes = $obj.MaximumSizeInBytes
    $LogMode = $obj.LogMode

    echo @"
<h2><a name="Windowsログ">Windowsログ</a></h2>
<h3>アプリケーション</h3>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>ログのパス</th> <td>$LogFilePath</td> <td>Get-WinEvent -listlog Application</td> <td>LogFilePath</td> <td></td></tr>
  <tr><th>最大ログサイズ</th> <td>$MaximumSizeInBytes</td> <td>Get-WinEvent -listlog Application</td> <td>MaximumSizeInBytes</td> <td></td></tr>
  <tr><th>イベントログサイズが最大値に達したとき</th> <td>$LogMode</td> <td>Get-WinEvent -listlog Application</td> <td>LogMode</td> <td>
    Circular = 必要に応じて上書きする<br>
    AutoBackup = 上書きせず保存する<br>
    Retain = 上書きしない
  </td></tr>
</table>

"@

    # セキュリティ
    $obj = Get-WinEvent -listlog Security
    $LogFilePath = $obj.LogFilePath
    $MaximumSizeInBytes = $obj.MaximumSizeInBytes
    $LogMode = $obj.LogMode

    echo @"
<h3>セキュリティ</h3>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>ログのパス</th> <td>$LogFilePath</td> <td>Get-WinEvent -listlog Security</td> <td>LogFilePath</td> <td></td></tr>
  <tr><th>最大ログサイズ</th> <td>$MaximumSizeInBytes</td> <td>Get-WinEvent -listlog Security</td> <td>MaximumSizeInBytes</td> <td></td></tr>
  <tr><th>イベントログサイズが最大値に達したとき</th> <td>$LogMode</td> <td>Get-WinEvent -listlog Security</td> <td>LogMode</td> <td>
    Circular = 必要に応じて上書きする<br>
    AutoBackup = 上書きせず保存する<br>
    Retain = 上書きしない
  </td></tr>
</table>

"@

    # システム
    $obj = Get-WinEvent -listlog System
    $LogFilePath = $obj.LogFilePath
    $MaximumSizeInBytes = $obj.MaximumSizeInBytes
    $LogMode = $obj.LogMode

    echo @"
<h3>システム</h3>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>ログのパス</th> <td>$LogFilePath</td> <td>Get-WinEvent -listlog System</td> <td>LogFilePath</td> <td></td></tr>
  <tr><th>最大ログサイズ</th> <td>$MaximumSizeInBytes</td> <td>Get-WinEvent -listlog System</td> <td>MaximumSizeInBytes</td> <td></td></tr>
  <tr><th>イベントログサイズが最大値に達したとき</th> <td>$LogMode</td> <td>Get-WinEvent -listlog System</td> <td>LogMode</td> <td>
    Circular = 必要に応じて上書きする<br>
    AutoBackup = 上書きせず保存する<br>
    Retain = 上書きしない
  </td></tr>
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_windowsupdate
#       自動更新を構成する
#-------------------------------------------------------------------------------
function fnc_windowsupdate {

    # 自動更新を構成する
    $obj = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
    $NoAutoUpdate = $obj.NoAutoUpdate

    echo @"
<h2><a name="Windows Update">Windows Update</a></h2>
<h3>自動更新</h3>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>自動更新を構成する</th> <td>$NoAutoUpdate</td> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"</td> <td>NoAutoUpdate</td> <td>0 = 有効<br>1 = 無効<br>null or 2 = 未構成【既定】</td></tr>
</table>

"@

    if ( -not ($NoAutoUpdate -eq 0) ) {
        return
    }

    # 自動更新の構成
    # (WIN_UPDATEに"0"を指定した場合に有効)
    $AUOptions = $obj.AUOptions

    echo @"
<h3>自動更新の構成</h3>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>自動更新の構成</th> <td>$AUOptions</td> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"</td> <td>AUOptions</td> <td>
    1 = 自動更新無効【既定】<br>
    2 = ダウンロードとインストールを通知<br>
    3 = 自動ダウンロードしインストールを通知<br>
    4 = 自動ダウンロードしインストール日時を指定<br>
    5 = ローカル管理者の設定選択を許可
  </td></tr>
</table>

"@

    if ( $AUOptions -eq 4 ) {
        # 自動ダウンロードしインストール日時を指定
        # (WIN_UPDATE_OPTIONに"4"を指定した場合に有効)

        # 自動メンテナンス時にインストールする
        $AutomaticMaintenanceEnabled = $obj.AutomaticMaintenanceEnabled

        # インストールを実行する日
        $ScheduledInstallDay = $obj.ScheduledInstallDay

        # インストールを実行する時間
        $ScheduledInstallTime = $obj.ScheduledInstallTime

        echo @"
<h3>自動ダウンロードしインストール日時を指定</h3>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>自動メンテナンス時にインストールする</th> <td>$AutomaticMaintenanceEnabled</td> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"</td> <td>AutomaticMaintenanceEnabled</td> <td>0 = 無効【既定】<br>1 = 有効</td></tr>
  <tr><th>インストールを実行する日</th> <td>$AutomaticMaintenanceEnabled</td> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"</td> <td>AutomaticMaintenanceEnabled</td> <td>
    0 = 毎日<br>
    1 = 毎週日曜<br>
    2 = 毎週月曜<br>
    3 = 毎週火曜<br>
    4 = 毎週水曜<br>
    5 = 毎週木曜<br>
    6 = 毎週金曜<br>
    7 = 毎週土曜
  </td></tr>
  <tr><th>インストールを実行する時間</th> <td>$ScheduledInstallTime</td> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"</td> <td>ScheduledInstallTime</td> <td>0 ～ 23 = 時刻を10進数で指定</td></tr>
</table>

"@
    }

    # 他のMicrosoft製品の更新プログラムのインストール
    $AllowMUUpdateService = $obj.AllowMUUpdateService

    echo @"
<h3>他のMicrosoft製品の更新プログラムのインストール</h3>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>他のMicrosoft製品の更新プログラムのインストール</th> <td>$AllowMUUpdateService</td> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"</td> <td>AllowMUUpdateService</td> <td>null or 0 = 無効【既定】<br>1 = 有効</td></tr>
</table>

"@

    # カスタム自動更新の検出頻度
    $DetectionFrequencyEnabled = $obj.DetectionFrequencyEnabled

    # 更新を確認する時間
    $DetectionFrequency = $obj.DetectionFrequency

    echo @"
<h3>自動更新の検出頻度</h3>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>カスタム検出時間の有効化</th> <td>$DetectionFrequencyEnabled</td> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"</td> <td>DetectionFrequencyEnabled</td> <td>0 = 無効<br>1 = 有効<br>2 = 未構成【既定】</td></tr>
  <tr><th>カスタム検出時間</th> <td>$DetectionFrequency</td> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"</td> <td>DetectionFrequency</td> <td></td></tr>
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_windowsupdate_wuserver
#       更新を検出するためのイントラネットの更新サービスを設定する
#-------------------------------------------------------------------------------
function fnc_windowsupdate_wuserver {

    # イントラネットのMicrosoft更新サービスの場所を指定する
    $obj = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
    $UseWUServer = $obj.UseWUServer

    echo @"
<h2><a name="WSUS">WSUS</a></h2>
<h3>イントラネットのMicrosoft更新サービスの場所を指定する</h3>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>イントラネットのMicrosoft更新サービスの場所を指定する</th> <td>$UseWUServer</td> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"</td> <td>UseWUServer</td> <td>0 = 無効<br>1 = 有効<br>2 = 未構成【既定】</td></tr>
"@

    if ( $UseWUServer -eq 1 ) {
        $obj = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"

        # 更新を検出するイントラネットの更新サービス(サーバ名かIPを指定)
        $WUServer = $obj.WUServer

        # イントラネット統計サーバーの設定(サーバ名かIPを指定)
        $WUStatusServer = $obj.WUStatusServer

        # 代替ダウンロードサーバーの設定
        $UpdateServiceUrlAlternate = $obj.UpdateServiceUrlAlternate

        # 代替ダウンロードサーバーが設定されている場合は、メタデータにURLが示されていないファイルをダウンロードします
        # (WIN_UPDATE_UPDATESERVICEURLALTERNATEが指定されている場合に有効)
        $FillEmptyContentUrls = $obj.FillEmptyContentUrls

        echo @"
  <tr><th>更新を検出するイントラネットの更新サービス</th> <td>$WUServer</td> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"</td> <td>WUServer</td> <td></td></tr>
  <tr><th>イントラネット統計サーバーの設定</th> <td>$WUStatusServer</td> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"</td> <td>WUStatusServer</td> <td></td></tr>
  <tr><th>代替ダウンロードサーバーの設定</th> <td>$UpdateServiceUrlAlternate</td> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"</td> <td>UpdateServiceUrlAlternate</td> <td></td></tr>
  <tr><th>代替ダウンロードサーバーが設定されている場合は、メタデータにURLが示されていないファイルをダウンロードします</th> <td>$FillEmptyContentUrls</td> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"</td> <td>FillEmptyContentUrls</td> <td>0 = 無効【既定】<br>1 = 有効</td></tr>
"@
    }

    echo @"
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_taskscheduler
#       タスクスケジューラ
#-------------------------------------------------------------------------------
function fnc_taskscheduler {

    echo @"
<h2><a name="タスクスケジューラ">タスクスケジューラ</a></h2>
<table>
  <tr><th>確認コマンド</th> <td>SCHTASKS /Query /V</td></tr>
</table>
<textarea  rows="10" wrap="off" readonly>
"@
    SCHTASKS /Query /V
    echo @"
</textarea>

"@
}


#-------------------------------------------------------------------------------
#   fnc_iesecurity
#       IEセキュリティ強化の構成
#-------------------------------------------------------------------------------
function fnc_iesecurity {
    $adm = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}").IsInstalled
    $usr = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}").IsInstalled

    echo @"
<h2><a name="IEセキュリティ強化の構成">IEセキュリティ強化の構成</a></h2>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>Administratorグループ</th> <td>$adm</td> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"</td> <td>IsInstalled</td> <td>0 = 無効<br>1 = 有効【既定】</td></tr>
  <tr><th>Usersグループ</th> <td>$usr</td> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"</td> <td>IsInstalled</td> <td>0 = 無効<br>1 = 有効【既定】</td></tr>
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_powerplan
#       電源プラン
#-------------------------------------------------------------------------------
function fnc_powerplan {
    $obj = Get-WmiObject -Namespace root\cimv2\power -Class win32_PowerPlan | Where-Object { $_.IsActive }
    $ElementName = $obj.ElementName

    echo @"
<h2><a name="電源プラン">電源プラン</a></h2>
<table>
  <tr><th>項目</th> <th>値</th> <th>コマンドレット</th> <th>プロパティ</th> <th>備考</th></tr>
  <tr><th>電源プラン</th> <td>$ElementName</td> <td>Get-WmiObject -Namespace root\cimv2\power -Class win32_PowerPlan</td> <td>ElementName</td> <td></td></tr>
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_netadapter
#       ネットワーク - 接続名称とデバイス情報
#-------------------------------------------------------------------------------
function fnc_netadapter {
    $obj = Get-NetAdapter | Sort-Object -Property Name

    echo @"
<h2><a name="接続名称とデバイス情報">接続名称とデバイス情報</a></h2>
<table>
  <tr><th>コマンドレット</th> <td>Get-NetAdapter</td> <th>プロパティ</th> <td>Name,Status,InterfaceDescription,MacAddress,LinkSpeed</td></tr>
</table>
<table>
  <tr><th>接続名称</th> <th>状態</th> <th>デバイス名</th> <th>MACアドレス</th> <th>SPEED＆DUPLEX</th> <th>備考</th></tr>
"@

    for($i = 0; $i -lt $obj.Count; $i++){
        Write-Output "  <tr><td>"$obj.Name[$i]"</td> <td>"$obj.Status[$i]"</td> <td>"$obj.InterfaceDescription[$i]"</td> <td>"$obj.MacAddress[$i]"</td> <td>"$obj.LinkSpeed[$i]"</td> <td></td></tr>"
    } 

    echo @"
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_netteam
#       ネットワーク - チーミング
#-------------------------------------------------------------------------------
function fnc_netteam {
    # ネットワーク - チーミング
    $obj1 = Get-NetLbfoTeam
    $obj2 = Get-NetLbfoTeamNic

    if ( $obj1.Name -eq $null ) {
        return
    } 

    echo @"
<h2><a name="チーミングデバイス情報">チーミングデバイス情報</a></h2>
<h3>チーミングアダプタ</h3>
<table>
  <tr><th rowspan=2>コマンドレット</th> <td>Get-NetLbfoTeam</td> <th rowspan=2>プロパティ</th> <td>Name,TeamingMode,LoadBalancingAlgorithm</td></tr>
  <tr>                                  <td>Get-NetLbfoTeamNic</td>                            <td>InterfaceDescription</td></tr>
</table>
<table>
  <tr><th>チーミング名</th> <th>チーミングモード</th> <th>負荷分散アルゴリズム</th> <th>デバイス名</th> <th>備考</th></tr>
"@

    for($i = 0; $i -lt $obj1.Count; $i++){
        Write-Output "  <tr><td>"$obj1.Name[$i]"</td> <td>"$obj1.TeamingMode[$i]"</td> <td>"$obj1.LoadBalancingAlgorithm[$i]"</td> <td>"$obj2.InterfaceDescription[$i]"</td> <td></td></tr>"
    } 

    echo @"
</table>

"@

    # ネットワーク - チーミングメンバ
    $obj = Get-NetLbfoTeamMember | Sort-Object -Property Team,Name

    echo @"
<h3>チーミングメンバー</h3>
<table>
  <tr><th>コマンドレット</th> <td>Get-NetLbfoTeamMember</td> <th>プロパティ</th> <td>Team,Name,AdministrativeMode,TransmitLinkSpeed,ReceiveLinkSpeed</td></tr>
</table>
<table>
  <tr><th>チーミング名</th> <th>アダプタ名</th> <th>Active/Standby</th> <th>Transmit Speed</th> <th>Recieve Speed</th> <th>備考</th></tr>
"@
    foreach($line in $obj){
        Write-Output "  <tr><td>"$line.Team"</td> <td>"$line.Name"</td> <td>"$line.AdministrativeMode"</td> <td>"$line.TransmitLinkSpeed"</td> <td>"$line.ReceiveLinkSpeed"</td> <td></td></tr>"
    }

    echo @"
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_netadapterbinding
#       ネットワーク - 接続プロパティ
#-------------------------------------------------------------------------------
function fnc_netadapterbinding {
    $obj = Get-NetAdapterBinding | Sort-Object -Property Name

    echo @"
<h2><a name="接続プロパティ">接続プロパティ</a></h2>
<table>
  <tr><th>コマンドレット</th> <td>Get-NetAdapterBinding</td> <th>プロパティ</th> <td>Name,Description,ComponentID,Enabled</td></tr>
</table>
<div style="height:300px; width=100%; overflow-y:scroll;">
<table>
  <tr><th>接続名称</th> <th>説明</th> <th>ID</th> <th>Enabled</th> <th>備考</th></tr>
"@

    foreach($line in $obj){
        Write-Output "  <tr><td>"$line.Name"</td> <td>"$line.Description"</td> <td>"$line.ComponentID"</td> <td>"$line.Enabled"</td> <td></td></tr>"
    }

    echo @"
</table>
</div>

"@
}


#-------------------------------------------------------------------------------
#   fnc_netipaddress
#       ネットワーク - IPアドレス
#-------------------------------------------------------------------------------
function fnc_netipaddress {
    $obj = Get-NetIPConfiguration

    echo @"
<h2><a name="Internet Protocol Version 4 (TCP/IPv4)のプロパティ">Internet Protocol Version 4 (TCP/IPv4)のプロパティ</a></h2>
<table>
  <tr><th rowspan="2">コマンドレット</th> <td>Get-NetIPConfiguration</td> <th rowspan="2">プロパティ</th> <td>InterfaceAlias,IPv4Address,IPv4DefaultGateway</td></tr>
  <tr>                                    <td>Get-WmiObject Win32_NetworkAdapterConfiguration</td> <td>IPSubnet</td></tr>
</table>
<table>
  <tr><th>接続名称</th> <th>IPアドレス</th> <th>サブネットマスク</th> <th>ゲートウェイ</th> <th>備考</th></tr>
"@

    foreach($line in $obj){
        $obj2 = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.IPAddress -eq $line.IPv4Address}
        $buf = $obj2.IPSubnet -split " "
        Write-Output "  <tr><td>"$line.InterfaceAlias"</td> <td>"$line.IPv4Address.IPAddress"</td> <td>"$buf[0]"</td> <td>"$line.IPv4DefaultGateway.NextHop"</td> <td></td></tr>"
    }

    echo @"
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_netdns
#       ネットワーク - DNS
#-------------------------------------------------------------------------------
function fnc_netdns {
    $obj = Get-DnsClientServerAddress | Where-Object { ($_.ServerAddresses -ne $null) -and ($_.AddressFamily -eq 2) }
    $obj = $obj.ServerAddresses -join ","

    echo @"
<h2><a name="DNSサーバーアドレス">DNSサーバーアドレス</a></h2>
<table>
  <tr><th>コマンドレット</th> <td>Get-DnsClientServerAddress</td> <th>プロパティ</th> <td>ServerAddresses</td></tr>
</table>
<table>
  <tr><th>DNSサーバーアドレス（使用順）</th> <th>備考</th></tr>
  <tr><td>$obj</td> <td></td></tr>
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_nethosts
#       ネットワーク - hosts設定
#-------------------------------------------------------------------------------
function fnc_nethosts {
    echo @"
<h2><a name="hosts">hosts</a></h2>
<table>
  <tr><th>ファイル</th> <td>$env:windir\System32\Drivers\etc\hosts</td></tr>
</table>
<textarea  rows="10" wrap="off" readonly>
"@
    Get-Content $env:windir\System32\Drivers\etc\hosts
    echo @"
</textarea>

"@
}


#-------------------------------------------------------------------------------
#   fnc_netroute
#       ネットワーク - Static Route設定
#-------------------------------------------------------------------------------
function fnc_netroute {
    echo @"
<h2><a name="静的経路">静的経路</a></h2>
<table>
  <tr><th>コマンドレット</th> <td>Get-NetRoute</td> <th>プロパティ</th> <td></td></tr>
</table>
<textarea  rows="10" wrap="off" readonly>
"@
    Get-NetRoute | Format-Table
    echo @"
</textarea>

"@
}


#-------------------------------------------------------------------------------
#   fnc_netsnp
#       ネットワーク - Scalable Networking Pack設定
#-------------------------------------------------------------------------------
function fnc_netsnp {
    $obj = netsh interface tcp show global

    echo @"
<h2><a name="Scalable Networking Pack設定">Scalable Networking Pack設定</a></h2>
<table>
  <tr><th>確認コマンド</th> <td>netsh interface tcp show global</td></tr>
</table>
<table>
  <tr><th>項目</th> <th>値</th> <th>備考</th></tr>
"@

    for ( $i = 4; $i -lt ($obj.Count - 1); $i++ ){
        $buf = $obj[$i] -split ":"
        Write-Output "  <tr><td>"$buf[0]"</td> <td>"$buf[1]"</td> <td></td></tr>"
    }

    echo @"
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_netsmb
#       ネットワーク - SMB Multichannel
#-------------------------------------------------------------------------------
function fnc_netsmb {
    $svr = $(Get-SmbServerConfiguration).EnableMultiChannel
    $cli = $(Get-SmbClientConfiguration).EnableMultiChannel

    echo @"
<h2><a name="SMBマルチチャネル機能">SMBマルチチャネル機能</a></h2>
<table>
  <tr><th rowspan="2">コマンドレット</th> <td>Get-SmbServerConfiguration</td> <th rowspan="2">プロパティ</th> <td>EnableMultiChannel</td></tr>
  <tr>                                    <td>Get-SmbClientConfiguration</td>                                 <td>EnableMultiChannel</td></tr>
</table>
<table>
  <tr><th>項目</th>   <th>SMBマルチチャネル機能有効</th> <th>備考</th></tr>
  <tr><td>Server</td> <td>$svr</td> <td></td></tr>
  <tr><td>Client</td> <td>$cli</td> <td></td></tr>
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_netfirewall
#       ネットワーク - ファイアウォール
#-------------------------------------------------------------------------------
function fnc_netfirewall {
    $obj = Get-NetFirewallProfile

    echo @"
<h2><a name="プロファイル">プロファイル</a></h2>
<table>
  <tr><th>コマンドレット</th> <td>Get-NetFirewallProfile</td> <th>プロパティ</th> <td>Name,Enabled</td></tr>
</table>
<table>
  <tr><th>プロファイル名</th> <th>Enabled</th> <th>備考</th></tr>
"@

    foreach($line in $obj){
        Write-Output "  <tr><td>"$line.Name"</td> <td>"$line.Enabled"</td> <td></td></tr>"
    }

    echo @"
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_netfirewallrule
#       ネットワーク - ファイアウォールルール
#-------------------------------------------------------------------------------
function fnc_netfirewallrule {
    $obj = Get-NetFirewallRule | Sort-Object -Property DisplayName

    echo @"
<h2><a name="ファイアウォールルール">ファイアウォールルール</a></h2>
<table>
  <tr><th>コマンドレット</th> <td>Get-NetFirewallRule</td> <th>プロパティ</th> <td>DisplayName,Profile,Direction,Action</td></tr>
</table>
<div style="height:300px; width=100%; overflow-y:scroll;">
<table>
  <tr><th>名前</th> <th>プロファイル</th> <th>送信/受信</th> <th>許可</th> <th>備考</th></tr>
"@

    foreach($line in $obj){
        Write-Output "  <tr><td>"$line.DisplayName"</td> <td>"$line.Profile"</td> <td>"$line.Direction"</td> <td>"$line.Action"</td> <td></td></tr>"
    }

    echo @"
</table>
</div>

"@
}


#-------------------------------------------------------------------------------
#   fnc_diskdevice
#       ディスク - ディスクデバイス
#-------------------------------------------------------------------------------
function fnc_diskdevice {
    $obj = Get-WmiObject Win32_DiskDrive | Sort-Object -Property Partitions

    echo @"
<h2><a name="ディスクデバイス">ディスクデバイス</a></h2>
<table>
  <tr><th>コマンドレット</th> <td>Get-WmiObject Win32_DiskDrive</td> <th>プロパティ</th> <td>Partitions,DeviceID,Caption,InterfaceType,Size</td></tr>
</table>
<table>
  <tr><th>Partition</th> <th>DeviceID</th> <th>Caption</th> <th>InterfaceType</th> <th>Size</th> <th>備考</th></tr>
"@

    foreach ($line in $obj) {
        Write-Output "  <tr><td>"$line.Partitions"</td> <td>"$line.DeviceID"</td> <td>"$line.Caption"</td> <td>"$line.InterfaceType"</td> <td>"$line.Size"</td> <td></td></tr>"
    }

    echo @"
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_diskvolume
#       ディスク - ボリューム/パーティション情報
#-------------------------------------------------------------------------------
function fnc_diskvolume {
    $obj = Get-Volume

    echo @"
<h2><a name="ボリューム/パーティション">ボリューム/パーティション</a></h2>
<table>
  <tr><th>コマンドレット</th> <td>Get-Volume</td> <th>プロパティ</th> <td>DriveLetter,FileSystemLabel,DriveType,FileSystemType,Size</td></tr>
</table>
<table>
  <tr><th>DriveLetter</th> <th>FileSystemLabel</th> <th>DriveType</th> <th>FileSystemType</th> <th>Size</th> <th>備考</th></tr>
"@

    foreach ($line in $obj) {
        Write-Output "  <tr><td>"$line.DriveLetter"</td> <td>"$line.FileSystemLabel"</td> <td>"$line.DriveType"</td> <td>"$line.FileSystemType"</td> <td>"$line.Size"</td> <td></td></tr>"
    }

    echo @"
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_group
#       アカウント - グループ
#-------------------------------------------------------------------------------
function fnc_group {
    $obj = Get-WmiObject Win32_Group

    echo @"
<h2><a name="グループ">グループ</a></h2>
<table>
  <tr><th>コマンドレット</th> <td>Get-WmiObject Win32_Group</td> <th>プロパティ</th> <td>Domain,Name</td></tr>
</table>
<table>
  <tr><th>ドメイン</th> <th>グループ名</th> <th>備考</th></tr>
"@

    foreach ($line in $obj) {
        Write-Output "  <tr><td>"$line.Domain"</td> <td>"$line.Name"</td> <td></td></tr>"
    }

    echo @"
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_user
#       アカウント - ユーザー
#-------------------------------------------------------------------------------
function fnc_user {
    $obj = Get-WmiObject Win32_UserAccount

    echo @"
<h2><a name="ユーザー">ユーザー</a></h2>
<table>
  <tr><th>コマンドレット</th> <td>Get-WmiObject Win32_UserAccount</td> <th>プロパティ</th> <td>Domain,Name,FullName,PasswordChangeable,PasswordRequired,Disabled</td></tr>
</table>
<table>
  <tr><th>ドメイン</th> <th>ユーザー名</th> <th>フルネーム</th> <th>ユーザーはパスワードを変更可能(True)/不可(False)</th> <th>ユーザーはパスワードが必要(True)/不要(False)</th> <th>ユーザーアカウントの無効(True)/有効(False)</th> <th>備考</th></tr>
"@

    foreach ($line in $obj) {
        Write-Output "  <tr><td>"$line.Domain"</td> <td>"$line.Name"</td> <td>"$line.FullName"</td> <td>"$line.PasswordChangeable"</td> <td>"$line.PasswordRequired"</td> <td>"$line.Disabled"</td> <td></td></tr>"
    }

    echo @"
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_accpolicy
#       グループポリシー - パスワード/ロックアウトポリシー
#-------------------------------------------------------------------------------
function fnc_accpolicy {
    $obj = net accounts

    echo @"
<h2><a name="パスワード/ロックアウトポリシー">パスワード/ロックアウトポリシー</a></h2>
<table>
  <tr><th>設定画面表示コマンド</th> <td>gpedit.msc</td></tr>
  <tr><th>確認コマンド</th> <td>net accounts</td></tr>
</table>
<table>
  <tr><th>項目</th> <th>値</th> <th>備考</th></tr>
"@

    for ( $i = 0; $i -lt ($obj.Count - 2); $i++ ){
        $buf = $obj[$i] -split ":"
        Write-Output "  <tr><td>"$buf[0]"</td> <td>"$buf[1]"</td> <td></td></tr>"
    }

    echo @"
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_uac
#       アカウント - ユーザーアカウント制御(UAC)
#-------------------------------------------------------------------------------
function fnc_uac {
    $obj = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
    $buf1=$obj.EnableLUA
    $buf2=$obj.ConsentPromptBehaviorAdmin
    $buf3=$obj.PromptOnSecureDesktop
    $buf4 = $(Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\luafv").Start

    echo @"
<h2><a name="ユーザーアカウント制御(UAC)">ユーザーアカウント制御(UAC)</a></h2>
<table>
  <tr><th>設定画面表示コマンド</th> <td>UserAccountControlSettings.exe</td></tr>
</table>
<h3>UAC無効化</h3>
<table>
  <tr><th>コマンドレット</th> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"</td> <th>プロパティ</th> <td>EnableLUA</td></tr>
</table>
<table>
  <tr><th>項目</th> <th>EnableLUA</th> <th>備考</th></tr>
  <tr><th>UAC無効化</th> <td>$buf1</td> <td>0 = UAC無効<br>null or 1 = UAC有効【既定】</td></tr>
</table>

"@

    if ( $obj -eq 0 ) {
        # UAC無効ならこの先は実行しない
        return
    }

    echo @"
<h3>コンピュータに対する変更の通知を受け取るタイミングの選択</h3>
<table>
  <tr><th>コマンドレット</th> <td>Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"</td> <th>プロパティ</th> <td>ConsentPromptBehaviorAdmin,PromptOnSecureDesktop</td></tr>
</table>
<table>
  <tr><th>項目</th> <th>ConsentPromptBehaviorAdmin</th> <th>PromptOnSecureDesktop</th> <th>備考</th></tr>
  <tr><th>コンピュータに対する変更の通知を受け取るタイミングの選択</th> <td>$buf2</td> <td>$buf3</td>
    <td>
    ConsentPromptBehaviorAdmin=2, PromptOnSecureDesktop=1 = 常に通知する<br>
    ConsentPromptBehaviorAdmin=5, PromptOnSecureDesktop=1 = アプリがコンピューターに変更を加えようとする場合のみ通知する(既定)<br>
    ConsentPromptBehaviorAdmin=5, PromptOnSecureDesktop=0 = アプリがコンピューターに変更を加えようとする場合のみ通知する(デスクトップを暗転しない)<br>
    ConsentPromptBehaviorAdmin=0, PromptOnSecureDesktop=0 = 通知しない
  </td></tr>
</table>

"@

    echo @"
<h3>UAC File Virtualization サービス</h3>
<table>
  <tr><th>コマンドレット</th> <td>Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\luafv"</td> <th>プロパティ</th> <td>Start</td></tr>
</table>
<table>
  <tr><th>項目</th> <th>Start</th> <th>備考</th></tr>
  <tr><th>UAC File Virtualization サービス</th> <td>$buf4</td> <td>2 = サービス有効<br>4 = サービス無効</td></tr>
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_installedsoftware
#       導入プログラム
#-------------------------------------------------------------------------------
function fnc_installedsoftware {
    $obj = Get-WmiObject Win32_Product | Sort-Object -Property Name

    echo @"
<h2><a name="導入プログラム">導入プログラム</a></h2>
<table>
  <tr><th>コマンドレット</th> <td>Get-WmiObject Win32_Product</td> <th>プロパティ</th> <td>Name,Vendor,Version</td></tr>
</table>
<table>
  <tr><th>名前</th> <th>ベンダー</th> <th>バージョン</th> <th>備考</th></tr>
"@

    foreach ($line in $obj) {
        Write-Output "  <tr><td>"$line.Name"</td> <td>"$line.Vendor"</td> <td>"$line.Version"</td> <td></td></tr>"
    }

    echo @"
</table>

"@
}


#-------------------------------------------------------------------------------
#   fnc_fixlist
#       適用プログラム - 適用パッチ
#-------------------------------------------------------------------------------
function fnc_fixlist {
    $obj = Get-HotFix | Sort-Object -Property HotFixID

    echo @"
<h2><a name="インストールされた更新プログラム">インストールされた更新プログラム</a></h2>
<table>
  <tr><th>コマンドレット</th> <td>Get-HotFix</td> <th>プロパティ</th> <td>Description,HotFixID,InstalledOn</td></tr>
</table>
"@

    # HotFixIDが10以上になったらtableにスクロールバーをつける
    if ( $obj.count -gt 10 ) {
        echo @"
<div style="height:300px; width=100%; overflow-y:scroll;">
"@
    }

    echo @"
<table>
  <tr><th>Description</th> <th>HotFixID</th> <th>InstalledOn</th> <th>備考</th></tr>
"@

    foreach ($line in $obj) {
        Write-Output "  <tr><td>"$line.Description"</td> <td>"$line.HotFixID"</td> <td>"$line.InstalledOn"</td> <td></td></tr>"
    }

    if ( $obj.count -gt 10 ) {
        echo @"
</table>
</div>

"@
    } else {
        echo @"
</table>

"@
    }
}


#-------------------------------------------------------------------------------
#   fnc_uacfilevirt
#       役割と機能
#-------------------------------------------------------------------------------
function fnc_rolefunction {
    $obj = Get-WindowsFeature

    echo @"
<h2><a name="役割と機能">役割と機能</a></h2>
<table>
  <tr><th>コマンドレット</th> <td>Get-WindowsFeature</td> <th>プロパティ</th> <td>DisplayName,Name,InstallState</td></tr>
</table>
<div style="height:300px; width=100%; overflow-y:scroll;">
<table>
  <tr><th>表示名</th> <th>名前</th> <th>導入状態</th> <th>備考</th></tr>
"@

    foreach ($line in $obj) {
        Write-Output "  <tr><td>"$line.DisplayName"</td> <td>"$line.Name"</td> <td>"$line.InstallState"</td> <td></td></tr>"
    }

    echo @"
</table>
</div>

"@
}


#-------------------------------------------------------------------------------
#   fnc_service
#       サービス
#-------------------------------------------------------------------------------
function fnc_service {
    $obj = Get-Service

    echo @"
<h2><a name="サービス">サービス</a></h2>
<table>
  <tr><th>コマンドレット</th> <td>Get-Service</td> <th>プロパティ</th> <td>DisplayName,Name,StartType,Status</td></tr>
</table>
<div style="height:300px; width=100%; overflow-y:scroll;">
<table>
  <tr><th>表示名</th> <th>名前</th> <th>状態</th> <th>スタートアップの種類<th>備考</th></tr>
"@

    foreach ($line in $obj) {
        Write-Output "  <tr><td>"$line.DisplayName"</td> <td>"$line.Name"</td> <td>"$line.Status"</td> <td>"$line.StartType"</td> <td></td></tr>"
    }

    echo @"
</table>
</div>

"@
}


#-------------------------------------------------------------------------------
#   main
#-------------------------------------------------------------------------------
$hostname=hostname

# HTML header and start body
fnc_htmlheader
fnc_index
echo "<h1>$hostname Parameter Sheet</h1>"

echo "<h1>systeminfo</h1>"
fnc_systeminfo

echo "<h1>システムのプロパティ</h1>"
echo "<table><tr><th>設定画面表示コマンド</th> <td>control /name Microsoft.System</td></tr></table>"
fnc_operatingsystem
fnc_workgroup
fnc_dnssuffix
fnc_hardware
fnc_visualeffect
fnc_prioritycontrol
fnc_recoveros
fnc_env
fnc_remote_desktop

echo "<h1>時計と地域</h1>"
echo "<table><tr><th>設定画面表示コマンド</th> <td>control /name Microsoft.DateAndTime</td></tr></table>"
fnc_timezone

echo "<h1>時刻同期サーバーサービス</h1>"
echo "<table><tr><th>設定画面表示コマンド</th> <td>設定画面なし</td></tr></table>"
fnc_ntpserver

echo "<h1>イベントログ</h1>"
echo "<table><tr><th>設定画面表示コマンド</th> <td>eventvwr.msc</td></tr></table>"
fnc_systemlog

echo "<h1>Windows Update</h1>"
echo "<table><tr><th>設定画面表示コマンド</th> <td>gpedit.msc</td></tr></table>"
fnc_windowsupdate
fnc_windowsupdate_wuserver

echo "<h1>タスクスケジューラ</h1>"
echo "<table><tr><th>設定画面表示コマンド</th> <td>taskschd.msc</td></tr></table>"
fnc_taskscheduler

echo "<h1>IEセキュリティ強化の構成</h1>"
echo "<table><tr><th>設定画面表示コマンド</th> <td>ServerManager.exe</td></tr></table>"
fnc_iesecurity

echo "<h1>電源プラン</h1>"
echo "<table><tr><th>設定画面表示コマンド</th> <td>control /name Microsoft.PowerOptions</td></tr></table>"
fnc_powerplan

echo "<h1>ネットワークアダプタ</h1>"
echo "<table><tr><th>設定画面表示コマンド</th> <td>ServerManager.exe</td></tr></table>"
fnc_netadapter
fnc_netteam

echo "<h1>ネットワーク</h1>"
echo "<table><tr><th>設定画面表示コマンド</th> <td>control /name Microsoft.NetworkAndSharingCenter</td></tr></table>"
fnc_netadapterbinding
fnc_netipaddress
fnc_netdns
fnc_nethosts
fnc_netroute
fnc_netsnp
fnc_netsmb

echo "<h1>ファイヤーウォール</h1>"
echo "<table><tr><th>設定画面表示コマンド</th> <td>control /name Microsoft.WindowsFirewall</td></tr></table>"
fnc_netfirewall
fnc_netfirewallrule

echo "<h1>ディスク</h1>"
echo "<table><tr><th>設定画面表示コマンド</th> <td>diskmgmt.msc</td></tr></table>"
fnc_diskdevice
fnc_diskvolume

echo "<h1>アカウント</h1>"
echo "<table><tr><th>設定画面表示コマンド</th> <td>compmgmt.msc</td></tr></table>"
fnc_group
fnc_user
fnc_accpolicy
fnc_uac

echo "<h1>プログラムと機能</h1>"
echo "<table><tr><th>設定画面表示コマンド</th> <td>control /name Microsoft.ProgramsAndFeatures</td></tr></table>"
fnc_installedsoftware
fnc_fixlist
fnc_rolefunction

echo "<h1>サービス</h1>"
echo "<table><tr><th>設定画面表示コマンド</th> <td>services.msc</td></tr></table>"
fnc_service

# HTML close body
fnc_htmlfooter