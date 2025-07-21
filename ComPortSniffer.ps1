<################################################# COMポートスニファ Ver1.2 2021/02/27
●スクリプト説明
2つのCOMポートを使用してバケツリレーのように双方向にデータを中継することで、通信をモニターするスクリプトです。

バッチまたは、ショートカットでの起動方法
  powershell -WindowStyle Hidden -ExecutionPolicy RemoteSigned -command ".\スクリプトファイル名"  

●使い方
シリアル通信パラメータ設定を設定し、実行するとデータがコンソールに表示されます。
起動オプションでCOMポートのパラメータを指定可能。

例）COMポートスニファ.exe COM1 19200 None 8 One COM2 38400 None 8 One -byte -log

●変更履歴
V1.0  初版
V1.1  起動オプション設定を追加
V1.2  CTRL＋Cで入力でlog出力して終了を追加

#>
#---+---+---+---+---+---+---+---+---+---+ 起動オプション パラメータ取得
Param(
    [string]$Arg1 = "COM1",    #COM番号
    [string]$Arg2 = "38400",   #bps
    [string]$Arg3 = "None",    #パリティ
    [string]$Arg4 = "8",       #データ長
    [string]$Arg5 = "One",     #ストップビット

    [string]$Arg6 = "COM2",    #COM番号
    [string]$Arg7 = "38400",   #bps
    [string]$Arg8 = "None",    #パリティ
    [string]$Arg9 = "8",       #データ長
    [string]$ArgA = "One",     #ストップビット

    [Switch]$Byte = $false,    #バイト値表示する
    [Switch]$Log = $false      #log出力する
)

Set-StrictMode -Version Latest    #コーディング規則を設定して適用します。（デバッグに有用）

################################################## 初期設定
#---+---+---+---+---+---+---+---+---+---+ 変数宣言
# シリアル通信パラメータ設定
$ComParam1 = $Arg1, $Arg2, $Arg3, $Arg4, $Arg5    #1番目 COMポート設定 COM番号、ボーレート、データ長、パリティ、ストップ
$ComParam2 = $Arg6, $Arg7, $Arg8, $Arg9, $ArgA    #2番目 COMポート設定 COM番号、ボーレート、データ長、パリティ、ストップ

#---+---+---+---+---+---+---+---+---+---+ エラー発生時動作
$ErrorActionPreference = "stop"    #エラー発生時にスクリプトを中断します。

################################################## 関数定義
#---+---+---+---+---+---+---+---+---+---+ シリアル通信の制御線確認
Function CtsDsrChk {
    While (-not ($ComPortObj.CtsHolding -and $ComPortObj.DsrHolding)) {
        $Input = Read-Host "機器の電源および接続を確認してください。"
    }
}
#---+---+---+---+---+---+---+---+---+---+ 1番目 COM送信処理
# 送信
Function Send1($Str) {
        $ComPortObj1.Write($Str)
}

#---+---+---+---+---+---+---+---+---+---+ 2番目 COM送信処理
# 送信
Function Send2($Str) {
        $ComPortObj2.Write($Str)
}

#---+---+---+---+---+---+---+---+---+---+ 文字列 ⇒ バイト表示
Function Bytes($Str) {
        $Ary = [system.text.encoding]::Default.GetBytes("$Str")
        $Byte_Str = ""
        foreach ( $B in $Ary ){
            $Byte_Str += [System.Convert]::ToString($B, 16) + " "    #16進数変換
        }
        Write-Host " $Byte_Str"
}

################################################## イニシャル処理
#---+---+---+---+---+---+---+---+---+---+ 1番目 COMポート
# COMポート生成、パラメータ設定
$ComPortObj1 = New-Object System.IO.Ports.SerialPort $ComParam1    #COM番号、ボーレート、データ長、パリティ、ストップ
$ComPortObj1.DtrEnable = $true    #DTR設定
$ComPortObj1.RtsEnable = $true    #RTS設定
$ComPortObj1.Handshake = "None"   # ハンドシェイク設定 (None、XOnXOff、RequestToSend、RequestToSendXOnXOff)
$ComPortObj1.NewLine = "`r"       #改行文字設定（WriteLineやReadLineメソッドに適用）
$ComPortObj1.Encoding=[System.Text.Encoding]::GetEncoding("Shift_JIS")    # 文字コード設定
# シリアルポート受信イベント
$ComRecvEventObj1 = Register-ObjectEvent -InputObject $ComPortObj1 -EventName "DataReceived" `
    -Action {Param([System.IO.Ports.SerialPort]$sender, [System.EventArgs]$e) Write-Output $sender.ReadExisting()}
# シリアルポートエラーイベント（エラー発生時に$trueを返す）通信でパリティエラーが発生した場合など
$ComErrEventObj1 = Register-ObjectEvent -InputObject $ComPortObj1 -EventName "ErrorReceived" -Action {$true}
# シリアルポート非データ信号イベント（制御ピンが変化した場合に$trueを返す）通信側との接続に異常がある場合など？
#$ComPinEventObj = Register-ObjectEvent -InputObject $ComPortObj -EventName "PinChanged" -Action {$true}

#---+---+---+---+---+---+---+---+---+---+ 2番目 COMポート
# COMポート生成、パラメータ設定
$ComPortObj2 = New-Object System.IO.Ports.SerialPort $ComParam2    #COM番号、ボーレート、データ長、パリティ、ストップ
$ComPortObj2.DtrEnable = $true    #DTR設定
$ComPortObj2.RtsEnable = $true    #RTS設定
$ComPortObj2.Handshake = "None"   # ハンドシェイク設定 (None、XOnXOff、RequestToSend、RequestToSendXOnXOff)
$ComPortObj2.NewLine = "`r"       #改行文字設定（WriteLineやReadLineメソッドに適用）
$ComPortObj2.Encoding=[System.Text.Encoding]::GetEncoding("Shift_JIS")    # 文字コード設定
# シリアルポート受信イベント
$ComRecvEventObj2 = Register-ObjectEvent -InputObject $ComPortObj2 -EventName "DataReceived" `
    -Action {Param([System.IO.Ports.SerialPort]$sender, [System.EventArgs]$e) Write-Output $sender.ReadExisting()}
# シリアルポートエラーイベント（エラー発生時に$trueを返す）通信でパリティエラーが発生した場合など
$ComErrEventObj2 = Register-ObjectEvent -InputObject $ComPortObj2 -EventName "ErrorReceived" -Action {$true}
# シリアルポート非データ信号イベント（制御ピンが変化した場合に$trueを返す）通信側との接続に異常がある場合など？
#$ComPinEventObj = Register-ObjectEvent -InputObject $ComPortObj -EventName "PinChanged" -Action {$true}

#---+---+---+---+---+---+---+---+---+---+ COMポートオープン
$ErrorActionPreference = "SilentlyContinue"    #エラー発生時に無視します。
try {
    $ComPortObj1.Open()
} catch {
    $Input = Read-Host "$ComParam1 ポートのオープンが出来ません。"
    Exit    #スクリプト終了
}
try {
    $ComPortObj2.Open()
} catch {
    $Input = Read-Host "$ComParam2 ポートのオープンが出来ません。"
    $ComPortObj1.Close()
    Exit    #スクリプト終了
}
$ErrorActionPreference = "stop"    #エラー発生時にスクリプトを中断します。

#---+---+---+---+---+---+---+---+---+---+ logファイル出力
$now = Get-Date -Format 'yyMMdd-HHmmss'    #現在日時
$logname = "log_$now.txt"    #logファイル名
if($Log) { Start-Transcript $logname }    #log記録開始

################################################## メイン処理
Write-Host "CTRL+C を入力すると終了"
Try {
    While($true) {
        #1番目 COM受信確認
        $RecvStr1 = Receive-job -job $ComRecvEventObj1
        if($RecvStr1) {
            Write-Host ">$RecvStr1"
            if($Byte) { Bytes $RecvStr1 }
            Send2 $RecvStr1
        }

        #2番目 COM受信確認
        $RecvStr2 = Receive-job -job $ComRecvEventObj2
        if($RecvStr2) {
            Write-Host "<$RecvStr2"
            if($Byte) { Bytes $RecvStr2 }
            Send1 $RecvStr2
        }

        #通信エラー確認
        If (Receive-job -job $ComErrEventObj1) { Write-Host "通信エラーが発生しました。" }
        If (Receive-job -job $ComErrEventObj2) { Write-Host "通信エラーが発生しました。" }
        #If (Receive-job -job $ComPinEventObj1) { Write-Host "接続エラー" }
        #If (Receive-job -job $ComPinEventObj2) { Write-Host "接続エラー" }
    }
}
################################################## 終了処理
finally {
    #---+---+---+---+---+---+---+---+---+---+ COMポートクローズ
    $ComPortObj1.Close()
    $ComPortObj2.Close()

    if($Log) { Stop-Transcript }    #logファイル出力

    Exit    #スクリプト終了
}
################################################## End Of File

#シリアル通信のエラー(パリティエラーとか)については、ErrorReceivedイベント処理を追加すれば対応できる。
#シリアル通信の制御信号については、PinChangedイベント処理を追加すれば対応できる。

<#
フィールド 
InfiniteTimeout     タイムアウトが発生しないことを示します。

プロパティ 
BaudRate            シリアル ボー レートを取得または設定します。
BreakState          ブレーク シグナルの状態を取得または設定します。
BytesToRead         受信バッファー内のデータのバイト数を取得します。
BytesToWrite        送信バッファー内のデータのバイト数を取得します。
CDHolding           ポートのキャリア検出ラインの状態を取得します。
CtsHolding          Clear To Send ラインの状態を取得します。
DataBits            バイトごとのデータ ビットの標準の長さを取得または設定します。
DiscardNull         ポートと受信バッファー間での送信時に、null バイトを無視するかどうかを示す値を取得または設定します。
DsrHolding          DSR (Data Set Ready) シグナルの状態を取得します。
DtrEnable           シリアル通信中に、DTR (Data Terminal Ready) シグナルを有効にする値を取得または設定します。
Encoding            テキストの伝送前変換と伝送後変換のバイト エンコーディングを取得または設定します。
Handshake           Handshake からの値を使用したデータのシリアル ポート伝送のハンドシェイク プロトコルを取得または設定します。
IsOpen              SerialPort  オブジェクトの開いている状態または閉じた状態を示す値を取得します。
NewLine             ReadLine() メソッドと WriteLine(String) メソッドの呼び出しの末尾を解釈する際に使用する値を取得または設定します。
Parity             パリティ チェック プロトコルを取得または設定します。
ParityReplace      パリティ エラーの発生時に、データ ストリーム内の無効なバイトを置き換えるバイトを取得または設定します。
PortName           通信用のポートを取得または設定します。このポートには、使用可能なすべての COM ポートが含まれますが、これに限定されるわけではありません。
ReadBufferSize     SerialPort の入力バッファーのサイズを取得または設定します。
ReadTimeout        読み取り操作が完了していないときに、タイムアウトになるまでのミリ秒数を取得または設定します。
ReceivedBytesThreshold DataReceived イベントが発生する前の、内部入力バッファーのバイト数を取得または設定します。
RtsEnable          シリアル通信中に、RTS (Request To Send) シグナルが有効かどうかを示す値を取得または設定します。
StopBits           バイトごとのストップ ビットの標準の数を取得または設定します。
WriteBufferSize    シリアル ポートの出力バッファーのサイズを取得または設定します。
WriteTimeout       書き込み操作が完了していないときに、タイムアウトになるまでのミリ秒数を取得または設定します。

メソッド 
Close()            ポート接続を閉じ、IsOpen プロパティを false に設定し、内部 Stream オブジェクトを破棄します。
DiscardInBuffer()  シリアル ドライバーの受信バッファーからデータを破棄します。
DiscardOutBuffer() シリアル ドライバーの送信バッファーからデータを破棄します。
Dispose(Boolean)   SerialPort によって使用されているアンマネージド リソースを解放し、オプションでマネージド リソースも解放します。
GetPortNames()     現在のコンピューターのシリアル ポート名の配列を取得します。
Open()             新しいシリアル ポート接続を開きます。
Read(Byte[], Int32, Int32)    SerialPort の入力バッファーから複数のバイトを読み取り、読み取ったバイトを指定したオフセットでバイト配列に書き込みます。
Read(Char[], Int32, Int32)    SerialPort の入力バッファーから複数の文字を読み取り、読み取った文字を指定したオフセットで文字配列に書き込みます。
ReadByte()         SerialPort の入力バッファーから、同期で 1 バイトを読み取ります。
ReadChar()         SerialPort の入力バッファーから、同期で 1 文字を読み取ります。
ReadExisting()     ストリームと SerialPort オブジェクトの入力バッファーの両方で、エンコーディングに基づいて、即座に使用できるすべてのバイトを読み取ります。
ReadLine()         入力バッファー内の NewLine 値まで読み取ります。
ReadTo(String)     入力バッファー内の指定した value まで文字列を読み取ります。
Write(Byte[], Int32, Int32)    バッファーのデータを使用して、指定したバイト数をシリアル ポートに書き込みます。
Write(Char[], Int32, Int32)    バッファーのデータを使用して、指定した文字数をシリアル ポートに書き込みます。
Write(String)      指定した文字列をシリアル ポートに書き込みます。
WriteLine(String)  指定した文字列と NewLine 値を出力バッファーに書き込みます。

イベント
メインスレッドではなくセカンダリスレッドで発生するため、メインスレッドの一部の要素を変更しようとすると、
スレッド例外が発生する可能性があります。要素を変更する必要がある場合は、Invokeを使用して変更要求を戻します。
DataReceived       SerialPort オブジェクトによって表されるポートを介してデータが受信されたことを示します。
ErrorReceived      SerialPort オブジェクトによって表されるポートでエラーが発生したことを示します。
PinChanged         非データ信号イベントが SerialPort オブジェクトによって表されるポートで発生したことを示します。
#>