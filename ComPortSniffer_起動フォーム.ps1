<################################################# ComPortSniffer起動フォーム Ver1.1 2025/07/21
●スクリプト説明
・プログラムを起動する際に、オプションパラメータを簡単に選択して起動できるフォームです。

●使い方
・変更したい設定項目を入力すると起動オプションの欄に文字列が表示されます。
・実行ボタンを押すと表示された起動オプションのパラメータを使用して、プログラムを起動します。
・起動オプション欄のパラメータは、文字列をコピーしてバッチやショートカットのパラメータとして使用できます。
#>

################################################## 初期設定
#---+---+---+---+---+---+---+---+---+---+ コーディング規則を設定して適用します。（デバッグに有用）
Set-StrictMode -Version Latest    #コーディング規則を設定して適用します。（デバッグに有用）

#---+---+---+---+---+---+---+---+---+---+ 変数宣言
$WinTitle = "ComPortSniffer起動フォーム"  #ウィンドウタイトル
$ScrName = ".\ComPortSniffer.exe"    #起動するファイル

################################################## イニシャル処理
#---+---+---+---+---+---+---+---+---+---+ アセンブリのロード
Add-Type -AssemblyName System.Windows.Forms

#---+---+---+---+---+---+---+---+---+---+
Set-Location -Path $PSScriptRoot    #カレントディレクトリをスクリプトのディレクトリへ変更

#---+---+---+---+---+---+---+---+---+---+ タイマー（起動オプション テキスト表示）
$Timer = New-Object System.Windows.Forms.Timer
$Timer.Interval = 1000    #イベントを発生させる間隔(ms)
$Time = {
    $TextStr = ""
    #処理
    If ($Combo_Com.Text -ne "") { $TextStr += " " + $Combo_Com.Text + " " + $Combo_Baud.Text + " " + $Combo_Parity.Text + " " + $Combo_Bit.Text + " " + $Combo_Stop.Text }
    If ($Combo_Com2.Text -ne "") { $TextStr += " " + $Combo_Com2.Text + " " + $Combo_Baud2.Text + " " + $Combo_Parity2.Text + " " + $Combo_Bit2.Text + " " + $Combo_Stop2.Text }
    If ($CheckBox_Bin.Checked) { $TextStr += " -byte" }
    If ($CheckBox_Log.Checked) { $TextStr += " -log" }
    $TextBox_Arg.Text = $TextStr
}
$Timer.Add_Tick($Time)

#---+---+---+---+---+---+---+---+---+---+ メインフォーム
#フォーム作成
$Form = New-Object System.Windows.Forms.Form
$Form.Text = $WinTitle
$Form.Size = "410, 280"
$Form.StartPosition = "CenterScreen"
$Form.MaximizeBox = $false    #最大化ボタン
$Form.MinimizeBox = $false    #最小化ボタン

#---+---+---+---+---+---+---+---+---+---+ COMポート1設定
$X=10; $Y=10
#ラベル
$Label_Serial = New-Object System.Windows.Forms.Label
$Label_Serial.Location = "$X, $Y"
$Label_Serial.AutoSize = $true   #文字の長さに合わせ自動調整
$Label_Serial.Text = "シリアル通信設定 #1"
$Form.Controls.Add($Label_Serial)

# コンボボックス COM番号
$Combo_Com = New-Object System.Windows.Forms.Combobox
$Combo_Com.Location = "$X, $($Y + 20)"
$Combo_Com.size = "60, 30"
$Combo_Com.DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
<#
# コンボボックスに項目を追加
[void] $Combo_Com.Items.Add("COM1")
[void] $Combo_Com.Items.Add("COM2")
[void] $Combo_Com.Items.Add("COM3")
#$Combo_Com.SelectedIndex = 0
#>
$Form.Controls.Add($Combo_Com)
#ドロップダウン時イベント
$Combo_Com.Add_DropDown({
    $Combo_Com.Items.Clear()
    $ComAry = [System.IO.Ports.SerialPort]::GetPortNames()  #COMポート一覧取得
    If ($ComAry.Count -gt 0) {
        [void] $Combo_Com.Items.AddRange($ComAry)  #アイテム追加
    }
})
#結果取得
#$Result = $Combo_Com.Text    #コンボボックス選択項目取得

# コンボボックス ボーレート
$Combo_Baud = New-Object System.Windows.Forms.Combobox
$Combo_Baud.Location = "$($X + 70), $($Y + 20)"
$Combo_Baud.size = "60, 30"
$Combo_Baud.DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
# コンボボックスに項目を追加
[void] $Combo_Baud.Items.Add("9600")
[void] $Combo_Baud.Items.Add("19200")
[void] $Combo_Baud.Items.Add("38400")
[void] $Combo_Baud.Items.Add("57600")
[void] $Combo_Baud.Items.Add("115200")
$Combo_Baud.SelectedIndex = 2
$Form.Controls.Add($Combo_Baud) 
#結果取得
#$Result = $Combo_Baud.Text    #コンボボックス選択項目取得

# コンボボックス パリティ
$Combo_Parity = New-Object System.Windows.Forms.Combobox
$Combo_Parity.Location = "$($X + 140), $($Y + 20)"
$Combo_Parity.size = "60, 30"
$Combo_Parity.DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
# コンボボックスに項目を追加
[void] $Combo_Parity.Items.Add("None")
[void] $Combo_Parity.Items.Add("Odd")
[void] $Combo_Parity.Items.Add("Even")
[void] $Combo_Parity.Items.Add("Mark")
[void] $Combo_Parity.Items.Add("Space")
$Combo_Parity.SelectedIndex = 0
$Form.Controls.Add($Combo_Parity) 
#結果取得
#$Result = $Combo_Parity.Text    #コンボボックス選択項目取得

# コンボボックス データ長
$Combo_Bit = New-Object System.Windows.Forms.Combobox
$Combo_Bit.Location = "$($X + 210), $($Y + 20)"
$Combo_Bit.size = "60, 30"
$Combo_Bit.DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
# コンボボックスに項目を追加
[void] $Combo_Bit.Items.Add("8")
[void] $Combo_Bit.Items.Add("7")
[void] $Combo_Bit.Items.Add("6")
[void] $Combo_Bit.Items.Add("5")
$Combo_Bit.SelectedIndex = 0
$Form.Controls.Add($Combo_Bit) 
#結果取得
#$Result = $Combo_Bit.Text    #コンボボックス選択項目取得

# コンボボックス ストップビット
$Combo_Stop = New-Object System.Windows.Forms.Combobox
$Combo_Stop.Location = "$($X + 280), $($Y + 20)"
$Combo_Stop.size = "90, 30"
$Combo_Stop.DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
# コンボボックスに項目を追加
[void] $Combo_Stop.Items.Add("One")
[void] $Combo_Stop.Items.Add("OnePointFive")
[void] $Combo_Stop.Items.Add("Two")
$Combo_Stop.SelectedIndex = 0
$Form.Controls.Add($Combo_Stop) 
#結果取得
#$Result = $Combo_Stop.Text    #コンボボックス選択項目取得

#---+---+---+---+---+---+---+---+---+---+ COMポート2設定
$X=10; $Y=60
#ラベル
$Label_Serial2 = New-Object System.Windows.Forms.Label
$Label_Serial2.Location = "$X, $Y"
$Label_Serial2.AutoSize = $true            #文字の長さに合わせ自動調整
$Label_Serial2.Text = "シリアル通信設定 #2"
$Form.Controls.Add($Label_Serial2)

# コンボボックス COM番号
$Combo_Com2 = New-Object System.Windows.Forms.Combobox
$Combo_Com2.Location = "$X, $($Y + 20)"
$Combo_Com2.size = "60, 30"
$Combo_Com2.DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
<#
# コンボボックスに項目を追加
[void] $Combo_Com2.Items.Add("COM1")
[void] $Combo_Com2.Items.Add("COM2")
[void] $Combo_Com2.Items.Add("COM3")
#$Combo_Com2.SelectedIndex = 0
#>
$Form.Controls.Add($Combo_Com2) 
#ドロップダウン時イベント
$Combo_Com2.Add_DropDown({
    $Combo_Com2.Items.Clear()
    $ComAry = [System.IO.Ports.SerialPort]::GetPortNames()  #COMポート一覧取得
    If ($ComAry.Count -gt 0) {
        [void] $Combo_Com2.Items.AddRange($ComAry)  #アイテム追加
    }
})
#結果取得
#$Result = $Combo_Com2.Text    #コンボボックス選択項目取得

# コンボボックス ボーレート
$Combo_Baud2 = New-Object System.Windows.Forms.Combobox
$Combo_Baud2.Location = "$($X + 70), $($Y + 20)"
$Combo_Baud2.size = "60, 30"
$Combo_Baud2.DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
# コンボボックスに項目を追加
[void] $Combo_Baud2.Items.Add("9600")
[void] $Combo_Baud2.Items.Add("19200")
[void] $Combo_Baud2.Items.Add("38400")
[void] $Combo_Baud2.Items.Add("57600")
[void] $Combo_Baud2.Items.Add("115200")
$Combo_Baud2.SelectedIndex = 2
$Form.Controls.Add($Combo_Baud2) 
#結果取得
#$Result = $Combo_Baud2.Text    #コンボボックス選択項目取得

# コンボボックス パリティ
$Combo_Parity2 = New-Object System.Windows.Forms.Combobox
$Combo_Parity2.Location = "$($X + 140), $($Y + 20)"
$Combo_Parity2.size = "60, 30"
$Combo_Parity2.DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
# コンボボックスに項目を追加
[void] $Combo_Parity2.Items.Add("None")
[void] $Combo_Parity2.Items.Add("Odd")
[void] $Combo_Parity2.Items.Add("Even")
[void] $Combo_Parity2.Items.Add("Mark")
[void] $Combo_Parity2.Items.Add("Space")
$Combo_Parity2.SelectedIndex = 0
$Form.Controls.Add($Combo_Parity2) 
#結果取得
#$Result = $Combo_Parity2.Text    #コンボボックス選択項目取得

# コンボボックス データ長
$Combo_Bit2 = New-Object System.Windows.Forms.Combobox
$Combo_Bit2.Location = "$($X + 210), $($Y + 20)"
$Combo_Bit2.size = "60, 30"
$Combo_Bit2.DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
# コンボボックスに項目を追加
[void] $Combo_Bit2.Items.Add("8")
[void] $Combo_Bit2.Items.Add("7")
[void] $Combo_Bit2.Items.Add("6")
[void] $Combo_Bit2.Items.Add("5")
$Combo_Bit2.SelectedIndex = 0
$Form.Controls.Add($Combo_Bit2) 
#結果取得
#$Result = $Combo_Bit2.Text    #コンボボックス選択項目取得

# コンボボックス ストップビット
$Combo_Stop2 = New-Object System.Windows.Forms.Combobox
$Combo_Stop2.Location = "$($X + 280), $($Y + 20)"
$Combo_Stop2.size = "90, 30"
$Combo_Stop2.DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
# コンボボックスに項目を追加
[void] $Combo_Stop2.Items.Add("One")
[void] $Combo_Stop2.Items.Add("OnePointFive")
[void] $Combo_Stop2.Items.Add("Two")
$Combo_Stop2.SelectedIndex = 0
$Form.Controls.Add($Combo_Stop2) 
#結果取得
#$Result = $Combo_Stop2.Text    #コンボボックス選択項目取得

#---+---+---+---+---+---+---+---+---+---+ バイナリ表示
$X=10; $Y=110  #表示位置
# チェックボックス
$CheckBox_Bin = New-Object System.Windows.Forms.CheckBox
$CheckBox_Bin.Location = "$X, $Y"
$CheckBox_Bin.Size = "100, 30"
$CheckBox_Bin.Text = "バイナリ表示"
$Form.Controls.Add($CheckBox_Bin)
#結果取得
#$Result = $CheckBox_Bin.Checked    #チェックボックス取得 $true $false

#---+---+---+---+---+---+---+---+---+---+ ログ出力
$X=120; $Y=110  #表示位置
# チェックボックス
$CheckBox_Log = New-Object System.Windows.Forms.CheckBox
$CheckBox_Log.Location = "$X, $Y"
$CheckBox_Log.Size = "100, 30"
$CheckBox_Log.Text = "ログ出力"
$Form.Controls.Add($CheckBox_Log)
#結果取得
#$Result = $CheckBox_Log.Checked    #チェックボックス取得 $true $false

#---+---+---+---+---+---+---+---+---+---+
$X=10; $Y=150
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$X, $Y"
$Label.AutoSize = $True            #文字の長さに合わせ自動調整
$Label.Text = '起動オプション'
$Form.Controls.Add($Label)

#テキストボックス 起動オプション
$TextBox_Arg = New-Object System.Windows.Forms.TextBox
$TextBox_Arg.Location = "$X, $($Y + 20)"
#$TextBox_Arg.MaxLength = 128       #最大入力文字数
$TextBox_Arg.ReadOnly = $true      #編集不可
$TextBox_Arg.Anchor = (([System.Windows.Forms.AnchorStyles]::Left)`
              -bor ([System.Windows.Forms.AnchorStyles]::Top)`
              -bor ([System.Windows.Forms.AnchorStyles]::Right)`
              -bor ([System.Windows.Forms.AnchorStyles]::Bottom))    #位置固定(画面サイズ変更時など)
$TextBox_Arg.Size = "365, 10"
$Form.Controls.Add($TextBox_Arg)

#ボタン OK
$OkButton = New-Object System.Windows.Forms.Button
$OkButton.Location = "$($X + 280), $($Y + 50)"
$OkButton.Size = "70, 23"
$OkButton.Text = '実行'
$OkButton.Anchor = (([System.Windows.Forms.AnchorStyles]::Right)`
               -bor ([System.Windows.Forms.AnchorStyles]::Bottom))    #位置固定(画面サイズ変更時など)
$OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$Form.AcceptButton = $OkButton
$Form.Controls.Add($OkButton)

#ボタン キャンセル
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = "$($X + 470), $($Y + 50)"
$CancelButton.Size = "75, 23"
$CancelButton.Text = 'Cancel'
$CancelButton.Anchor = (([System.Windows.Forms.AnchorStyles]::Right)`
                   -bor ([System.Windows.Forms.AnchorStyles]::Bottom))    #位置固定(画面サイズ変更時など)
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$Form.CancelButton = $CancelButton
$Form.Controls.Add($CancelButton)

################################################## メイン処理
#フォーム表示
$Form.Topmost = $true    #フォームを常に手前に表示
#$Form.Add_Shown({$TextBox_Arg.Select()})    #フォームをアクティブにし、テキストボックスにフォーカスを設定

$Timer.Start()    #タイマー起動
$result = $Form.ShowDialog()    #フォーム表示
$Timer.Stop()    #タイマー終了

If ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    If ($TextBox_Arg.Text -ne "") {
        Start-Process -FilePath $ScrName -ArgumentList $TextBox_Arg.Text    #外部プログラムexe起動
        #PowerShellで起動時は下記にて行う
        #$ArgStr = "-ExecutionPolicy RemoteSigned -command " + $ScrName + $TextBox_Arg.Text  #-WindowStyle Hidden 
        #Start-Process -FilePath powershell.exe -ArgumentList $ArgStr -Wait    #外部プログラムps1起動、終了待ち
    } Else {
        Start-Process -FilePath $ScrName    #パラメータ指定無いときは、そのまま外部プログラム起動
    }
}

################################################## 終了処理
Exit    #スクリプト終了

################################################## End Of File