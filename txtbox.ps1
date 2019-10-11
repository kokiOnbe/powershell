# ユーザーフォームを作る　- チェックボックス編 -
function txtbox([string]$msglist)
{
# アセンブリのロード
Add-Type -AssemblyName System.Windows.Forms

# フォームの作成
$form = New-Object System.Windows.Forms.Form
$form.Size = "250,280"
$form.Startposition = "CenterScreen"
$form.Text = "選択"

# ラベルを作成
$label = New-Object System.Windows.Forms.Label
$label.Location = "10,20"
$label.Size = "250,20"

# チェックボックスを作成
$CheckedBox = New-Object System.Windows.Forms.CheckedListBox
$CheckedBox.Location = "5,40"
$CheckedBox.Size = "220,150"

# 配列を作成
$RETU = $msglist -split "、"

# チェックボックスに項目を追加
$CheckedBox.Items.AddRange($RETU)

# すべての既存の選択をクリア
$CheckedBox.ClearSelected()

# OKボタンの設定
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = "40,200"
$OKButton.Size = "75,30"
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

# キャンセルボタンの設定
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = "130,200"
$CancelButton.Size = "75,30"
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

# フォームにアイテムを追加
$form.Controls.Add($label)
$form.Controls.Add($OKButton)
$form.Controls.Add($CancelButton)
$form.Controls.Add($CheckedBox)

# キーとボタンの関係
$form.AcceptButton = $OKButton
$form.CancelButton = $CancelButton

# 最前面に表示：する
$form.Topmost = $True

# フォームを表示
$result = $Form.ShowDialog()

# 処理分岐
if ( $result -eq "OK" )
{
    $AAA = @($CheckedBox.CheckedItems)
}else{
    exit
}

}