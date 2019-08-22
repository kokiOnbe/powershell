[void][reflection.assembly]::loadwithpartialname("system.windows.forms")

$url = "https://asia.nikkei.com/?n_cid=NARAN101"
# シェルを取得
$shell = New-Object -ComObject Shell.Application
# IE起動
$ie = New-Object -ComObject InternetExplorer.Application
# 可視化
$ie.Visible = $true
# HWNDを記憶
$hwnd = $ie.HWND
# URLオープン(キャッシュ無効)
$ie.Navigate($url,4)

# IEアクティブ化
Add-Type -AssemblyName Microsoft.VisualBasic
$window_process = Get-Process -Name "iexplore" | ? {$_.MainWindowHandle -eq $ie.HWND}
[Microsoft.VisualBasic.Interaction]::AppActivate($window_process.ID) | Out-Null

while ($ie.busy -or $ie.readystate -ne 4) {
    Start-Sleep -s 1
}

#パソコン重すぎ
Start-Sleep -s 10

#印刷プレビュー
[system.windows.forms.sendkeys]::sendwait("%")
Start-Sleep -s 2
[system.windows.forms.sendkeys]::sendwait("f")
Start-Sleep -s 2
[system.windows.forms.sendkeys]::sendwait("v")
Start-Sleep -s 13

#サイズ変更60%
[system.windows.forms.sendkeys]::sendwait("(%s)")
Start-Sleep -s 4
[system.windows.forms.sendkeys]::sendwait("{DOWN 3}")
Start-Sleep -s 5

#印刷設定
[system.windows.forms.sendkeys]::sendwait("(%p)")
Start-Sleep -s 3

#カラーに変更
[system.windows.forms.sendkeys]::sendwait("(%r)")
Start-Sleep -s 3
[system.windows.forms.sendkeys]::sendwait("{DOWN 4}")
Start-Sleep -s 2

#印刷実行
[system.windows.forms.sendkeys]::sendwait("(%a)")
Start-Sleep -s 2
[system.windows.forms.sendkeys]::sendwait("{ENTER}")
Start-Sleep -s 1
[system.windows.forms.sendkeys]::sendwait("{ENTER}")
Start-Sleep -s 10

#閉じる
$ie.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie)
