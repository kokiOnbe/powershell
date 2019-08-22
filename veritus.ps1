[void][reflection.assembly]::loadwithpartialname("system.windows.forms")

$url = "https://asia.nikkei.com/?n_cid=NARAN101"

# シェルを取得
$shell = New-Object -ComObject Shell.Application

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

$book = $excel.Workbooks.Open("\\10.27.46.23\hatsuban\（ヴェリタス）電子版発番表 2017.07～最新.xls")

Add-Type -AssemblyName Microsoft.VisualBasic
$excel_process = Get-Process -Name "EXCEL" | ? {$_.MainWindowHandle -eq $excel.HWND}
[Microsoft.VisualBasic.Interaction]::AppActivate($excel_process.ID) | Out-Null

Start-Sleep -s 15

#印刷プレビュー
[system.windows.forms.sendkeys]::sendwait("^p")
Start-Sleep -s 3
[system.windows.forms.sendkeys]::sendwait("%")
Start-Sleep -s 1
[system.windows.forms.sendkeys]::sendwait("p")
Start-Sleep -s 1
[system.windows.forms.sendkeys]::sendwait("n")
Start-Sleep -s 1
[system.windows.forms.sendkeys]::sendwait("4")
Start-Sleep -s 1
[system.windows.forms.sendkeys]::sendwait("{ENTER}")
Start-Sleep -s 1

#$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($book)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)