[void][reflection.assembly]::loadwithpartialname("system.windows.forms")

$url = "https://asia.nikkei.com/?n_cid=NARAN101"
# �V�F�����擾
$shell = New-Object -ComObject Shell.Application
# IE�N��
$ie = New-Object -ComObject InternetExplorer.Application
# ����
$ie.Visible = $true
# HWND���L��
$hwnd = $ie.HWND
# URL�I�[�v��(�L���b�V������)
$ie.Navigate($url,4)

# IE�A�N�e�B�u��
Add-Type -AssemblyName Microsoft.VisualBasic
$window_process = Get-Process -Name "iexplore" | ? {$_.MainWindowHandle -eq $ie.HWND}
[Microsoft.VisualBasic.Interaction]::AppActivate($window_process.ID) | Out-Null

while ($ie.busy -or $ie.readystate -ne 4) {
    Start-Sleep -s 1
}

#�p�\�R���d����
Start-Sleep -s 10

#����v���r���[
[system.windows.forms.sendkeys]::sendwait("%")
Start-Sleep -s 2
[system.windows.forms.sendkeys]::sendwait("f")
Start-Sleep -s 2
[system.windows.forms.sendkeys]::sendwait("v")
Start-Sleep -s 13

#�T�C�Y�ύX60%
[system.windows.forms.sendkeys]::sendwait("(%s)")
Start-Sleep -s 4
[system.windows.forms.sendkeys]::sendwait("{DOWN 3}")
Start-Sleep -s 5

#����ݒ�
[system.windows.forms.sendkeys]::sendwait("(%p)")
Start-Sleep -s 3

#�J���[�ɕύX
[system.windows.forms.sendkeys]::sendwait("(%r)")
Start-Sleep -s 3
[system.windows.forms.sendkeys]::sendwait("{DOWN 4}")
Start-Sleep -s 2

#������s
[system.windows.forms.sendkeys]::sendwait("(%a)")
Start-Sleep -s 2
[system.windows.forms.sendkeys]::sendwait("{ENTER}")
Start-Sleep -s 1
[system.windows.forms.sendkeys]::sendwait("{ENTER}")
Start-Sleep -s 10

#����
$ie.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie)
