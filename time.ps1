Add-Type -AssemblyName Microsoft.VisualBasic

. "C:\Users\ns31896\Documents\txtbox.ps1"

<#
.SYNOPSIS
    時間が来たらメッセージ表示

.DESCRIPTION
    指定時間（分）後 または 指定時刻になったらメッセージを表示する

.PARAMETER time
    分または時刻（HH:mm形式）

.PARAMETER msg
    時間が来たら表示するカスタムメッセージ（省略可）

.NOTES
    時刻指定で現在時刻より前を指定した場合、skip
　　flagで判定

.EXAMPLE
    zzz 10
    10分待つ

.EXAMPLE
    zzz 08:30
    午前8時半まで待つ

.EXAMPLE
    zzz 23:00 もう寝ましょう
    午後23時まで待つ（任意のメッセージ表示）
#>

function zzz([string]$time, [string]$msg = "", [string]$boxtype = "txtbox")
{
    Clear-Host
    #$flag = 0

    [string]$mode = ""
    # 時刻指定っぽい？
    if ($time -match "^\d+:\d+$")
    {
        # 厳密にチェック
        if ($time -match "^([0-1][0-9])|([2][0-3]):[0-5][0-9]$")
        {
            $mode = "TIME";
        }
        else
        {
            Write-Host "時と分はそれぞれ2桁で、実在する時刻を指定してください"
            Write-Host "例1) 01:00"
            Write-Host "例2) 23:59"
            return
        }
    }
    elseif ($time -match "^[0-9]+$")
    {
        # 数字のみなら「分」指定とみなす
        $mode = "MINUTES";
    }

    if ($mode -eq "")
    {
        # パラメータ不正時は使用方法表示
        Write-Host "使用例1) 10分後にメッセージ表示"
        Write-Host "  zzz 10"
        Write-Host ""
        Write-Host "使用例2) 3分後に任意のメッセージ表示"
        Write-Host "  zzz 3 ラーメンできたよ！"
        Write-Host ""
        Write-Host "使用例3) 13時5分にメッセージ表示(時、分はそれぞれ2桁で指定)"
        Write-Host "         ※現在時刻より前を指定した場合、翌日の時刻とみなす"
        Write-Host "  zzz 13:05"
        return
    }

    $now = Get-Date
    Write-Host $now.ToString("HH:mm:ss 待ち開始") -ForegroundColor DarkGray
    #Write-Host "`n"

    if ($mode -eq "TIME")
    {
        # 「時刻」指定

        [string]$nowTime = $now.ToString("HH:mm")
        [string]$msgTime = $time
        [int]$waitSecond = 0

        if ($nowTime -eq $time)
        {
            Write-Host "指定した時刻と現在時刻が同じです。"
	    $flag = 1
            return
        }
        elseif ($nowTime -lt $time)
        {
            # 現在時刻 ＜ 指定時刻

            # 現在時刻～指定時刻の差（秒）を求める
            $waitSecond = (New-TimeSpan $now.ToString("HH:mm:ss") "${time}:00").TotalSeconds
        }
        else
        {
            # 現在時刻 ＞ 指定時刻
	     Write-Host "指定した時刻が過ぎています。"
	     $flag = 0
             return
        }

        Write-Host "${msgTime}になったらメッセージを表示します。"
        foreach ($i in (1..$waitSecond))
        {
            # 待ち
            Start-Sleep -Second 1

            if ((Get-Date -Format "HH:mm") -eq $time)
            {
                # 指定した時間が来たので終了
		$flag = 1
                break
            }
        }
    }
    else
    {
        # 「分」指定

        Write-Host "${time}分後にメッセージを表示します。　Zzz…(u_u)"
        [int]$waitSecond = [int]$time * 60

        # 指定した秒数分、繰り返す
        foreach ($i in (1..$waitSecond))
        {
            # 待ち
            Start-Sleep -Second 1
        }
    }

    # 目覚め
    #Write-Host "`n`n(@_@)`n"
    $flag = 0

    if ($msg -eq "")
    {
        # 引数でメッセージが指定されていない場合はデフォルトのメッセージを使う
        $msg = "`n`n　時間です　" + (Get-Date).ToString("HH:mm") + "n`n"
    }

    
    # onならtxtbox
    if ($boxtype -eq "check"){
        txtbox $msg
    }
    else
    {
    # メッセージボックス表示
    [void][Microsoft.VisualBasic.Interaction]::MsgBox($msg,
        ([Microsoft.VisualBasic.MsgBoxStyle]::Yes -bor [Microsoft.VisualBasic.MsgBoxStyle]::SystemModal),"PowerShell")
    }
}







