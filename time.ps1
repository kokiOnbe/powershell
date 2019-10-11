Add-Type -AssemblyName Microsoft.VisualBasic

. "C:\Users\ns31896\Documents\txtbox.ps1"

<#
.SYNOPSIS
    ���Ԃ������烁�b�Z�[�W�\��

.DESCRIPTION
    �w�莞�ԁi���j�� �܂��� �w�莞���ɂȂ����烁�b�Z�[�W��\������

.PARAMETER time
    ���܂��͎����iHH:mm�`���j

.PARAMETER msg
    ���Ԃ�������\������J�X�^�����b�Z�[�W�i�ȗ��j

.NOTES
    �����w��Ō��ݎ������O���w�肵���ꍇ�Askip
�@�@flag�Ŕ���

.EXAMPLE
    zzz 10
    10���҂�

.EXAMPLE
    zzz 08:30
    �ߑO8�����܂ő҂�

.EXAMPLE
    zzz 23:00 �����Q�܂��傤
    �ߌ�23���܂ő҂i�C�ӂ̃��b�Z�[�W�\���j
#>

function zzz([string]$time, [string]$msg = "", [string]$boxtype = "txtbox")
{
    Clear-Host
    #$flag = 0

    [string]$mode = ""
    # �����w����ۂ��H
    if ($time -match "^\d+:\d+$")
    {
        # �����Ƀ`�F�b�N
        if ($time -match "^([0-1][0-9])|([2][0-3]):[0-5][0-9]$")
        {
            $mode = "TIME";
        }
        else
        {
            Write-Host "���ƕ��͂��ꂼ��2���ŁA���݂��鎞�����w�肵�Ă�������"
            Write-Host "��1) 01:00"
            Write-Host "��2) 23:59"
            return
        }
    }
    elseif ($time -match "^[0-9]+$")
    {
        # �����݂̂Ȃ�u���v�w��Ƃ݂Ȃ�
        $mode = "MINUTES";
    }

    if ($mode -eq "")
    {
        # �p�����[�^�s�����͎g�p���@�\��
        Write-Host "�g�p��1) 10����Ƀ��b�Z�[�W�\��"
        Write-Host "  zzz 10"
        Write-Host ""
        Write-Host "�g�p��2) 3����ɔC�ӂ̃��b�Z�[�W�\��"
        Write-Host "  zzz 3 ���[�����ł�����I"
        Write-Host ""
        Write-Host "�g�p��3) 13��5���Ƀ��b�Z�[�W�\��(���A���͂��ꂼ��2���Ŏw��)"
        Write-Host "         �����ݎ������O���w�肵���ꍇ�A�����̎����Ƃ݂Ȃ�"
        Write-Host "  zzz 13:05"
        return
    }

    $now = Get-Date
    Write-Host $now.ToString("HH:mm:ss �҂��J�n") -ForegroundColor DarkGray
    #Write-Host "`n"

    if ($mode -eq "TIME")
    {
        # �u�����v�w��

        [string]$nowTime = $now.ToString("HH:mm")
        [string]$msgTime = $time
        [int]$waitSecond = 0

        if ($nowTime -eq $time)
        {
            Write-Host "�w�肵�������ƌ��ݎ����������ł��B"
	    $flag = 1
            return
        }
        elseif ($nowTime -lt $time)
        {
            # ���ݎ��� �� �w�莞��

            # ���ݎ����`�w�莞���̍��i�b�j�����߂�
            $waitSecond = (New-TimeSpan $now.ToString("HH:mm:ss") "${time}:00").TotalSeconds
        }
        else
        {
            # ���ݎ��� �� �w�莞��
	     Write-Host "�w�肵���������߂��Ă��܂��B"
	     $flag = 0
             return
        }

        Write-Host "${msgTime}�ɂȂ����烁�b�Z�[�W��\�����܂��B"
        foreach ($i in (1..$waitSecond))
        {
            # �҂�
            Start-Sleep -Second 1

            if ((Get-Date -Format "HH:mm") -eq $time)
            {
                # �w�肵�����Ԃ������̂ŏI��
		$flag = 1
                break
            }
        }
    }
    else
    {
        # �u���v�w��

        Write-Host "${time}����Ƀ��b�Z�[�W��\�����܂��B�@Zzz�c(u_u)"
        [int]$waitSecond = [int]$time * 60

        # �w�肵���b�����A�J��Ԃ�
        foreach ($i in (1..$waitSecond))
        {
            # �҂�
            Start-Sleep -Second 1
        }
    }

    # �ڊo��
    #Write-Host "`n`n(@_@)`n"
    $flag = 0

    if ($msg -eq "")
    {
        # �����Ń��b�Z�[�W���w�肳��Ă��Ȃ��ꍇ�̓f�t�H���g�̃��b�Z�[�W���g��
        $msg = "`n`n�@���Ԃł��@" + (Get-Date).ToString("HH:mm") + "n`n"
    }

    
    # on�Ȃ�txtbox
    if ($boxtype -eq "check"){
        txtbox $msg
    }
    else
    {
    # ���b�Z�[�W�{�b�N�X�\��
    [void][Microsoft.VisualBasic.Interaction]::MsgBox($msg,
        ([Microsoft.VisualBasic.MsgBoxStyle]::Yes -bor [Microsoft.VisualBasic.MsgBoxStyle]::SystemModal),"PowerShell")
    }
}







