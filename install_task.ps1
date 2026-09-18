# PCS 스케줄 자동 수집 - Windows 작업 스케줄러 등록
#
#   실행 방법: 이 파일을 우클릭 → "PowerShell에서 실행"
#             또는  powershell -ExecutionPolicy Bypass -File install_task.ps1
#
#   등록되는 작업 2개
#     PCS_Schedule_AM : 매일 09:00  →  "YYYY.MM.DD 오전" 스냅샷
#     PCS_Schedule_PM : 매일 18:00  →  "YYYY.MM.DD 오후" 스냅샷
#
#   PC가 꺼져 있어서 시간을 놓친 경우, 켜진 뒤 자동으로 밀린 작업을 실행합니다
#   (StartWhenAvailable). 이때도 오전/오후 이름은 예약된 시간대로 저장됩니다.

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

# Microsoft Store 앱 별칭(0바이트 스텁)은 작업 스케줄러에서 동작하지 않으므로
# 실제로 실행 중인 인터프리터 경로를 파이썬 자신에게 물어본다.
$real = (& python -c "import sys; print(sys.executable)").Trim()
if (-not $real -or -not (Test-Path $real)) {
    throw "파이썬을 찾을 수 없습니다. 명령 프롬프트에서 'python --version' 이 되는지 확인하세요."
}
$pythonw = Join-Path (Split-Path -Parent $real) "pythonw.exe"   # 창 없이 실행
if (-not (Test-Path $pythonw)) { $pythonw = $real }
Write-Host "사용할 파이썬: $pythonw"
Write-Host "작업 폴더    : $here"

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 10)

function Register-PcsTask {
    param([string]$Name, [string]$Session, [string]$Time, [string]$Desc)

    $action  = New-ScheduledTaskAction -Execute $pythonw `
                 -Argument "`"$here\scraper.py`" --session $Session" -WorkingDirectory $here
    $trigger = New-ScheduledTaskTrigger -Daily -At $Time

    if (Get-ScheduledTask -TaskName $Name -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $Name -Confirm:$false
        Write-Host "기존 작업 '$Name' 을 교체합니다."
    }
    Register-ScheduledTask -TaskName $Name -Action $action -Trigger $trigger `
        -Settings $settings -Description $Desc | Out-Null
    Write-Host "등록 완료: $Name  ($Time)"
}

Register-PcsTask -Name "PCS_Schedule_AM" -Session "AM" -Time "09:00" -Desc "PCS LINE 스케줄 오전 수집"
Register-PcsTask -Name "PCS_Schedule_PM" -Session "PM" -Time "18:00" -Desc "PCS LINE 스케줄 오후 수집"

Write-Host ""
Write-Host "완료되었습니다. 지금 바로 한 번 실행해서 확인하려면:" -ForegroundColor Green
Write-Host "    Start-ScheduledTask -TaskName PCS_Schedule_AM"
Write-Host ""
Write-Host "등록 해제하려면:"
Write-Host "    Unregister-ScheduledTask -TaskName PCS_Schedule_AM,PCS_Schedule_PM -Confirm:`$false"
