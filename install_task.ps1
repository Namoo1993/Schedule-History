# PCS 스케줄 자동 수집 - Windows 작업 스케줄러 등록
#
#   실행 방법: 이 파일을 우클릭 → "PowerShell에서 실행"  (관리자 권한 권장)
#             또는  powershell -ExecutionPolicy Bypass -File install_task.ps1
#
#   등록되는 작업 2개
#     PCS_Schedule_AM : 매일 09:00  →  "YYYY.MM.DD 오전" 스냅샷
#     PCS_Schedule_PM : 매일 18:00  →  "YYYY.MM.DD 오후" 스냅샷
#
#   각 작업은 run_and_push.ps1 을 실행한다 (수집 → 성공 시 GitHub 푸시).
#
#   ── 절전 상태에서도 수집되게 하는 설정 ──────────────────────────────
#   PC 를 끄지 않고 절전으로 두면 예약 시각에 스스로 깨어나 수집하고 다시 잠든다.
#   그러려면 두 가지가 모두 필요하다.
#     1) 작업에 WakeToRun (절전 모드 해제하여 실행)
#     2) 전원 옵션의 RTCWAKE (절전 모드 해제 타이머 허용) 가 "사용" 이어야 함
#   2번은 기본값이 꺼져 있어서 이 스크립트가 켜준다. 관리자 권한이 필요하다.
#
#   완전히 종료(셧다운)한 경우에는 깨울 수 없다. 절전 또는 최대 절전으로 두어야 한다.
#
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

$isAdmin = ([Security.Principal.WindowsPrincipal] `
            [Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

# --- 실행 파일 경로 -------------------------------------------------------
# Microsoft Store 앱 별칭(0바이트 스텁)은 작업 스케줄러에서 동작하지 않으므로
# 실제로 실행 중인 인터프리터 경로를 파이썬 자신에게 물어본다.
$python = (& python -c "import sys; print(sys.executable)").Trim()
if (-not $python -or -not (Test-Path $python)) {
    throw "파이썬을 찾을 수 없습니다. 명령 프롬프트에서 'python --version' 이 되는지 확인하세요."
}
$gitCmd = Get-Command git -ErrorAction SilentlyContinue
$git = if ($gitCmd) { $gitCmd.Source } else { "" }
if (-not $git) { Write-Warning "git 을 찾지 못했습니다. 수집은 되지만 GitHub 반영이 안 됩니다." }

Write-Host "사용할 파이썬: $python"
Write-Host "사용할 git   : $git"
Write-Host "작업 폴더    : $here"

# --- 1. 전원 옵션: 깨우기 타이머 허용 -------------------------------------
$UNATTENDSLEEP = "7bc4a2f9-d8fc-4469-b07b-33eb785aaca0"   # 무인 시스템 절전 제한시간

if ($isAdmin) {
    Write-Host ""
    Write-Host "전원 옵션 조정 중..." -ForegroundColor Cyan
    # 예약 시각에 스스로 깨어날 수 있게 한다 (기본값은 '사용 안 함')
    powercfg /setacvalueindex SCHEME_CURRENT SUB_SLEEP RTCWAKE 1
    powercfg /setdcvalueindex SCHEME_CURRENT SUB_SLEEP RTCWAKE 1

    # 타이머로 깨어난 뒤(사용자 조작이 없으면) 다시 잠들기까지의 시간.
    # 수집은 2분 내에 끝나므로 5분이면 충분하다. 이 설정은 기본적으로 숨겨져 있어
    # 먼저 노출시킨 다음 값을 넣는다.
    powercfg /attributes SUB_SLEEP $UNATTENDSLEEP -ATTRIB_HIDE
    powercfg /setacvalueindex SCHEME_CURRENT SUB_SLEEP $UNATTENDSLEEP 300

    powercfg /setactive SCHEME_CURRENT
    Write-Host "  깨우기 타이머 허용 = 사용"
    Write-Host "  무인 절전 제한시간 = 5분 (수집 후 자동으로 다시 잠듦)"
} else {
    Write-Warning "관리자 권한이 아니라 전원 옵션을 바꾸지 못했습니다."
    Write-Warning "이 상태로는 절전 중에 깨어나지 않습니다. 관리자 권한으로 다시 실행하세요."
}

# --- 2. 작업 등록 ---------------------------------------------------------
$settings = New-ScheduledTaskSettingsSet `
    -WakeToRun `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 10) `
    -MultipleInstances IgnoreNew

function Register-PcsTask {
    param([string]$Name, [string]$Session, [string]$Time, [string]$Desc)

    $argline = ('-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass ' +
                '-File "{0}\run_and_push.ps1" -Session {1} -Python "{2}" -Git "{3}"' `
                -f $here, $Session, $python, $git)

    $action  = New-ScheduledTaskAction -Execute "powershell.exe" `
                 -Argument $argline -WorkingDirectory $here
    $trigger = New-ScheduledTaskTrigger -Daily -At $Time

    if (Get-ScheduledTask -TaskName $Name -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $Name -Confirm:$false
        Write-Host "기존 작업 '$Name' 을 교체합니다."
    }
    Register-ScheduledTask -TaskName $Name -Action $action -Trigger $trigger `
        -Settings $settings -Description $Desc | Out-Null
    Write-Host "등록 완료: $Name  ($Time)"
}

Write-Host ""
Register-PcsTask -Name "PCS_Schedule_AM" -Session "AM" -Time "09:00" -Desc "PCS LINE 스케줄 오전 수집"
Register-PcsTask -Name "PCS_Schedule_PM" -Session "PM" -Time "18:00" -Desc "PCS LINE 스케줄 오후 수집"

# --- 3. 확인 -------------------------------------------------------------
Write-Host ""
Write-Host "현재 상태" -ForegroundColor Green
Get-ScheduledTask -TaskName "PCS_Schedule_*" | ForEach-Object {
    "  {0}  WakeToRun={1}  다음실행={2}" -f `
        $_.TaskName, $_.Settings.WakeToRun, (Get-ScheduledTaskInfo $_.TaskName).NextRunTime
}
# powercfg 출력은 Windows 표시 언어에 따라 달라지므로 한글/영문 문자열을 찾지 않고
# 16진 값만 뽑는다. 출력 순서가 AC -> DC 라서 첫 값이 AC 설정이다.
$hex = [regex]::Matches((powercfg /q SCHEME_CURRENT SUB_SLEEP RTCWAKE | Out-String),
                        '0x[0-9A-Fa-f]{8}')
if ($hex.Count -ge 1) {
    $ac = [Convert]::ToInt32($hex[0].Value, 16)
    $txt = @{ 0 = "사용 안 함"; 1 = "사용"; 2 = "중요한 타이머만" }[$ac]
    if (-not $txt) { $txt = "알 수 없음($ac)" }
    Write-Host "  깨우기 타이머(AC): $txt"
    if ($ac -eq 0) {
        Write-Host "    ^ 이 값이 '사용' 이 아니면 절전 중에 깨어나지 않습니다." -ForegroundColor Red
        Write-Host "      관리자 권한 PowerShell 에서 이 스크립트를 다시 실행하세요." -ForegroundColor Red
    }
} else {
    Write-Host "  깨우기 타이머(AC): 확인 실패"
}

Write-Host ""
Write-Host "지금 바로 한 번 실행해서 확인하려면:" -ForegroundColor Green
Write-Host "    Start-ScheduledTask -TaskName PCS_Schedule_AM"
Write-Host "    Get-Content logs\task_$(Get-Date -Format 'yyyy-MM').log -Tail 20"
Write-Host ""
Write-Host "등록 해제하려면:"
Write-Host "    Unregister-ScheduledTask -TaskName PCS_Schedule_AM,PCS_Schedule_PM -Confirm:`$false"
Write-Host ""
Write-Host "!! 주말에도 수집되게 하려면 노트북을 '종료' 하지 말고 '절전' 으로 두고," -ForegroundColor Yellow
Write-Host "   전원 어댑터를 꽂아 두세요. 완전히 꺼진 상태에서는 깨울 수 없습니다." -ForegroundColor Yellow
