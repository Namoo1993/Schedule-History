# PCS 스케줄 수집 + GitHub 반영
#
#   작업 스케줄러가 호출한다. 직접 실행할 때는:
#       powershell -ExecutionPolicy Bypass -File run_and_push.ps1 -Session AM
#
#   수집이 성공했을 때만 커밋/푸시한다. (scraper.py 는 0건이면 종료코드 1 을 준다)
#
param(
    # 생략하면 현재 시각으로 오전/오후를 판정한다 (배치 파일에서 인수 없이 부를 때)
    [ValidateSet("AM", "PM", "")][string]$Session = "",
    [string]$Python = "",
    [string]$Git = ""
)

if (-not $Session) { $Session = if ((Get-Date).Hour -lt 12) { "AM" } else { "PM" } }

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here
New-Item -ItemType Directory -Force (Join-Path $here "logs") | Out-Null
$log = Join-Path $here ("logs\task_{0}.log" -f (Get-Date -Format "yyyy-MM"))

# git 과 python 은 UTF-8 로 출력하는데 Windows PowerShell 5.1 은 콘솔 코드페이지
# (한국어 환경이면 CP949)로 해석한다. 그대로 두면 로그의 한글이 깨져서 장애가 났을 때
# 읽을 수 없다. 양쪽을 UTF-8 로 맞춘다. 콘솔이 없는 환경에서는 조용히 넘어간다.
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $env:PYTHONIOENCODING = 'utf-8'
} catch { }

function W($m) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $m
    Write-Host $line
    Add-Content -Path $log -Value $line -Encoding utf8
}

# --- 실행 파일 경로 -------------------------------------------------------
# 작업 스케줄러 환경은 PATH 가 로그인 세션과 다를 수 있어 절대경로를 우선한다.
if (-not $Python -or -not (Test-Path $Python)) {
    $c = Get-Command python -ErrorAction SilentlyContinue
    if ($c) { $Python = $c.Source }
}
if (-not $Git -or -not (Test-Path $Git)) {
    $c = Get-Command git -ErrorAction SilentlyContinue
    if ($c) { $Git = $c.Source }
    elseif (Test-Path "C:\Program Files\Git\cmd\git.exe") { $Git = "C:\Program Files\Git\cmd\git.exe" }
}

# --- 수집이 끝나기 전에 다시 잠들지 않게 막는다 ---------------------------
# 타이머로 깨어난 뒤 사용자 조작이 없으면 Windows 는 '무인 절전 제한시간'(5분)
# 뒤에 스스로 다시 잠든다. 작업 스케줄러는 실행 중인 작업을 이유로 절전을
# 막아주지 않기 때문에, 항로가 늘어 수집이 길어지면 도중에 잘릴 수 있다.
# SetThreadExecutionState 로 "시스템이 필요한 상태" 를 선언해 둔다.
# 이 선언은 프로세스가 끝나면 자동으로 풀리므로 따로 해제하지 않아도 된다.
try {
    Add-Type -Namespace Native -Name Power -MemberDefinition @'
[DllImport("kernel32.dll", SetLastError = true)]
public static extern uint SetThreadExecutionState(uint esFlags);
'@ -ErrorAction Stop
    $ES_CONTINUOUS      = [uint32]2147483648   # 0x80000000
    $ES_SYSTEM_REQUIRED = [uint32]1            # 0x00000001
    [Native.Power]::SetThreadExecutionState($ES_CONTINUOUS -bor $ES_SYSTEM_REQUIRED) | Out-Null
    $held = $true
} catch {
    $held = $false
}

W "=========================================================="
W "$Session 수집 작업 시작"
if (-not $Python) { W "중단: 파이썬을 찾지 못했습니다."; exit 2 }
if ($held) { W "절전 방지 설정 적용 (수집 중 다시 잠들지 않음)" }
else       { W "경고: 절전 방지 설정 실패. 수집이 5분을 넘기면 중단될 수 있습니다." }

# --- 절전에서 깨어난 직후엔 네트워크가 아직 안 붙어 있다 -------------------
# 무선 재연결에 보통 몇 초, 늦으면 30초 넘게 걸린다. 사이트가 응답할 때까지 기다린다.
$ready = $false
foreach ($i in 1..30) {
    try {
        $r = Invoke-WebRequest -Uri "https://ebiz.pcsline.co.kr/" -Method Head `
                               -TimeoutSec 10 -UseBasicParsing -ErrorAction Stop
        if ($r.StatusCode -eq 200) { $ready = $true; break }
    } catch { }
    Start-Sleep -Seconds 5
}
if ($ready) { W "네트워크 준비 완료" }
else        { W "경고: 사이트 확인 실패. 그래도 수집을 시도합니다." }

# --- 수집 ---------------------------------------------------------------
& $Python (Join-Path $here "scraper.py") --session $Session 2>&1 | ForEach-Object { W "  $_" }
$code = $LASTEXITCODE
W "scraper.py 종료코드: $code"

if ($code -ne 0) {
    W "수집 실패 - 커밋하지 않고 종료합니다. (logs 폴더 확인)"
    exit $code
}

# --- GitHub 반영 --------------------------------------------------------
# 이게 있어야 GitHub Pages 뷰어가 갱신된다. 예전에는 Actions 가 해주던 일.
if (-not $Git) { W "경고: git 을 찾지 못해 푸시를 건너뜁니다."; exit 0 }

& $Git add data/ 2>&1 | ForEach-Object { W "  git: $_" }
& $Git diff --staged --quiet
if ($LASTEXITCODE -eq 0) {
    W "데이터 변경 없음 - 푸시 생략"
    exit 0
}

$msg = "수집: {0} {1}" -f (Get-Date -Format "yyyy.MM.dd"), $Session
& $Git commit -m $msg 2>&1 | ForEach-Object { W "  git: $_" }

# 다른 데서 먼저 올라간 커밋이 있으면 rebase 로 붙인다 (충돌 시 푸시 포기)
& $Git pull --rebase --autostash 2>&1 | ForEach-Object { W "  git: $_" }
if ($LASTEXITCODE -ne 0) {
    W "경고: pull --rebase 실패. 다음 실행 때 다시 시도합니다."
    exit 0
}

& $Git push 2>&1 | ForEach-Object { W "  git: $_" }
if ($LASTEXITCODE -eq 0) { W "푸시 완료 - 잠시 뒤 GitHub Pages 가 갱신됩니다." }
else { W "경고: 푸시 실패. 다음 실행 때 다시 시도합니다." }

exit 0
