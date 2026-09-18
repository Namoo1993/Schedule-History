@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem 수집만 하는 것이 아니라 GitHub 반영까지 한다.
rem scraper.py 만 돌리면 로컬 data 폴더만 바뀌고 뷰어(GitHub Pages)는 갱신되지 않는다.
rem
rem   run_scrape.bat        현재 시각으로 오전/오후 자동 판정
rem   run_scrape.bat AM     오전 스냅샷으로 강제
rem   run_scrape.bat PM     오후 스냅샷으로 강제

if "%~1"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_and_push.ps1"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_and_push.ps1" -Session %~1
)

echo.
echo 자세한 기록은 logs 폴더의 task_*.log 에 남습니다.
pause
