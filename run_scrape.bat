@echo off
chcp 65001 >nul
cd /d "%~dp0"
python scraper.py %*
if errorlevel 2 (
  echo.
  echo [!] 수집 중 오류가 발생했습니다. logs 폴더를 확인하세요.
  pause
)
