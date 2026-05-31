@echo off
chcp 65001 >nul
set "GIT=C:\Program Files\Git\cmd\git.exe"

echo ========================================
echo  abcos - GitHub sync ^& push
echo  Target: irismonapark/seobju (ABCOS/)
echo ========================================
echo.

if not exist "%GIT%" (
  echo [ERROR] Git not found at %GIT%
  echo         Install from https://git-scm.com/download/win
  exit /b 1
)

cd /d "%~dp0.."

echo [1/2] Local commit in abcos...
"%GIT%" add -A
"%GIT%" status --short
echo.
set /p LOCAL_MSG=Local commit message (Enter = skip local commit): 
if not "%LOCAL_MSG%"=="" (
  "%GIT%" commit -m "%LOCAL_MSG%"
)

echo.
echo [2/2] Sync to seobju-github and push...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0sync-github.ps1" -Push
if errorlevel 1 exit /b 1

echo.
echo ========================================
echo  Done! https://github.com/irismonapark/seobju
echo ========================================
pause
