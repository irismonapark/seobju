@echo off
chcp 65001 >nul
echo ========================================
echo  ResumeConverter - GitHub sync ^& push
echo  Target: irismonapark/seobju (ResumeConverter/)
echo ========================================
echo.

cd /d "%~dp0.."

echo [1/2] Local commit...
set "GIT=C:\Program Files\Git\bin\git.exe"
if not exist "%GIT%" set "GIT=C:\Program Files\Git\cmd\git.exe"
"%GIT%" add -A
"%GIT%" status --short
echo.
set /p LOCAL_MSG=Local commit message (Enter = skip): 
if not "%LOCAL_MSG%"=="" (
  "%GIT%" commit -m "%LOCAL_MSG%"
)

echo.
echo [2/2] Sync to seobju-github and push...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0sync-github.ps1" -Push -Message "(RESUME)깃허브저장완료"
if errorlevel 1 exit /b 1

echo.
echo ========================================
echo  Done! https://github.com/irismonapark/seobju/tree/main/ResumeConverter
echo ========================================
pause
