@echo off
chcp 65001 >nul
echo ========================================
echo  abcos - GitHub push helper
echo ========================================
echo.

where git >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Git이 설치되어 있지 않습니다.
  echo         https://git-scm.com/download/win 에서 설치 후 다시 실행하세요.
  exit /b 1
)

cd /d "%~dp0"

echo 1. GitHub에서 새 저장소 생성 (https://github.com/new)
echo    - Repository name: abcos (원하는 이름)
echo    - Public 또는 Private 선택
echo    - README / .gitignore 추가하지 말 것 (빈 repo)
echo.
set /p GITHUB_URL=2. GitHub 저장소 URL 입력 (예: https://github.com/아이디/abcos.git): 

if "%GITHUB_URL%"=="" (
  echo URL이 비어 있습니다.
  exit /b 1
)

git status >nul 2>&1
if errorlevel 1 (
  git init
  git branch -M main
)

git remote remove origin 2>nul
git remote add origin "%GITHUB_URL%"

git add .
git status
echo.
set /p CONFIRM=위 파일들을 커밋하고 push 할까요? (y/N): 
if /i not "%CONFIRM%"=="y" exit /b 0

git commit -m "Prepare payroll system for Vercel deployment" 2>nul
if errorlevel 1 (
  echo 변경사항이 없거나 커밋 실패. git status 를 확인하세요.
)

git push -u origin main
if errorlevel 1 (
  echo.
  echo [TIP] push 실패 시:
  echo   - GitHub 로그인: gh auth login 또는 Git Credential Manager
  echo   - 브랜치가 master면: git push -u origin master
  exit /b 1
)

echo.
echo ========================================
echo  GitHub push 완료!
echo  다음: https://vercel.com/new 에서 Import Git Repository
echo ========================================
pause
