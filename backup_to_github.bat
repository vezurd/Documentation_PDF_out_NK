@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "GIT=C:\Program Files\Git\cmd\git.exe"
if not exist "%GIT%" (
  echo Git not found. Install Git for Windows first.
  pause
  exit /b 1
)

echo Saving project source to GitHub...
"%GIT%" add -A
"%GIT%" status
"%GIT%" diff --cached --quiet
if errorlevel 1 (
  "%GIT%" commit -m "backup %DATE% %TIME%"
  if errorlevel 1 (
    echo Commit failed.
    pause
    exit /b 1
  )
) else (
  echo No file changes to save.
)

"%GIT%" push -u origin main
if errorlevel 1 (
  echo Push failed. Check internet and GitHub login.
  pause
  exit /b 1
)

echo Done. Copy is on GitHub.
pause
