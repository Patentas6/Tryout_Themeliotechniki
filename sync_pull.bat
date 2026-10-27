@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ========================================================
echo   ANTIGRAVITY WORKSPACE SYNC - PULL (Git -> PC)
echo   Workspace: %CD%
echo ========================================================
echo.

git remote get-url origin >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [NOTE] No GitHub remote 'origin' is configured yet.
    echo Once linked, this script will pull the latest workspace and .agent-memory/ from GitHub.
    python "%~dp0sync_helper.py" pull
    pause
    exit /b 0
)

echo [1/2] Pulling latest workspace and .agent-memory/ from GitHub...
git pull --rebase --autostash

echo.
echo [2/2] Verifying .agent-memory/ files...
python "%~dp0sync_helper.py" pull

echo [DONE] Ready! Open this folder in Antigravity and start your session.
pause
