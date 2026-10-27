@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ========================================================
echo   ANTIGRAVITY WORKSPACE SYNC - PULL (GitHub -> PC)
echo   Workspace: %CD%
echo   Remote:    https://github.com/Patentas6/Tryout_Themeliotechniki.git
echo ========================================================
echo.

REM 1. Ensure GitHub account is authenticated in Git Credential Manager
for /f %%i in ('git credential-manager github list 2^>nul') do set GH_USER=%%i
if "%GH_USER%"=="" (
    echo [AUTH] First-time GitHub login on this PC. Opening browser to authorize...
    git credential-manager github login --web
)

echo [1/2] Pulling latest workspace and .agent-memory/ from GitHub...
git pull --rebase --autostash origin main

echo.
echo [2/2] Verifying .agent-memory/ files...
python "%~dp0sync_helper.py" pull

echo [DONE] Ready! Open this folder in Antigravity and start your session.
pause
