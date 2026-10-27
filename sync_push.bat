@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ========================================================
echo   ANTIGRAVITY WORKSPACE SYNC - PUSH (PC -> GitHub)
echo   Workspace: %CD%
echo   Remote:    https://github.com/Patentas6/Tryout_Themeliotechniki.git
echo ========================================================
echo.

REM 1. Copy latest brain markdown files into .agent-memory\
python "%~dp0sync_helper.py" push

REM 2. Stage and commit all changes in workspace and .agent-memory\
echo [2/2] Staging and committing workspace + .agent-memory/...
git add -A

git diff --cached --quiet
if %ERRORLEVEL% EQU 0 (
    echo   -> No new local file changes to commit.
) else (
    git commit -m "Sync workspace & .agent-memory [%DATE% %TIME%]"
)

REM 3. Ensure GitHub account is authenticated in Git Credential Manager
for /f %%i in ('git credential-manager github list 2^>nul') do set GH_USER=%%i
if "%GH_USER%"=="" (
    echo.
    echo [AUTH] First-time GitHub login on this PC. Opening browser to authorize...
    git credential-manager github login --web
)

REM 4. Push to remote
echo.
echo Pushing to GitHub (origin/main)...
git push -u origin main
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [!] Push did not complete. Please verify your GitHub login or repository access.
    pause
    exit /b 1
)

echo.
echo [DONE] Workspace and .agent-memory/ synced to GitHub!
pause
