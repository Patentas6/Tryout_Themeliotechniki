@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ========================================================
echo   ANTIGRAVITY WORKSPACE SYNC - PUSH (PC -> Git)
echo   Workspace: %CD%
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
    for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set TODAY=%%a-%%b-%%c
    git commit -m "Sync workspace & .agent-memory [%DATE% %TIME%]"
)

REM 3. Push to remote if configured
git remote get-url origin >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [NOTE] Local commit succeeded, but no GitHub remote 'origin' is linked yet.
    echo To link your private GitHub repository once, run:
    echo   git remote add origin https://github.com/YOUR_USERNAME/YOUR_PRIVATE_REPO.git
    echo   git branch -M main
    echo   git push -u origin main
    echo.
    pause
    exit /b 0
)

echo.
echo Pushing to remote repository (origin)...
git push
if %ERRORLEVEL% NEQ 0 (
    echo [!] git push encountered an issue. Trying 'git push -u origin main'...
    git push -u origin main
)

echo.
echo [DONE] Workspace and .agent-memory/ synced!
pause
