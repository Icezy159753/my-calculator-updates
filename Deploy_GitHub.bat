@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>&1
if errorlevel 1 (
    python -X utf8 "%~dp0tools\deploy_release.py" %*
) else (
    py -3 -X utf8 "%~dp0tools\deploy_release.py" %*
)
set "DEPLOY_EXIT=%ERRORLEVEL%"
echo.
pause
exit /b %DEPLOY_EXIT%
