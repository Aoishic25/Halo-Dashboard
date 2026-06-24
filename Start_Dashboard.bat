@echo off
title HALO Dashboard Server
cd /d "%~dp0"

:: Find Python
set PYTHON=
for /f "delims=" %%i in ('where python 2^>nul') do (
    set PYTHON=%%i
    goto :found
)
if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" (
    set PYTHON=%LOCALAPPDATA%\Programs\Python\Python314\python.exe
    goto :found
)
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
    set PYTHON=%LOCALAPPDATA%\Programs\Python\Python313\python.exe
    goto :found
)
:found

if "%PYTHON%"=="" (
    echo ERROR: Python not found. Install Python and try again.
    pause
    exit /b 1
)

echo.
echo   HALO Dashboard Server
echo   =====================
echo   Local:  http://localhost:8765
echo   Press Ctrl+C to stop.
echo.

start "" cmd /c "timeout /t 2 /nobreak >nul & start http://localhost:8765"
"%PYTHON%" _server.py
pause
