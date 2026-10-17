@echo off
setlocal
cd /d "%~dp0"

echo ========================================
echo Random Question Bank - Startup Check
echo ========================================
echo.

set "PYTHON="
where py >nul 2>nul
if not errorlevel 1 set "PYTHON=py"
if not defined PYTHON (
    where python >nul 2>nul
    if not errorlevel 1 set "PYTHON=python"
)

if not defined PYTHON goto install_python
%PYTHON% --version >nul 2>nul
if errorlevel 1 goto install_python
goto check_files

:install_python
echo [CHECK] Python was not found. Downloading the official Python installer...
echo An internet connection is required. Installation is for the current Windows user.
set "INSTALLER=%TEMP%\python-3.12.10-amd64.exe"
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe' -OutFile (Join-Path $env:TEMP 'python-3.12.10-amd64.exe') } catch { exit 1 }"
if errorlevel 1 (
    echo [ERROR] Python download failed. Check the internet connection and try again.
    goto show_error
)

"%INSTALLER%" /quiet InstallAllUsers=0 TargetDir="%LOCALAPPDATA%\Programs\Python\Python312" Include_exe=1 Include_pip=1 Include_launcher=0
if errorlevel 1 (
    echo [ERROR] Python installation failed.
    goto show_error
)
set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not exist "%PYTHON%" (
    echo [ERROR] Python installer completed, but python.exe was not found.
    goto show_error
)

echo [CHECK] Python installed successfully.

:check_files
if not exist "app.py" (
    echo [ERROR] app.py was not found. Run this file from the project folder.
    goto show_error
)
if not exist "requirements.txt" (
    echo [ERROR] requirements.txt was not found.
    goto show_error
)

echo.
echo [CHECK] Installing or checking required packages...
%PYTHON% -m pip --version >nul 2>nul
if errorlevel 1 (
    %PYTHON% -m ensurepip --upgrade
    if errorlevel 1 (
        echo [ERROR] pip is unavailable.
        goto show_error
    )
)
%PYTHON% -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Package installation failed. Check your internet connection and the messages above.
    goto show_error
)

echo.
echo [START] Starting the application...
echo Quiz page:  http://127.0.0.1:5000/
echo Admin page: http://127.0.0.1:5000/admin
echo Close this window to stop the server.
echo.
%PYTHON% app.py
if errorlevel 1 (
    echo [ERROR] The application failed to start. Check the messages above.
    goto show_error
)
goto end

:show_error
echo.
echo Press any key to close this window...
pause >nul
:end
endlocal