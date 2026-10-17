@echo off
setlocal
cd /d "%~dp0"

echo ========================================
echo Random Question Bank - Startup Check
echo ========================================
echo.

where py >nul 2>nul
if not errorlevel 1 goto use_py
where python >nul 2>nul
if not errorlevel 1 goto use_python

echo [ERROR] Python was not found.
echo Install Python 3.10 or newer and enable "Add Python to PATH".
echo Then run this file again.
goto failed

:use_py
set "PYTHON=py"
goto check_python

:use_python
set "PYTHON=python"

:check_python
%PYTHON% --version
if errorlevel 1 (
    echo [ERROR] Python could not start.
    goto failed
)

if not exist "app.py" (
    echo [ERROR] app.py was not found. Run this file from the project folder.
    goto failed
)
if not exist "requirements.txt" (
    echo [ERROR] requirements.txt was not found.
    goto failed
)

echo.
echo [CHECK] Installing or checking required packages...
%PYTHON% -m pip --version >nul 2>nul
if errorlevel 1 (
    echo [CHECK] Enabling pip...
    %PYTHON% -m ensurepip --upgrade
    if errorlevel 1 (
        echo [ERROR] pip is unavailable.
        goto failed
    )
)

%PYTHON% -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERROR] Package installation failed. Check your internet connection and the messages above.
    goto failed
)

echo.
echo [START] Starting the application...
echo Quiz page:  http://127.0.0.1:5000/
echo Admin page: http://127.0.0.1:5000/admin
echo Close this window to stop the server.
echo.
%PYTHON% app.py
if errorlevel 1 (
    echo.
    echo [ERROR] The application failed to start. Check the messages above.
    goto failed
)
goto end

:failed
echo.
echo Press any key to close this window...
pause >nul
:end
endlocal
