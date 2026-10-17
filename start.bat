@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ========================================
echo       随机题库答题系统启动检查
echo ========================================
echo.

where py >nul 2>nul
if not errorlevel 1 (
    set "PYTHON=py"
) else (
    where python >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON=python"
    ) else (
        echo [错误] 未检测到 Python。
        echo 请安装 Python 3.10 或更新版本，并在安装时勾选 Add Python to PATH。
        echo 安装完成后重新运行本脚本。
        goto :failed
    )
)

%PYTHON% --version
if errorlevel 1 (
    echo [错误] Python 无法正常启动。
    goto :failed
)

if not exist "app.py" (
    echo [错误] 未找到 app.py，请确认启动脚本位于项目根目录。
    goto :failed
)

if not exist "requirements.txt" (
    echo [错误] 未找到 requirements.txt，无法检查项目依赖。
    goto :failed
)

echo.
echo [检查] 正在安装或检查项目依赖...
%PYTHON% -m pip --version >nul 2>nul
if errorlevel 1 (
    echo [检查] 正在启用 pip...
    %PYTHON% -m ensurepip --upgrade
    if errorlevel 1 (
        echo [错误] pip 不可用，无法安装项目依赖。
        goto :failed
    )
)

%PYTHON% -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [错误] 依赖安装失败。请检查网络连接、Python/pip 安装及上方错误信息。
    goto :failed
)

echo.
echo [启动] 正在启动答题系统...
echo 服务地址：http://127.0.0.1:5000/
echo 管理页面：http://127.0.0.1:5000/admin
echo 关闭此窗口即可停止服务。
echo.
%PYTHON% app.py
if errorlevel 1 (
    echo.
    echo [错误] 系统启动失败，请检查上方错误信息。
    goto :failed
)

goto :end

:failed
echo.
echo 按任意键关闭窗口...
pause >nul
:end
endlocal
