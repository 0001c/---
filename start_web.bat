@echo off
chcp 65001 >nul

echo ===========================================
echo   xx调剂定制班 - Web 服务启动
echo ===========================================
echo.

:: 检查虚拟环境
if exist ".venv\Scripts\activate.bat" (
    echo [检查] 找到虚拟环境，正在激活...
    call .venv\Scripts\activate.bat
) else if exist "..\.venv\Scripts\activate.bat" (
    echo [检查] 找到上级虚拟环境，正在激活...
    call ..\.venv\Scripts\activate.bat
) else (
    echo [警告] 未找到 .venv 虚拟环境，将使用系统 Python。
)

:: 检查是否安装了 flask
python -c "import flask" 2>nul
if errorlevel 1 (
    echo [安装] 正在安装 Web 依赖...
    pip install flask flask-socketio pandas openpyxl openai
    echo.
)

:: 进入 web 目录并启动服务
cd /d "%~dp0web"
echo [启动] 正在启动 Web 服务...
echo.
echo   访问地址: http://localhost:5000
echo   按 Ctrl+C 停止服务
echo.
echo ===========================================

python app.py

pause
