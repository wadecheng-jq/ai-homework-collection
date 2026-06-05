@echo off
chcp 65001 > nul
echo ============================================
echo   AI作业收集系统 - 启动脚本
echo ============================================
echo.

REM 找到Python路径（优先使用WorkBuddy管理的Python）
set PYTHON_EXE=C:\Users\wade\.workbuddy\binaries\python\versions\3.13.12\python.exe
if not exist "%PYTHON_EXE%" (
    set PYTHON_EXE=python
)

echo [INFO] 使用Python：%PYTHON_EXE%
echo.

REM 进入脚本所在目录
cd /d %~dp0

REM 检查依赖
echo [INFO] 检查Python依赖...
"%PYTHON_EXE%" -c "import flask, openpyxl" 2>nul
if errorlevel 1 (
    echo [WARN] 缺少依赖，正在自动安装...
    "%PYTHON_EXE%" -m pip install -r requirements.txt -q
    if errorlevel 1 (
        echo [ERROR] 依赖安装失败，请手动执行：pip install -r requirements.txt
        pause
        exit /b 1
    )
)

echo [INFO] 依赖检查通过。
echo.
echo [INFO] 启动Web服务...
echo [INFO] 启动后访问：http://localhost:5000
echo [INFO] 管理后台：  http://localhost:5000/admin
echo [INFO] 关闭窗口即可停止服务。
echo ============================================
echo.

"%PYTHON_EXE%" app.py

pause
