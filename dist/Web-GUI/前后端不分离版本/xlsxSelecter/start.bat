@echo off
setlocal

echo 正在检查 Python 环境和依赖...

:: 将命令行编码设置为 UTF-8
chcp 65001 > nul
if %errorlevel% neq 0 (
    echo.
    echo 错误：无法将命令行编码设置为 UTF-8。
    echo 请以管理员身份运行此脚本。
    pause
    exit /b 1
)

:: 检查 Python 是否安装
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo.
    echo 错误：未找到 Python。
    echo 请前往 https://www.python.org/ 下载并安装。
    echo 确保在安装时勾选 "Add Python to PATH"。
    echo.
    pause
    exit /b 1
)

:: 尝试导入并检查依赖
python -c "import flask, pandas, numpy, openpyxl, fuzzywuzzy" >nul 2>nul
if %errorlevel% neq 0 (
    echo.
    echo 缺少必要的依赖库，正在自动安装...
    echo.
    
    python -m pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo.
        echo 错误：依赖安装失败。
        echo 请检查网络连接或手动运行 "pip install -r requirements.txt"。
        echo.
        pause
        exit /b 1
    )
    echo.
    echo 依赖安装成功。
) else (
    echo 所有依赖已安装。
)

echo.
echo 正在启动应用...
echo 请在浏览器中访问 http://127.0.0.1:5000

:: 启动应用
python app.py

endlocal
pause