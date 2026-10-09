@echo off
chcp 65001 >nul
setlocal

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM 自动探测 Python 解释器（py / python / python3 均可）
set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY (where python >nul 2>nul && set "PY=python")
if not defined PY (where python3 >nul 2>nul && set "PY=python3")
if not defined PY (
    echo 未检测到 Python，请先到 https://www.python.org 安装，并务必勾选 "Add Python to PATH"。
    pause
    exit /b 1
)

if not exist config.json (
    echo 未找到 config.json，请先复制 config.example.json 为 config.json 并填入账号密码。
    pause
    exit /b 1
)

echo 正在从 gyl.zyred.com 抓取正单/返单数据...
"%PY%" "%SCRIPT_DIR%scrape.py" --config "%SCRIPT_DIR%config.json"

if %errorlevel% neq 0 (
    echo 抓取失败，请检查 config.json 中的账号密码或网络连接。
    pause
    exit /b %errorlevel%
)

echo.
echo 抓取完成，合并结果已生成。
pause
endlocal
