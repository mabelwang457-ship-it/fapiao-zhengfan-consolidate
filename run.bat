@echo off
chcp 65001 >nul
setlocal
REM ============================================================
REM  农夫开票明细 · 正单/返单 按品类合计 —— 一键运行
REM  用法：把要处理的 xlsx 文件直接拖到本 bat 上即可
REM ============================================================

set "SCRIPT=%~dp0consolidate.py"

REM 自动探测 Python 解释器（py / python / python3 均可）
set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY (where python >nul 2>nul && set "PY=python")
if not defined PY (where python3 >nul 2>nul && set "PY=python3")
if not defined PY (
    echo 未检测到 Python，请先安装并勾选 "Add Python to PATH"。
    pause
    exit /b 1
)

if "%~1"=="" (
    echo 用法：将需要处理的 Excel（含「正单」「返单」两个 sheet）拖到本文件上。
    echo 也可以在该文件夹打开命令行，执行：
    echo     python consolidate.py 你的文件.xlsx
    pause
    exit /b
)

"%PY%" "%SCRIPT%" "%~1"
if errorlevel 1 (
    echo.
    echo [错误] 运行失败。请确认已安装 Python 并勾选 PATH；所需依赖会在运行时自动安装。
    echo.
    pause
)
