@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo 智能论文整理系统 - 本地服务器启动
echo ========================================
echo.
echo 当前目录: %CD%
echo 正在启动 Python HTTP 服务器...
echo 服务器地址: http://localhost:8000
echo.
echo 按 Ctrl+C 可以停止服务器
echo ========================================
echo.

python -m http.server 8000

pause
