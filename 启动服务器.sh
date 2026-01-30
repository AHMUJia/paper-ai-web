#!/bin/bash
echo "========================================"
echo "智能论文整理系统 - 本地服务器启动"
echo "========================================"
echo ""
echo "正在启动 Python HTTP 服务器..."
echo "服务器地址: http://localhost:8000"
echo ""
echo "按 Ctrl+C 可以停止服务器"
echo "========================================"
echo ""

python3 -m http.server 8000
