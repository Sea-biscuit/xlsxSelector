@echo off
setlocal

chcp 65001 > nul

echo 正在启动前端服务器，监听端口 8000...
echo 请在浏览器中访问 http://127.0.0.1:8000
python -m http.server 8000

pause