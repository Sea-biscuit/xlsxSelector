@echo off
setlocal

chcp 65001 > nul

echo 正在检查和安装后端依赖...
python -m pip install -r ..\requirements.txt

echo 正在启动后端服务，监听端口 5000...
python app.py