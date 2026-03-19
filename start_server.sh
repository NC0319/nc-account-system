#!/bin/bash
# NC台账管理系统启动脚本

cd "$(dirname "$0")"

# 检查是否已在运行
if lsof -i :5001 > /dev/null 2>&1; then
    echo "服务已在运行中"
    exit 0
fi

# 启动Flask服务
nohup python3 app.py > server.log 2>&1 &

# 等待服务启动
sleep 2

# 启动外网穿透
pkill -f "serveo.net" 2>/dev/null
nohup ssh -o StrictHostKeyChecking=no -R 80:localhost:5001 serveo.net > tunnel.log 2>&1 &

sleep 3

# 获取外网地址
grep -o 'https://[^.]*\.serveousercontent\.com' tunnel.log | head -1
