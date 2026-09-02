#!/data/data/com.termux/files/usr/bin/bash

sleep 5

LOG=~/storage/shared/Download/server.log

# 防止重复启动
pkill -f "python.*server.py" 2>/dev/null

cd /data/data/com.termux/files/home/py

nohup python server.py >> "$LOG" 2>&1 &
