#!/bin/bash
source /root/brandshop-scanner-tg-bot/myenv/bin/activate
nohup python3 /root/brandshop-scanner-tg-bot/main.py > /dev/null 2>&1 &
