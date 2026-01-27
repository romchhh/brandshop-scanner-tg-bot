#!/bin/bash
source /root/BrandshopScannerBot/myenv/bin/activate
nohup python3 /root/BrandshopScannerBot/main.py > /dev/null 2>&1 &
