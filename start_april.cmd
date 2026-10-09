@echo off
chcp 65001 >nul
cd /d "%~dp0"
python webui.py --input "data\4月份排产.xlsx"
pause
