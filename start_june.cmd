@echo off
chcp 65001 >nul
cd /d "%~dp0"
python webui.py --input "data\630厂热处理车间6月排产.xlsx"
pause
