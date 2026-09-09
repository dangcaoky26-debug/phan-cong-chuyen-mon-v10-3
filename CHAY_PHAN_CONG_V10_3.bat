@echo off
chcp 65001 >nul
python "phan_cong_v10_3_dragdrop.py"
if errorlevel 1 pause
