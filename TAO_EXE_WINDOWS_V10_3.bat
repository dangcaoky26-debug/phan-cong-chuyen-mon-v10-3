@echo off
chcp 65001 >nul
where py >nul 2>nul
if errorlevel 1 (
  echo Khong tim thay Python Launcher "py". Hay cai Python 3 truoc.
  pause
  exit /b 1
)
py -m pip install --upgrade pyinstaller
py -m PyInstaller --noconfirm --clean --onefile --noconsole ^
  --name "PHAN_CONG_CHUYEN_MON_V10_3" ^
  --add-data "index_v10_3.html;." ^
  --add-data "v10_3_seed.json;." ^
  --add-data "MAU_NHAP_TKB_EXCEL_V10_3.xlsx;." ^
  "phan_cong_v10_3_dragdrop.py"
if errorlevel 1 pause
