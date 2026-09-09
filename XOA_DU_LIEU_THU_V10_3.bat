@echo off
chcp 65001 >nul
echo CANH BAO: thao tac nay xoa database V10.3.
choice /M "Ban chac chan muon xoa du lieu va nap lai du lieu goc"
if errorlevel 2 exit /b 0
if exist "phan_cong_v10_3.db" del "phan_cong_v10_3.db"
echo Da xoa database V10.3.
pause
