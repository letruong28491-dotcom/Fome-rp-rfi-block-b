@echo off
chcp 65001 >nul
title HE THONG XUAT BAO CAO RFI VA RP TU DONG
cd /d "%~dp0"

echo ======================================================================
echo    HỆ THỐNG XỬ LÝ VÀ XUẤT BÁO CÁO RFI & RP TỰ ĐỘNG
echo ======================================================================
echo.
echo [1/2] Đang khởi động máy chủ cục bộ...
echo [2/2] Địa chỉ truy cập: http://127.0.0.1:8765/
echo.
echo ======================================================================
echo   LƯU Ý:
echo   - Giữ nguyên cửa sổ này trong khi sử dụng.
echo   - Để phát sóng ra Internet cho máy khác dùng, hãy mở file:
echo     Phat_Song_Ra_Internet.bat
echo ======================================================================
echo.

start http://127.0.0.1:8765/
python fome_web.py

pause
