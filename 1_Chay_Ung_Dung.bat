@echo off
chcp 65001 >nul
title HE THONG XUAT BAO CAO RFI VA RP TU DONG
cd /d "%~dp0"

echo ======================================================================
echo    HỆ THỐNG XỬ LÝ VÀ XUẤT BÁO CÁO RFI & RP TỰ ĐỘNG
echo ======================================================================
echo.
echo [1/3] Máy chủ nội bộ: http://127.0.0.1:8765/
echo.
echo [2/3] ĐƯỜNG LINK ONLINE CỐ ĐỊNH VĨNH VIỄN (TRUY CẬP 24/7 TỪ MỌI NƠI):
echo       👉 https://fome-rp-rfi-block-b.onrender.com/
echo.
echo [3/3] ĐỒNG BỘ DỮ LIỆU LÊN ONLINE:
echo       👉 Nhấp đúp file: Dong_Bo_Du_Lieu_Len_Online.bat
echo.
echo ======================================================================
echo   LƯU Ý:
echo   - Giữ nguyên cửa sổ này khi dùng nội bộ trên máy này.
echo   - Đồng nghiệp, QC, Giám sát có thể mở trực tiếp link Render Online
echo     ở trên từ điện thoại, Zalo hay máy tính mà không cần cài phần mềm!
echo ======================================================================
echo.

start http://127.0.0.1:8765/
python fome_web.py

pause
