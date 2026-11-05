@echo off
chcp 65001 >nul
title HE THONG XUAT BAO CAO RFI VA RP TU DONG
cd /d "%~dp0"

echo ======================================================================
echo    HỆ THỐNG XỬ LÝ VÀ XUẤT BÁO CÁO RFI & RP TỰ ĐỘNG
echo ======================================================================
echo.
echo [1/2] Đang chuẩn bị môi trường chạy...
echo [2/2] Đang khởi động máy chủ cục bộ...
echo.
echo ======================================================================
echo   ĐỊA CHỈ TRUY CẬP: http://127.0.0.1:8765/
echo.
echo   LƯU Ý QUAN TRỌNG:
echo   - Trình duyệt web sẽ tự động mở sau khi máy chủ sẵn sàng.
echo   - GIỮ NGUYÊN CỬA SỔ NÀY TRONG KHI SỬ DỤNG.
echo   - Khi dùng xong, bạn có thể đóng cửa sổ này để tắt chương trình.
echo   - Nếu trình duyệt chưa tự mở, bạn có thể nhấp đúp file:
echo     2_Mo_Trinh_Duyet.bat
echo ======================================================================
echo.

"%~dp0FomeApp.exe"

pause
