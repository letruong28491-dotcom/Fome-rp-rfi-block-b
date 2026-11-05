@echo off
chcp 65001 >nul
title PHAT SONG FOME WEB RA TOAN CAU (INTERNET)
cd /d "%~dp0"

echo ======================================================================
echo    HỆ THỐNG PHÁT SÓNG FOME WEB RA TOÀN CẦU (INTERNET)
echo ======================================================================
echo.
echo [1/2] Đang khởi động máy chủ Fome cục bộ...
start "" "%~dp0Chay_Ngam_Khong_Cua_So.vbs"

echo [2/2] Đang kết nối mạng Cloudflare Internet toàn cầu...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command "
$logFile = Join-Path $env:TEMP 'fome_tunnel.log'
if (Test-Path $logFile) { Remove-Item $logFile -Force -ErrorAction SilentlyContinue }

$cloudflared = Join-Path '%~dp0' 'cloudflared.exe'
$proc = Start-Process -FilePath $cloudflared -ArgumentList 'tunnel --url http://127.0.0.1:8765 --no-autoupdate' -RedirectStandardError $logFile -PassThru

$url = ''
for ($i=0; $i -lt 35; $i++) {
    Start-Sleep -Milliseconds 600
    if (Test-Path $logFile) {
        $content = Get-Content $logFile -Raw -ErrorAction SilentlyContinue
        if ($content -match 'https://[a-zA-Z0-9-]+\.trycloudflare\.com') {
            $url = $matches[0]
            break
        }
    }
}

if ($url) {
    Clear-Host
    Write-Host '======================================================================' -ForegroundColor Green
    Write-Host '   CHÚC MỪNG! ỨNG DỤNG ĐÃ ĐƯỢC ĐƯA LÊN INTERNET THÀNH CÔNG!' -ForegroundColor Green
    Write-Host '======================================================================' -ForegroundColor Green
    Write-Host ''
    Write-Host '  ĐƯỜNG LINK ONLINE BẢO MẬT (TRUY CẬP TỪ MỌI NƠI):' -ForegroundColor Yellow
    Write-Host ('  ' + $url) -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  * Bạn có thể gửi link này qua Zalo/Email cho sếp hoặc đồng nghiệp.' -ForegroundColor White
    Write-Host '  * Mở trên điện thoại, máy tính bảng hay máy tính bất kỳ đều dùng được.' -ForegroundColor White
    Write-Host '  * Hoàn toàn không cần cài Python, không cần Office trên máy truy cập!' -ForegroundColor White
    Write-Host ''
    Write-Host '  LƯU Ý: GIỮ NGUYÊN CỬA SỔ NÀY ĐỂ DUY TRÌ ĐƯỜNG TRUYỀN INTERNET.' -ForegroundColor Yellow
    Write-Host '======================================================================' -ForegroundColor Green
    Start-Process $url
} else {
    Write-Host 'Chưa lấy được link tự động. Hãy đảm bảo máy tính có kết nối mạng Internet.' -ForegroundColor Red
}

Wait-Process -Id $proc.Id
"
pause
