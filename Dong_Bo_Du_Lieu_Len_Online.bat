@echo off
chcp 65001 >nul
title DONG BO DU LIEU LEN ONLINE (RENDER 24/7)
cd /d "%~dp0"

echo ======================================================================
echo    HỆ THỐNG ĐỒNG BỘ DỮ LIỆU TỰ ĐỘNG LÊN WEB ONLINE (RENDER 24/7)
echo    Đường link Online cố định: https://fome-rp-rfi-block-b.onrender.com/
echo ======================================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command "
Write-Host '[1/3] Kiem tra ket noi Git va du lieu thay doi...' -ForegroundColor Cyan

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host 'LOI: Khong tim thay git tren may tinh nay!' -ForegroundColor Red
    pause
    exit 1
}

Write-Host '  -> Dang kiem tra nhanh tu GitHub...' -ForegroundColor Gray
git pull --rebase origin main

Write-Host '[2/3] Dang tong hop du lieu nguon va bieu mau...' -ForegroundColor Cyan
git add Fome_Sites/
git add Fome_Web_Data/users.json
git add Fome_Web_Data/app_state.json
git add web/
git add fome_web.py
git add core_processor.py
git add nktc_processor.py
git add requirements.txt
git add render.yaml
git add .gitignore
git add *.bat

$status = git status --porcelain
if ($status) {
    $nowStr = Get-Date -Format 'dd/MM/yyyy HH:mm:ss'
    Write-Host ('  -> Phat hien du lieu moi, dang dong goi commit: ' + $nowStr) -ForegroundColor Yellow
    git commit -m ('sync: cap nhat du lieu tram ' + $nowStr)
    Write-Host '[3/3] Dang day du lieu len GitHub de Render tu dong cap nhat...' -ForegroundColor Cyan
    git push origin main
    if ($LASTEXITCODE -eq 0) {
        Write-Host '  -> [THANH CONG] Du lieu da duoc day len GitHub thanh cong!' -ForegroundColor Green
    } else {
        Write-Host '  -> [CANH BAO] Khong the push len GitHub. Vui long kiem tra ket noi mang.' -ForegroundColor Red
    }
} else {
    Write-Host '  -> Du lieu hien tai da dong bo hoan toan voi online.' -ForegroundColor Green
}

$url = 'https://fome-rp-rfi-block-b.onrender.com/'
Set-Clipboard -Value $url -ErrorAction SilentlyContinue

Write-Host ''
Write-Host '======================================================================' -ForegroundColor Green
Write-Host '   HE THONG ONLINE DA SAN SANG SU DUNG 24/7!' -ForegroundColor Green
Write-Host '======================================================================' -ForegroundColor Green
Write-Host ''
Write-Host '  DUONG LINK TRUY CAP CO DINH VINH VIEN:' -ForegroundColor Yellow
Write-Host ('  ' + $url) -ForegroundColor Cyan
Write-Host ''
Write-Host '  (Da tu dong COPY link nay vao Clipboard bo nho tam)' -ForegroundColor Gray
Write-Host '  * Bat ky ai co link nay deu co the dang nhap va su dung online.' -ForegroundColor White
Write-Host '  * Khong phu thuoc vao may tinh nay bat hay tat.' -ForegroundColor White
Write-Host '  * Render tu dong cap nhat ban moi nhat sau 1-2 phut.' -ForegroundColor White
Write-Host '======================================================================' -ForegroundColor Green
Write-Host ''
Write-Host 'Dang mo trinh duyet den trang web...' -ForegroundColor Cyan
Start-Process $url
"
pause
