@echo off
chcp 65001 >nul
title PHAT SONG FOME WEB RA TOAN CAU (INTERNET)
cd /d "%~dp0"

echo ======================================================================
echo    HỆ THỐNG TRUY CẬP VÀ PHÁT SÓNG FOME WEB ONLINE (INTERNET)
echo ======================================================================
echo.
echo  [1] MỞ LINK ONLINE VĨNH VIỄN (RENDER 24/7 - KHUYÊN DÙNG)
echo      👉 https://fome-rp-rfi-block-b.onrender.com/
echo.
echo  [2] ĐỒNG BỘ DỮ LIỆU NGUỒN MỚI LÊN LINK ONLINE RENDER (1-CLICK)
echo.
echo  [3] PHÁT SÓNG TRỰC TIẾP MÁY NÀY QUA CLOUDFLARE TUNNEL (ĐƯỜNG TRUYỀN TỨC THỜI)
echo.
echo ======================================================================
set /p opt="Vui lòng chọn thao tác (Nhấn 1, 2 hoặc 3, mặc định là 1): "

if "%opt%"=="2" (
    call "%~dp0Dong_Bo_Du_Lieu_Len_Online.bat"
    exit /b
)

if "%opt%"=="" set opt=1
if "%opt%"=="1" (
    powershell -NoProfile -Command "Set-Clipboard -Value 'https://fome-rp-rfi-block-b.onrender.com/'"
    echo.
    echo  Đã copy link: https://fome-rp-rfi-block-b.onrender.com/
    echo  Đang mở trình duyệt...
    start https://fome-rp-rfi-block-b.onrender.com/
    pause
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "
Write-Host '[1/2] Kiem tra may chu Fome Web cuc bo...' -ForegroundColor Cyan

$portActive = $false
try {
    $tcp = New-Object System.Net.Sockets.TcpClient
    $tcp.Connect('127.0.0.1', 8765)
    $tcp.Close()
    $portActive = $true
} catch {
    $portActive = $false
}

if (-not $portActive) {
    Write-Host '  -> Dang khoi dong may chu fome_web.py...' -ForegroundColor Yellow
    $pyProc = Start-Process -FilePath 'python' -ArgumentList 'fome_web.py' -WorkingDirectory '%~dp0' -WindowStyle Hidden -PassThru
    Start-Sleep -Seconds 2
} else {
    Write-Host '  -> May chu fome_web.py da san sang.' -ForegroundColor Green
}

Write-Host '[2/2] Dang ket noi mang Cloudflare Internet toan cau...' -ForegroundColor Cyan

$logFile = Join-Path $env:TEMP 'fome_tunnel.log'
if (Test-Path $logFile) { Remove-Item $logFile -Force -ErrorAction SilentlyContinue }

Get-Process cloudflared -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$cloudflared = Join-Path '%~dp0' 'cloudflared.exe'
$proc = Start-Process -FilePath $cloudflared -ArgumentList 'tunnel --url http://127.0.0.1:8765 --no-autoupdate' -RedirectStandardError $logFile -PassThru

$url = ''
for ($i=0; $i -lt 40; $i++) {
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
    Set-Clipboard -Value $url -ErrorAction SilentlyContinue
    Clear-Host
    Write-Host '======================================================================' -ForegroundColor Green
    Write-Host '   CHUC MUNG! UNG DUNG DA DUOC DUA LEN INTERNET THANH CONG!' -ForegroundColor Green
    Write-Host '======================================================================' -ForegroundColor Green
    Write-Host ''
    Write-Host '  DUONG LINK ONLINE CLOUDFLARE (TRUY CAP TRUC TIEP MAY NAY):' -ForegroundColor Yellow
    Write-Host ('  ' + $url) -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  DUONG LINK CO DINH VINH VIEN 24/7:' -ForegroundColor Yellow
    Write-Host '  https://fome-rp-rfi-block-b.onrender.com/' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  (Da tu dong COPY link Cloudflare vao bo nho tam Clipboard cua ban)' -ForegroundColor Gray
    Write-Host '  * Ban co the paste (Ctrl+V) link nay gui qua Zalo cho dong nghiep.' -ForegroundColor White
    Write-Host '  * Mo tren dien thoai, may tinh bang hay may tinh bat ky deu dung duoc.' -ForegroundColor White
    Write-Host ''
    Write-Host '  LUU Y: GIU NGUYEN CUA SO NAY DE DUY TRI DUONG TRUYEN INTERNET.' -ForegroundColor Yellow
    Write-Host '======================================================================' -ForegroundColor Green
    Start-Process $url
} else {
    Write-Host 'Chua lay duoc link tu dong. Hay dam bao may tinh co ket noi mang Internet.' -ForegroundColor Red
}

Wait-Process -Id $proc.Id
"
pause
