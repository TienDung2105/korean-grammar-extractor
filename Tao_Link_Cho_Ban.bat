@echo off
chcp 65001 >nul
title Chia sẻ web ra Internet (Cloudflare Tunnel)
echo ============================================================
echo   ĐANG TẠO ĐƯỜNG LINK ONLINE CHO BẠN CỦA BẠN TRUY CẬP...
echo ============================================================
echo.
echo [CÁCH SỬ DỤNG]:
echo 1. Nhìn vào các dòng chữ bên dưới, tìm dòng có link dạng:
echo.
echo       https://xxxxxxxx.trycloudflare.com
echo.
echo 2. Copy link đó gửi qua Zalo / Facebook cho bạn của bạn.
echo 3. Bạn của bạn bấm vào link trên điện thoại hoặc máy tính là dùng được!
echo.
echo * Lưu ý: Giữ cửa sổ này mở trong lúc bạn của bạn đang dùng. Khi bạn đóng cửa sổ này thì link sẽ tự tắt.
echo ============================================================
echo.

cloudflared.exe tunnel --url http://localhost:5000
pause

