@echo off
chcp 65001 >nul
title Korean Double Negation Extractor
echo ============================================================
echo   Đang khởi động chương trình trích xuất ngữ pháp tiếng Hàn...
echo ============================================================
echo.

python -m pip install flask openpyxl --break-system-packages --quiet 2>nul

echo Đang mở giao diện trên trình duyệt...
start http://localhost:5000

python app.py
pause

