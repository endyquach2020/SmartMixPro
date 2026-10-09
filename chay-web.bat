@echo off
chcp 65001 > nul
echo ===================================================
echo     DANG KHOI DONG SMARTMIX PRO LOCAL SERVER
echo ===================================================
echo Server dang chay tai: http://localhost:8000
echo Nhan Ctrl+C de dung server.
echo.

start http://localhost:8000
python -m http.server 8000
