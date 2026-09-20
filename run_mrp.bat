@echo off
title KMTECH MRP System
cd /d "%~dp0"

echo ========================================================
echo   KMTECH MRP Real-Time System Starting...
echo   URL: http://127.0.0.1:8000
echo ========================================================
echo.

python server.py
pause
