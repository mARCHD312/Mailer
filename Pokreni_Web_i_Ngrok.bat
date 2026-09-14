@echo off
title Zoho Mailer Web Start
echo ===================================================
echo Pokrecem Zoho Web Aplikaciju...
echo ===================================================
start "Zoho Web Server" cmd /k "python app.py"

echo.
echo ===================================================
echo Pokrecem Ngrok (Povezivanje na internet)...
echo ===================================================
start "Ngrok Link" cmd /k "ngrok http 5000"
