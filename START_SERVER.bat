@echo off
cd /d "%~dp0"
title PriceInstall Web Server
echo Starting PriceInstall server on http://localhost:8095/ ...
if exist .venv\Scripts\python.exe (
    .venv\Scripts\python.exe price_server.py
) else (
    python price_server.py
)
pause
