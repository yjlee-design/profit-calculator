@echo off
chcp 65001 > nul
cd /d "%~dp0"
title Profit Calculator - Web
python launch_web.py
echo.
pause
