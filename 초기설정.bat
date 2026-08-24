@echo off
chcp 65001 > nul
cd /d "%~dp0"
title Profit Calculator - Setup
python setup_config.py
pause
