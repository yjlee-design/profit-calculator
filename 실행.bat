@echo off
chcp 65001 > nul
cd /d "%~dp0"
title Profit Calculator - Batch
python calc.py
pause
