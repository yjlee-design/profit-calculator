@echo off
chcp 65001 > nul
cd /d "%~dp0"
title Sync to Supabase
python sync_to_db.py
