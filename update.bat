@echo off
cd /d "%~dp0"
echo Memperbarui data dashboard OH...
python update_data.py
if "%1"=="" pause
