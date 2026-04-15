@echo off
cd /d "%~dp0"
venv\Scripts\python.exe -m PyInstaller --clean --noconfirm --windowed main.py
pause