@echo off
cd /d "%~dp0"

venv\Scripts\python.exe -m PyInstaller --clean --noconfirm --windowed main.py

git add .
git diff --cached --quiet
if %errorlevel%==0 (
    echo Нет изменений для коммита.
) else (
    git commit -m "update build"
    git push
)

pause