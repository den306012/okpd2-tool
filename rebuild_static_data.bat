@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

python -m pip install --quiet openpyxl striprtf

echo ============================== >> rebuild_log.txt
echo %date% %time% >> rebuild_log.txt

python scripts\build_static_data.py >> rebuild_log.txt 2>&1
if errorlevel 1 (
    echo OSHIBKA pri peresborke - sm. tekst vyshe >> rebuild_log.txt
    exit /b 1
)

git add index.html >> rebuild_log.txt 2>&1
git diff --cached --quiet
if %errorlevel%==0 (
    echo Izmeneniy net, kommit ne nuzhen. >> rebuild_log.txt
) else (
    git commit -m "Obnovlenie klassifikatora/1875/719 %date%" >> rebuild_log.txt 2>&1
    git push >> rebuild_log.txt 2>&1
)

echo Gotovo. >> rebuild_log.txt
