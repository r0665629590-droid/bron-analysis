@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM Запуск без консольного вікна
start "" pythonw.exe "%~dp0bron_analysis_all.py"

REM Якщо pythonw не знайдено — пробуємо python
if %errorlevel% neq 0 (
    python "%~dp0bron_analysis_all.py"
    if %errorlevel% neq 0 (
        echo Помилка запуску. Спочатку запустіть install.bat
        pause
    )
)
