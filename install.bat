@echo off
chcp 65001 >nul
title Встановлення — Аналіз бронювання
setlocal enabledelayedexpansion

echo ============================================================
echo  АНАЛІЗ БРОНЬЮВАННЯ — Встановлення
echo ============================================================
echo.

REM ── 1. Перевірка наявності Python ──────────────────────────
echo [1/3] Перевірка Python...
python --version >nul 2>&1
if %errorlevel% equ 0 (
    for /f "tokens=2" %%v in ('python --version 2^>^&1') do set PYVER=%%v
    echo     ✓ Python !PYVER! вже встановлено
    goto :install_deps
)

echo     ✗ Python не знайдено. Встановлюю...
echo.

REM ── 2. Встановлення Python через winget ────────────────────
where winget >nul 2>&1
if %errorlevel% equ 0 (
    echo     Використовую winget...
    winget install -e --id Python.Python.3.12 --silent --accept-source-agreements --accept-package-agreements
    if !errorlevel! neq 0 (
        echo     ✗ Помилка winget. Спробую завантажити вручну...
        goto :manual_python
    )
    echo     ✓ Python встановлено через winget
    echo.
    echo     ВАЖЛИВО: Закрийте це вікно і запустіть install.bat ще раз,
    echo     щоб система побачила оновлений PATH.
    pause
    exit /b 0
)

:manual_python
echo     Завантажую Python з python.org...
set PYINSTALL=%TEMP%\python-installer.exe
powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe' -OutFile '%PYINSTALL%'"
if not exist "%PYINSTALL%" (
    echo     ✗ Не вдалося завантажити Python.
    echo     Встановіть вручну з https://www.python.org/downloads/
    pause
    exit /b 1
)

echo     Запускаю встановлення (тиха установка + додавання в PATH)...
"%PYINSTALL%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
del "%PYINSTALL%"

echo     ✓ Python встановлено.
echo.
echo     ВАЖЛИВО: Закрийте це вікно і запустіть install.bat ще раз,
echo     щоб система побачила оновлений PATH.
pause
exit /b 0

REM ── 3. Встановлення залежностей ────────────────────────────
:install_deps
echo.
echo [2/3] Оновлення pip...
python -m pip install --upgrade pip --quiet
if %errorlevel% neq 0 (
    echo     ✗ Помилка оновлення pip
    pause
    exit /b 1
)
echo     ✓ pip оновлено
echo.

echo [3/3] Встановлення залежностей з requirements.txt...
python -m pip install -r "%~dp0requirements.txt" --quiet
if %errorlevel% neq 0 (
    echo     ✗ Помилка встановлення залежностей
    pause
    exit /b 1
)
echo     ✓ Всі залежності встановлено
echo.

REM ── 4. Створення ярлика на робочому столі ──────────────────
echo Створення ярлика на робочому столі...
set SCRIPT_PATH=%~dp0bron_analysis_all.py
set SHORTCUT=%USERPROFILE%\Desktop\Аналіз бронювання.lnk

powershell -NoProfile -Command ^
    "$s=(New-Object -ComObject WScript.Shell).CreateShortcut('%SHORTCUT%');" ^
    "$s.TargetPath='pythonw.exe';" ^
    "$s.Arguments='\"%SCRIPT_PATH%\"';" ^
    "$s.WorkingDirectory='%~dp0';" ^
    "$s.IconLocation='pythonw.exe';" ^
    "$s.Save()"

if exist "%SHORTCUT%" (
    echo     ✓ Ярлик створено на робочому столі
) else (
    echo     ⚠ Не вдалося створити ярлик (не критично)
)

echo.
echo ============================================================
echo  ✓ ВСТАНОВЛЕННЯ ЗАВЕРШЕНО!
echo ============================================================
echo.
echo  Запустіть програму:
echo    - подвійним кліком на "run.bat"
echo    - або через ярлик на робочому столі
echo.
pause
