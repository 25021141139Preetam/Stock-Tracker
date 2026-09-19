@echo off
setlocal
title NSE Stock Price Tracker

cd /d "%~dp0"

echo =====================================================================
echo                     NSE STOCK PRICE TRACKER LAUNCHER
echo =====================================================================
echo.

set "PYTHON_EXE="

REM Test 1: Check G:\orange data miner\python.exe
if exist "G:\orange data miner\python.exe" (
    "G:\orange data miner\python.exe" --version >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_EXE=G:\orange data miner\python.exe"
        goto :FOUND
    )
)

REM Test 2: Check Codex primary runtime Python
if exist "C:\Users\ASUS\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" (
    "C:\Users\ASUS\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" --version >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_EXE=C:\Users\ASUS\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
        goto :FOUND
    )
)

REM Test 3: Check if python in PATH actually executes
python --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_EXE=python"
    goto :FOUND
)

REM Test 4: Check Windows py launcher
py -3 --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_EXE=py -3"
    goto :FOUND
)

REM Test 5: Check D:\Python3
if exist "D:\Python3\python.exe" (
    "D:\Python3\python.exe" --version >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_EXE=D:\Python3\python.exe"
        goto :FOUND
    )
)

REM Test 6: Check AppData Local Python versions
for %%V in (312 311 310 39) do (
    if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
        "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" --version >nul 2>&1
        if not errorlevel 1 (
            set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
            goto :FOUND
        )
    )
)

echo [ERROR] A working Python installation was not detected.
echo Please ensure Python 3.10+ is installed and configured in PATH.
echo.
pause
exit /b 1

:FOUND
echo Detected Python: "%PYTHON_EXE%"
echo Executing stock_tracker.py...
echo.

"%PYTHON_EXE%" stock_tracker.py
if errorlevel 1 (
    echo.
    echo [ERROR] Stock tracker encountered an error.
) else (
    echo.
    echo [SUCCESS] Stock tracker finished successfully.
)

echo.
pause
