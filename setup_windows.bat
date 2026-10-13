@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
echo =====================================================
echo GDG NTPU pandas environment installer - Windows
echo =====================================================
set "PYTHON_LAUNCHER="
where py >nul 2>nul
if not errorlevel 1 (
    py -3 -c "import sys;sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
    if not errorlevel 1 set "PYTHON_LAUNCHER=py -3"
)
if not defined PYTHON_LAUNCHER (
    python -c "import sys;sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
    if not errorlevel 1 set "PYTHON_LAUNCHER=python"
)
if not defined PYTHON_LAUNCHER (
    echo [ERROR] Python 3.10 or later not detected.
    echo Install Python from https://www.python.org/downloads/ and enable Add Python to PATH.
    goto failed
)
echo [1/3] Creating local virtual environment .venv ...
%PYTHON_LAUNCHER% -m venv .venv
if errorlevel 1 goto failed
echo [2/3] Updating pip ...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto failed
echo [3/3] Installing pandas and openpyxl ...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -c "import pandas,openpyxl;print('READY: pandas',pandas.__version__,'openpyxl',openpyxl.__version__)"
if errorlevel 1 goto failed
echo [PASS] Setup completed. Next: double click run_windows.bat
pause
exit /b 0
:failed
echo [FAIL] Setup did not finish. Check Python and internet access.
pause
exit /b 1
