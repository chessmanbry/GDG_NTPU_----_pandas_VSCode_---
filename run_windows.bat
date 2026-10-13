@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo [INFO] Dependencies not installed. Running setup_windows.bat first ...
    call "%~dp0setup_windows.bat"
    if errorlevel 1 goto failed
)
".venv\Scripts\python.exe" run_audit.py
if errorlevel 1 goto failed
".venv\Scripts\python.exe" verify_results.py
if errorlevel 1 goto failed
echo [PASS] Results are saved in outputs\
pause
exit /b 0
:failed
echo [FAIL] Please read the error message above.
pause
exit /b 1
