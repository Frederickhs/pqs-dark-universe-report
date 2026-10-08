@echo off
setlocal

cd /d "%~dp0"

title PQS Dark Universe Report

echo ============================================================
echo PQS DARK UNIVERSE REPORT
echo ============================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo ERROR: Python virtual environment was not found.
    echo Expected: %CD%\.venv\Scripts\python.exe
    echo.
    pause
    exit /b 1
)

echo [1/2] Checking required inputs...
".venv\Scripts\python.exe" "Scripts\input_scanner.py"
if errorlevel 1 (
    echo.
    echo REPORT STOPPED: One or more required inputs are missing.
    echo Review the messages above, correct the files, and run again.
    echo.
    pause
    exit /b 1
)

echo.
echo [2/2] Generating Excel report...
".venv\Scripts\python.exe" "Scripts\report_generator.py"
if errorlevel 1 (
    echo.
    echo REPORT FAILED: The Excel report was not generated successfully.
    echo Review the messages above.
    echo.
    pause
    exit /b 1
)

echo.
echo REPORT COMPLETED SUCCESSFULLY.
echo Opening the Output folder...
start "" "%CD%\Output"

endlocal
exit /b 0
