@echo off
cd /d "%~dp0"
echo === Renaming dumps ===
python rename_dumps.py --source incoming
if errorlevel 1 goto failed
echo.
echo === Filling the daily reports ===
python run_reports.py --group daily
if errorlevel 1 goto failed
echo.
echo Done. Reports are ready.
pause
exit /b 0

:failed
echo.
echo SOMETHING WENT WRONG - read the messages above.
pause
exit /b 1
