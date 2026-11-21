@echo off
cd /d "%~dp0"
echo === Renaming the MBR dumps ===
python rename_dumps.py --source incoming\mbr --mbr
if errorlevel 1 goto failed
echo.
echo === Filling the MBR reports ===
python run_reports.py --config config_mbr.json --group mbr
if errorlevel 1 goto failed
echo.
echo Done. MBR reports are ready.
pause
exit /b 0
:failed
echo.
echo SOMETHING WENT WRONG - read the messages above.
pause
exit /b 1
