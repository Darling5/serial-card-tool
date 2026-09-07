@echo off
rem Serial Card Tool launcher - requires Python 3.11 with tkinter/openpyxl/pyserial
set "PYW=C:\Python311\pythonw.exe"
if exist "%PYW%" (
    start "" "%PYW%" "%~dp0serial_card_tool.py"
    exit /b 0
)
where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "%~dp0serial_card_tool.py"
    exit /b 0
)
echo [ERROR] Python not found. Install Python 3.11 to C:\Python311
echo then run: C:\Python311\python.exe -m pip install pyserial
pause
