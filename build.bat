@echo off
rem Builds dist\windows_fixer.exe (the same way the GitHub workflow does)
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    py -3 -m venv .venv || python -m venv .venv
)
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt pyinstaller || (pause & exit /b 1)
".venv\Scripts\pyinstaller.exe" --noconfirm --onefile --noconsole --name windows_fixer ^
  --icon icon.ico --collect-data sv_ttk --add-data "icon.ico;." --add-data "kuwait.png;." --add-data "Success.wav;success.wav" ^
  windows_fixer.py || (pause & exit /b 1)
echo.
echo Built: dist\windows_fixer.exe
pause
