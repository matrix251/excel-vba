@echo off
REM Define registry path for Excel Trust Center Settings
set regPath="HKCU:\Software\Microsoft\Office\16.0\Excel\Security"

REM Set VBAWarnings key (1 = enable all, 4 = disable without notification)
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Set-ItemProperty -Path %regPath% -Name 'VBAWarnings' -Value 1 -ErrorAction SilentlyContinue"

REM Close Excel to apply the changes
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue"

REM Restart Excel to apply the changes
start excel.exe
