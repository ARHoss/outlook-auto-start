@echo off
start "" "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" /recycle
timeout /t 5 /nobreak > nul
for /f "tokens=2" %%i in ('tasklist /fi "imagename eq OUTLOOK.EXE" /fo list /nh') do set pid=%%i
powershell -Command "(Get-Process -id %pid%).MainWindowHandle | ForEach-Object {Invoke-Command -ScriptBlock {[Windows.UI.Shell.Taskbar]::GetDefault().InvalidateThumbnailPreview($_)}}"
timeout /t 1 /nobreak > nul
taskkill /f /im outlook.exe
exit
