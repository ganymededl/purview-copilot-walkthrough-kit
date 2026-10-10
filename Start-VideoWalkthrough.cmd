@echo off
setlocal
set "VW_ROOT=%~dp0"
pushd "%~dp0"
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if defined PROCESSOR_ARCHITEW6432 if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS_EXE=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS_EXE%" (
  echo Built-in Windows PowerShell was not found. Use an approved Windows host.
  pause
  exit /b 1
)
"%PS_EXE%" -NoProfile -File "%~dp0Deploy-VideoWalkthrough.ps1"
set "RESULT=%ERRORLEVEL%"
echo.
echo The walkthrough wizard exited with code %RESULT%.
echo See README.md and any deployment logs in the selected work folder.
pause
popd
exit /b %RESULT%
