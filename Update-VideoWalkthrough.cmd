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
if not exist "%~dp0Deploy-VideoWalkthrough.ps1" (
  echo Keep Deploy-VideoWalkthrough.ps1 and VideoWalkthrough.Common.ps1 with this updater.
  pause
  exit /b 1
)
if not exist "%~dp0VideoWalkthrough.Common.ps1" (
  echo Missing VideoWalkthrough.Common.ps1. Extract the entire package first.
  pause
  exit /b 1
)
"%PS_EXE%" -NoProfile -File "%~dp0Update-VideoWalkthrough.ps1"
set "RESULT=%ERRORLEVEL%"
echo.
echo The content updater exited with code %RESULT%.
echo Open QUICKSTART.html; see update.csv and pending-update.json in the WorkFolder.
pause
popd
exit /b %RESULT%
