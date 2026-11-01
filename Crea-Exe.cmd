@echo off
rem Crea ToolkitHelpDesk.exe (il programma da mettere sulla chiavetta) da questa cartella.
if not exist "%~dp0exe\Crea-Exe.ps1" (
  echo.
  echo  Questo file e' stato aperto direttamente dentro lo zip.
  echo  Prima estrai lo zip: tasto destro sul file .zip ^> Estrai tutto, poi riprova dalla cartella estratta.
  echo.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0exe\Crea-Exe.ps1"
echo.
pause
