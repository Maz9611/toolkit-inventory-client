@echo off
rem Censisce QUESTO PC e lo scrive nell'Excel del censimento (cartella Censimento della chiavetta).
rem Non modifica il PC: legge i dati e scrive solo sulla chiavetta.
cd /d "%~dp0"
if not exist "%~dp0Censisci-PC.ps1" (
  echo.
  echo  Questo file e' stato aperto direttamente dentro lo zip, senza estrarlo.
  echo  Chiudi questa finestra, fai tasto destro sul file .zip ^> Estrai tutto,
  echo  poi apri la cartella estratta e rilancia Censisci-PC.cmd da li'.
  echo.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Censisci-PC.ps1"
pause
