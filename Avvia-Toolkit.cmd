@echo off
rem Apre la finestra del Toolkit Help Desk (censimento, scansione della rete, verifiche).
cd /d "%~dp0"
if not exist "%~dp0Interfaccia.ps1" (
  echo.
  echo  Questo file e' stato aperto direttamente dentro lo zip, senza estrarlo.
  echo  Chiudi questa finestra, fai tasto destro sul file .zip ^> Estrai tutto,
  echo  poi apri la cartella estratta e rilancia Avvia-Toolkit.cmd da li'.
  echo.
  pause
  exit /b 1
)
start "Toolkit Help Desk" /min powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Interfaccia.ps1"
