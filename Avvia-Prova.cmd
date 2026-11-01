@echo off
rem Prova del toolkit da eseguire su una MACCHINA DI PROVA prima di usarlo dai clienti.
rem Non modifica il PC: scrive solo in %TEMP%\HelpDeskToolkit-Prova.
cd /d "%~dp0"
if not exist "%~dp0Prova-Toolkit.ps1" (
  echo.
  echo  Questo file e' stato aperto direttamente dentro lo zip, senza estrarlo.
  echo  Chiudi questa finestra, fai tasto destro sul file .zip ^> Estrai tutto,
  echo  poi apri la cartella estratta e rilancia Avvia-Prova.cmd da li'.
  echo.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Prova-Toolkit.ps1"
pause
