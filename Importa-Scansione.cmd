@echo off
rem Importa nel censimento l'export di Advanced IP Scanner (CSV o XML).
rem Puoi anche trascinare il file della scansione sopra questo file.
cd /d "%~dp0"
if not exist "%~dp0Importa-Scansione.ps1" (
  echo.
  echo  Questo file e' stato aperto direttamente dentro lo zip, senza estrarlo.
  echo  Chiudi questa finestra, fai tasto destro sul file .zip ^> Estrai tutto,
  echo  poi apri la cartella estratta e rilancia Importa-Scansione.cmd da li'.
  echo.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Importa-Scansione.ps1" %1
pause
