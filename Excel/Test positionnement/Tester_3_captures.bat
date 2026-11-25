@echo off
rem Test sur 3 questions : Q8 (grille), Q19 (graphique), R4 (TCD)
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Generer_captures.ps1" -Ids "Q8,Q19,R4"
pause
