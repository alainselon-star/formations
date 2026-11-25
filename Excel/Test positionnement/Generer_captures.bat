@echo off
rem Genere toutes les captures du test de positionnement Excel
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Generer_captures.ps1"
pause
