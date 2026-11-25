@echo off
rem Regenere une ou plusieurs captures. Exemple : Capturer.bat R3   ou   Capturer.bat Q8,N21
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Generer_captures.ps1" -Ids "%~1"
pause
