@echo off
REM ===== Démarrer Rendez-vous parents (Windows) =====
REM Un double-clic démarre l’assistant local ; le navigateur s’ouvre automatiquement.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ElternTermine.ps1"
