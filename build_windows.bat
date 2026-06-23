@echo off
REM ===== Baut ElternTermine.exe (einmalig, auf einem Windows-PC) =====
cd /d "%~dp0"
echo.
echo == PyInstaller installieren (falls noetig) ==
python -m pip install --upgrade pyinstaller certifi
echo.
echo == ElternTermine.exe bauen ==
python -m PyInstaller --onefile --name ElternTermine ^
  --add-data "graph.html;." ^
  --add-data "index.html;." ^
  --add-data "schullogo.png;." ^
  --hidden-import certifi ^
  server.py
echo.
echo ============================================================
echo  Fertig. Die Datei liegt unter:  dist\ElternTermine.exe
echo  Diese .exe gibst du den Kollegen.
echo ============================================================
pause
