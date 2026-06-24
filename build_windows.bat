@echo off
REM ===== Crée ElternTermine.exe (une fois, sur un PC Windows) =====
cd /d "%~dp0"
echo.
echo == Installer PyInstaller (si nécessaire) ==
python -m pip install --upgrade pyinstaller certifi
echo.
echo == Créer ElternTermine.exe ==
python -m PyInstaller --onefile --name ElternTermine ^
  --add-data "graph.html;." ^
  --add-data "index.html;." ^
  --add-data "schullogo.png;." ^
  --hidden-import certifi ^
  server.py
echo.
echo ============================================================
echo  Terminé. Le fichier se trouve dans :  dist\ElternTermine.exe
echo  Vous pouvez transmettre ce .exe aux collègues.
echo ============================================================
pause
