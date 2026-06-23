# Windows-.exe bauen (einmalig, durch dich)

Ziel: eine **einzige `ElternTermine.exe`**, die du den Kollegen gibst. Sie brauchen dann
**nichts** zu installieren — nur doppelklicken und einmal per Code bei Microsoft anmelden.

Das Bauen machst **du einmal** auf einem **Windows-PC** (auf dem Schul-PC, wo du lokaler
Admin bist, geht das). Danach ist die .exe fertig und beliebig oft verteilbar.

## Schritte (Windows, ~5 Minuten)

1. **Python installieren** (nur auf dem Bau-PC): https://www.python.org/downloads/
   → beim Setup **„Add python.exe to PATH"** anhaken.
2. Den Ordner `ElternTermineWeb` auf den Windows-PC kopieren.
3. In dem Ordner die Datei **`build_windows.bat` doppelklicken**.
   (Sie installiert PyInstaller und baut die .exe.)
4. Fertig: die Datei liegt dann unter **`dist\ElternTermine.exe`**.

Falls du es lieber von Hand machst — in der Eingabeaufforderung (cmd) im Ordner:

```
pip install pyinstaller certifi
pyinstaller --onefile --name ElternTermine --add-data "graph.html;." --add-data "index.html;." --add-data "schullogo.png;." --hidden-import certifi server.py
```

## Den Kollegen geben

- Nur **`ElternTermine.exe`** weitergeben (z. B. per Mail/USB/Teams).
- Der Kollege: **Doppelklick** → ein schwarzes Fenster + Browser öffnen sich →
  im Browser **„Mit Microsoft anmelden"** → Code eingeben, anmelden, zustimmen → loslegen.
- Beim ersten Start zeigt Windows evtl. **„Der Computer wurde geschützt"**
  (unbekannter Herausgeber, weil die .exe nicht teuer signiert ist) →
  **„Weitere Informationen" → „Trotzdem ausführen"**. Das ist **keine** Installation und
  **keine** Admin-Freigabe, nur ein einmaliger Klick.
- Das schwarze Fenster offen lassen, solange die App genutzt wird; Schließen beendet sie.

## Wichtig / ehrlich

- Manche Schul-PCs sind sehr streng gesperrt und blockieren das Ausführen **jeder**
  unbekannten .exe (Geräte-Richtlinie). Falls die .exe gar nicht startet, liegt es daran —
  dann am besten **vorher an einem echten Kollegen-PC testen**, bevor du sie breit verteilst.
- Jeder Kollege meldet sich mit **seinem eigenen** Konto an und sendet aus seinem eigenen
  Postfach. Der Login-Schlüssel (`graph_token.json`) bleibt lokal neben der .exe.
