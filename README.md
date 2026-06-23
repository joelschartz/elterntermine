# ElternTermine · Microsoft 365

Lokale App zum Planen von Elterngesprächen und zum Versenden der Bestätigungs-Mails über das eigene Microsoft-365-Konto.

## Neu in dieser Version

- Im Reiter **E-Mails & Versand** gibt es eine **vorgefertigte E-Mail-Fassung**.

- Die Standard-Signatur ist jetzt mit der Adresse von Joël Schartz vorausgefüllt.
- Der Abstand zwischen Trennlinie und Logo wurde reduziert.
- Die Logo-Auswahl ist kein Dropdown mehr, sondern robuste Schalter: **UNESCO-Logo**, **Eigenes Logo hochladen**, **Kein Logo**.
- Die **vorgefertigte E-Mail-Fassung** kann jetzt komplett eingeklappt werden, inklusive Signatur- und Logo-Bereich.
- Der E-Mail-Text ist sofort direkt bearbeitbar, ohne extra Edit-Text-Knopf.
- **Text und Signatur sind jetzt im gleichen Editor**: Die Signatur kann direkt im E-Mail-Text bearbeitet und formatiert werden.
- Oben rechts neben **Vorgefertigte E-Mail-Fassung** bleibt die Checkbox **Signature**, um die Signatur ein- oder auszuschalten.
- Das separate Signatur-Fenster wurde entfernt.
- Unter dem Editor gibt es jetzt eine klare Signatur-Leiste: **Logo in Signatur: UNESCO-Logo / Eigenes Logo / Kein Logo**.
- **Signatur zurücksetzen** stellt die Standard-Signatur wieder her und aktiviert wieder das UNESCO-Logo.
- Die Outlook-Signatur-Suche wurde entfernt.
- Die Vorschau aktualisiert sich automatisch; die früheren Buttons „Vorschau aktualisieren“ und „Testmail an mich“ wurden entfernt.
- Das Layout wurde aufgeräumt: flachere Karten, klarere Aktionsleiste und eine übersichtlichere Versandliste ohne verschachtelte Rahmen.
- Der E-Mail-Text, die Signatur und das Logo werden weiterhin in der Vorschau und beim Versand übernommen.
- Falls Port 8765 bereits belegt ist, nimmt die App automatisch einen freien Port.

## Starten auf dem Mac

1. ZIP entpacken.
2. Den Ordner **ElternTermineWeb** öffnen.
3. Doppelklick auf **Starten.command**.
4. Der Browser öffnet sich automatisch.

Falls macOS das Öffnen blockiert:

1. Rechtsklick auf **Starten.command**.
2. **Öffnen** wählen.
3. Noch einmal bestätigen.

Falls Python fehlt, im Terminal testen:

```bash
python3 --version
```

## Ablauf

1. **Schüler** eintragen oder per CSV importieren.
2. **Kalender** anlegen und Schüler den Terminen zuweisen.
3. **E-Mails & Versand** öffnen:
   - mit Microsoft anmelden,
   - direkt in den E-Mail-Text klicken und ihn wie in Outlook bearbeiten,
   - die Standard-Signatur direkt im Text bearbeiten,
   - oben rechts über **Signature** die Signatur ein- oder ausschalten,
   - unter dem Editor bei **Logo in Signatur** zwischen **UNESCO-Logo**, **Eigenes Logo hochladen** und **Kein Logo** wählen,
   - mit **Signatur zurücksetzen** die Standard-Signatur wiederherstellen,
   - die vorgefertigte Fassung bei Bedarf einklappen und einzelne E-Mails über **Bearbeiten** kontrollieren,
   - dann **Alle senden**.

## Platzhalter im E-Mail-Text

- `{DATE}` = Datum auf Französisch
- `{TIME}` = Uhrzeit
- `{NAME}` = dein Name aus dem Microsoft-Konto-Feld
- `{ENFANT}` = Name des Kindes

## Wichtig

Die App läuft lokal auf deinem Computer. Schülerdaten und Vorlagen werden lokal im Browser gespeichert. Der Versand läuft über Microsoft Graph mit deinem eigenen Microsoft-365-Konto.
