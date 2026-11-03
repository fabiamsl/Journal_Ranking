# Journal-Suche: ABDC und VHB nebeneinander

Zeitschrift eintippen und sofort sehen, welches Rating sie hat:

- **ABDC Journal Quality List 2025** (A\*, A, B, C)
- **VHB-Rating 2024**, alle 18 Fachbereiche (A+, A, B, C, D)

## Starten in drei Schritten

1. Auf GitHub oben rechts auf **Code → Download ZIP** klicken.
2. Die ZIP-Datei entpacken: Rechtsklick → „Alle extrahieren“. Direkt aus der ZIP heraus funktioniert es nicht.
3. Im entpackten Ordner **`Journal-Suche.html` doppelklicken**. Die Suche öffnet sich im Browser.

Mehr braucht es nicht: kein Python, keine Installation, kein Konto. Die Suche funktioniert auch offline.

## Bedienung

- **Suchen** nach Titel, Teil des Titels, Abkürzung (`JFE`, `JFQA`) oder ISSN (`1540-6261`). Tippfehler werden meist trotzdem gefunden.
- **VHB-Bereich** oben wählen, zum Beispiel Marketing oder Finanzierung. Der VHB bewertet je Fachbereich getrennt, dieselbe Zeitschrift kann in einem Bereich A+ und in einem anderen A haben. Die Auswahl merkt sich der Browser.
- **„Alle VHB-Bereiche“** unter einem Treffer zeigt die Bewertung in jedem Bereich.
- **„Kopieren“** legt eine Zeile wie `The Journal of Finance — ABDC 2025: A*; VHB 2024 (BAFI): A+` in die Zwischenablage, praktisch für die eigene Literaturliste.
- **„Gelistet in“** filtert: alle, nur in beiden Listen, nur ABDC, nur VHB.
- **Ohne Suchbegriff:** „Alle Zeitschriften in … anzeigen“ listet den gewählten Bereich nach Rating sortiert.
- Taste `/` springt in die Suche, `Esc` leert sie.

## Was im Ordner liegt

```text
Journal-Suche.html        ← doppelklicken, das ist die Suche
README.md                 diese Anleitung
data/
  journal_ratings.csv     alle Zeitschriften als Tabelle (öffnet in Excel)
  journal_ratings.json    dieselben Daten für Programme
  vhb_entries.csv         jeder einzelne VHB-Eintrag, zum Nachprüfen
  build_hinweise.txt      Auffälligkeiten beim Zusammenführen der Listen
python/                   nur für Fortgeschrittene, siehe unten
```

---

## Für Fortgeschrittene

### Suche auf der Kommandozeile

Python 3.9 oder neuer, nur die Standardbibliothek. Im Ordner ausführen:

```bash
python python/journal_search.py "journal of finance"
python python/journal_search.py jfe jfqa 1540-6261                 # mehrere Suchen
python python/journal_search.py "management science" --bereich STRAT --alle
python python/journal_search.py --datei meine_journals.txt         # eine Zeitschrift je Zeile
python python/journal_search.py "review of finance" --json         # maschinenlesbar
python python/journal_search.py --bereiche                         # Bereichskürzel
```

Aus Python:

```python
import sys; sys.path.insert(0, "python")
from journal_search import lookup, search
j = lookup("Journal of Finance")          # bester Treffer oder None
j["abdc"], j["vhb"]["BAFI"]["rating"]      # 'A*', 'A+'
```

### Daten aktualisieren (neue ABDC- oder VHB-Liste)

1. Originallisten herunterladen und in den Ordner `quellen/` legen (wird nicht mit hochgeladen):
   - ABDC: die Excel-Datei der Journal Quality List von abdc.edu.au als `quellen/ABDC-JQL-….xlsx`
   - VHB: die PDFs der Teilratings von vhbonline.org („VHB-Rating 2024“) als `quellen/VHB/VHB_Rating_2024_Area_rating_<KÜRZEL>.pdf`
2. `python python/build_ratings.py` ausführen. Die PDFs werden dabei mit `pdftotext -layout` in Text umgewandelt (Poppler; unter Windows bei Git for Windows dabei).
3. `data/build_hinweise.txt` lesen. `Journal-Suche.html` und alle Dateien in `data/` sind danach neu erzeugt.

Für eine neue ABDC-Ausgabe muss in `python/build_ratings.py` gegebenenfalls der Blattname (`"2025 JQL"`) angepasst werden. Aussehen und Suchlogik der Seite stehen in `python/template.html`.

### Wie die Listen zusammengeführt werden

- **VHB-PDFs lesen:** Die Tabellen behalten im Text ihre Spaltenpositionen. Lange Titel sind umgebrochen: Ein Teil steht über, ein Teil unter der Zeile mit ISSN, Rating und Stimmen. Das Skript fügt diese Teile anhand der Spalten, offener Klammern und von Zeilenenden wie „and“, „of“ oder „für“ wieder zusammen. Die Gruppenspalte der Wirtschaftsinformatik („Bezug zu WI“) wird abgetrennt.
- **Gleiche Zeitschrift:** VHB-Einträge aus verschiedenen Bereichen werden über ISSN und Titel zusammengefasst. Jede VHB-Zeitschrift wird höchstens einer ABDC-Zeitschrift zugeordnet, vorrangig über den Titel und, wenn kein Titel passt, über die ISSN. Bei sehr ähnlichen Titeln (Tippfehler in einer Liste) wird ab einer Ähnlichkeit von 0,96 verknüpft.
- Großschreibung, Sonderzeichen, „The“ am Anfang sowie „&“ und „and“ werden beim Vergleich ignoriert.

### Bekannte Grenzen

- Die Daten sind automatisch aus den PDFs gelesen. Alle 2.330 VHB-Einträge wurden gegen die Zeilen der Listen geprüft. Im Zweifel gilt trotzdem die Originalliste.
- **Rechnungswesen (RECH)** bewertet zusätzlich den Beitrag zur Praxis (Sterne). Zwölf Zeitschriften haben dort nur ein Praxis-Rating. Sie erscheinen als „nur Praxis“.
- **Steuerlehre (STEU):** Enthalten ist die Liste mit dem Kriterium „empirisch/analytisch“.
- **Wirtschaftsinformatik (WI)** enthält auch Konferenzbände. Diese sind als „Konferenz“ markiert.
- **Widersprüche zwischen den Listen** stehen in `data/build_hinweise.txt`. Beispiel: Der VHB nennt für *Review of Financial Economics* eine ISSN, die laut ABDC zu *Review of Corporate Finance* gehört. Hier gilt der Titel.

## Quellen und Haftung

Die Ratings stammen vom Australian Business Deans Council (ABDC) und vom Verband der Hochschullehrerinnen und Hochschullehrer für Betriebswirtschaft (VHB). Beide stellen ihre Listen öffentlich zum Download bereit. Dieses Tool ist kein Angebot dieser Organisationen und enthält die Originaldateien nicht. Alle Angaben ohne Gewähr.
