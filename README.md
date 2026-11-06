# Praesentationsassistent (Grundgeruest)

Visual-Identity-Praesentationsassistent fuer die Hotelgruppe (Marken:
Althoff Hotels, Althoff Collection, AMERON, URBAN LOFT). Ziel: aus
Dokumenten + Marken-Regeln automatisch PowerPoint-Praesentationen erzeugen,
die verbindlich eine der vier Marken-Visual-Identities einhalten; bestehende
.pptx anhand von Kommentaren/Aenderungslisten gezielt ueberarbeiten; dabei
nie Fakten erfinden.

**Stand heute:** Grundgeruest + Langdock-Anbindung + ein erster
Marken-Abgleich zum Anschauen (`erzeugen-mit-marke`, siehe unten). Der
volle Workflow (Dokumenten-Extraktion, Dramaturgie, Marken-Checkliste,
Revision via Kommentare, Lern-Gedaechtnis) fehlt noch.

## Installation auf dem Arbeitslaptop (Testphase, ohne GitHub)

Der Code liegt noch nicht auf dem Arbeits-GitHub (kommt spaeter) -- fuers
erste Ausprobieren reicht eine einmalige Kopie des Projektordners:

1. Diesen Projektordner als .zip auf den Arbeitslaptop bringen (z.B. per
   USB-Stick, eigener Cloud-Ordner oder Mail an dich selbst) und entpacken,
   z.B. nach `C:\Praesentationsassistent`.
2. Im entpackten Ordner `windows\einrichten.ps1` per Rechtsklick ->
   "Mit PowerShell ausfuehren" starten (installiert bei Bedarf Python,
   legt eine `.venv` an, installiert das Programm, legt `praesentation.env`
   aus der Vorlage an und eine Desktop-Verknuepfung).
3. `praesentation.env` oeffnen und den eigenen Langdock-API-Key eintragen.
4. Desktop-Verknuepfung "Praesentationsassistent" starten -- zeigt ein
   einfaches Menue (Testpraesentation erzeugen / Beispiel mit echter Marke /
   Kommentare lesen).

Sobald der Arbeits-GitHub-Zugang steht, kommt ein Selbst-Update dazu
(analog zum Rechnungsprogramm) -- bis dahin muss eine neue Version per
neuer .zip-Kopie manuell ersetzt werden.

## Beispiel-Praesentation mit echter Marke ansehen

```bash
export LANGDOCK_API_KEY=<dein eigener Langdock-API-Key>
praesentationsassistent erzeugen-mit-marke "URBAN LOFT" "Neueroeffnung Hotel X" --ausgabe beispiel.pptx
```

Erzeugt eine kurze Beispiel-Praesentation (3-4 Folien) mit den echten
Marken-Farben/-Schriften aus `praesentationsassistent/markenprofile/<marke>.json`.
Die Gliederung
(Folientitel + kurzer Text) wird von Langdock entworfen, bewusst OHNE
konkrete Zahlen/Fakten zu erfinden -- das ist noch kein Dokumenten-
gestuetzter Inhalt, nur ein visueller Abgleich: stimmen Schriftart und
Farbe mit der Marke? `LANGDOCK_API_KEY` bleibt dabei ausschliesslich lokal
in der Umgebungsvariable, wird nirgends gespeichert oder weitergegeben.

## Was heute schon da ist

- `praesentationsassistent/markenprofil.py` -- Datenmodell `Markenprofil`
  (Farben, Schriften, Logo-Regeln, Raster-Hinweise, Bildsprache, Tonalitaet,
  Verbote) + JSON-Laden/Speichern unter `<ordner>/<marke>.json`.
- `praesentationsassistent/markenprofile/{urban_loft,ameron,althoff_hotels,althoff_collection}.json`
  -- die vier echten Markenprofile, extrahiert aus Aarons CD-Manuals/Brand
  Guidelines (30.10.2026). Jeder Eintrag zitiert seine Quellseite im Dokument.
  Schema wurde dabei um `schriften.weitere_schriften` erweitert, weil AMERON
  eine dritte Schriftrolle definiert (TStar Mono Rounded fuer Infos/
  Bildunterschriften neben Haupt-/Nebenschrift). Luecken, weil die jeweiligen
  Manuals dazu nichts Explizites sagen: `tonalitaet` ist bei AMERON, Althoff
  Hotels und Althoff Collection leer (nur URBAN LOFT definiert Tone-of-Voice
  und Markenclaim); `bildsprache.hinweise` ist bei AMERON leer (Styleguide
  trifft keine Aussage zur Bildsprache).
- `praesentationsassistent/pptx_erzeugen.py` -- baut aus einfachen
  Folien-Beschreibungen eine valide `.pptx`: echte Textfelder (kein Text als
  Bild), Bilder als eigene Bild-Objekte, fehlende Bilder als klar
  gekennzeichneter Platzhalter im passenden Seitenverhaeltnis. Optional mit
  `markenprofil=...`: setzt Titel-/Fliesstext-Schriftart und Titelfarbe aus
  der Marke. Manche Markenprofile speichern beim Schriftnamen den vollen
  Zitat-Text (z.B. `"Full Sans (Light, Bold) -- Hauptschrift ... (S.27)"`)
  statt nur den Font-Namen -- `_schriftname_fuer_pptx()` schneidet davon nur
  fuer die PowerPoint-Darstellung den reinen Font-Namen ab, die gespeicherten
  Markenprofil-Dateien bleiben unveraendert.
- `praesentationsassistent/beispiel_erzeugen.py` -- `erzeuge_markenbeispiel()`:
  laedt ein echtes Markenprofil, laesst Langdock eine thematische Gliederung
  entwerfen (ausdruecklich ohne erfundene Zahlen/Fakten) und baut daraus eine
  marken-gestylte Beispiel-`.pptx`. CLI-Befehl: `erzeugen-mit-marke`.
- `praesentationsassistent/bildbearbeitung.py` -- zentrierter Bildzuschnitt
  auf ein Ziel-Seitenverhaeltnis (Pillow), unabhaengig testbar.
- `praesentationsassistent/kommentare.py` -- liest Kommentare aus
  bestehenden `.pptx`-Dateien direkt aus dem OOXML (python-pptx hat dafuer
  keine High-Level-API, siehe Docstring im Modul). Nur Lesen, keine Annahme
  ueber Bedeutung.
- `praesentationsassistent/cli.py` -- Unterbefehle `erzeugen` (Test-PPTX aus
  TESTDATEN) und `kommentare-lesen <datei.pptx>`.
- `praesentationsassistent/langdock.py` -- `generiere_text()`: echter
  Chat-Completions-Aufruf gegen Langdock (`urllib`, OpenAI-kompatibles
  Format). Liest den API-Key ausschliesslich aus `LANGDOCK_API_KEY` (lokale
  Umgebungsvariable); kein Key wird je zentral gespeichert oder mit mir
  geteilt -- entspricht Aarons Vorgabe (30.10.2026), jeder Nutzer verwendet
  seinen eigenen Account. UNGETESTET gegen einen echten Key (keiner in
  dieser Umgebung verfuegbar), nur gegen mehrere Mock-Szenarien
  (`tests/test_langdock.py`).

## Was bewusst noch fehlt

1. Echter Lauf gegen einen echten Langdock-Key (bisher nur gemockt getestet).
2. Keine automatische PDF-Regel-Extraktion zur Laufzeit (die vier Markenprofile
   wurden einmalig von Hand aus den PDFs extrahiert, siehe oben).
3. Keine Marken-Checkliste-Logik, kein voller Dramaturgie-/Dokumenten-
   gestuetzter Text-Generierungs-Workflow -- `erzeugen-mit-marke` ist nur ein
   visueller Abgleich (Farbe/Schrift), nicht echte Inhaltserzeugung.
4. Raster-Hinweise und Logo-Mindestmasse aus dem Markenprofil werden noch
   nicht als Platzierungs-/Pruefregeln umgesetzt (nur gespeichert/gelesen).
5. Revision-via-Kommentare-Workflow (Abschnitt 5 der Spezifikation) und das
   Lern-Gedaechtnis sind noch nicht gebaut.

## Kommentarformate (Stand 06.11.2026)

Beide Kommentarformate sind implementiert und mit je einer handgebauten
Testdatei abgedeckt (siehe `tests/test_kommentare.py`):

- **Legacy** (`ppt/comments/commentN.xml`, `ppt/commentAuthors.xml`) --
  wird auch von der Ueberarbeitungs-Oberflaeche zum SCHREIBEN genutzt.
- **Modern** (PowerPoint fuer Microsoft 365, Namespace `p188` laut
  [MS-PPTX]: `ppt/comments/modernComment_*.xml`, `ppt/authors.xml`) inkl.
  Antworten, Status "resolved" und Verankerung an Folienelementen. Die
  fruehere Annahme, moderne PowerPoint-Kommentare laegen unter
  `commentThreads` (das ist Excels Format), war falsch -- dadurch wurden
  Kommentare aus aktuellem PowerPoint bis Version 0.2.0 gar nicht gefunden.

Offen: Abgleich mit einer **echten**, in PowerPoint 365 kommentierten Datei.
Die Testdatei folgt dem dokumentierten Schema, ist aber handgebaut.

## Aenderungen 06.11.2026 (Version 0.2.1)

- Kommentare aus PowerPoint 365 werden gelesen (siehe oben); `kommentare-lesen`
  zeigt Antworten und Status mit an.
- `windows\start.cmd` startet Python mit `-P`: bisher wurde beim Start aus
  dem Projektordner der dort liegende (aeltere) Quellcode statt der in
  `.venv` installierten Version geladen -- Menuepunkt 4 (Oberflaeche) schlug
  deshalb fehl. Quellcode-Ordner und installierte Version sind jetzt gleich.
- `windows\einrichten.ps1` prueft Python durch echten Aufruf: `python` ist
  unter Windows oft nur der Microsoft-Store-Platzhalter (Abbruch mit Code
  9009, siehe `einrichten.log`); bevorzugt wird jetzt `py -3`.
- `pyproject.toml` nennt alle Abhaengigkeiten (fastapi/uvicorn fehlten).
- Sicherung der vorherigen Dateien: `sicherung_2026-11-06_vor_reparatur.zip`.

## Windows-Verteilung (spaeter)

Die Verteilung auf die Windows-Rechner der Kollegen folgt demselben Muster
wie beim Rechnungsprogramm (`tasks/rechnungsautomatisierung/app/`):
`windows/*-start.cmd`, ein `laptop-einrichten.ps1`-Installer, venv-basiert.
Wird erst gebaut, wenn das Grundgeruest steht und Aaron das sehen will.

## Entwicklung

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -e ".[dev]"
pytest -q
```
