# Praesentationsassistent (Grundgeruest)

Visual-Identity-Praesentationsassistent fuer die Hotelgruppe (Marken:
Althoff Hotels, Althoff Collection, AMERON, URBAN LOFT). Ziel: aus
Dokumenten + Marken-Regeln automatisch PowerPoint-Praesentationen erzeugen,
die verbindlich eine der vier Marken-Visual-Identities einhalten; bestehende
.pptx anhand von Kommentaren/Aenderungslisten gezielt ueberarbeiten; dabei
nie Fakten erfinden.

**Stand heute: nur das Grundgeruest.** Die eigentliche Marken-Regel-Logik
und der LLM-Anschluss (Langdock) fehlen bewusst noch.

## Was heute schon da ist

- `praesentationsassistent/markenprofil.py` -- Datenmodell `Markenprofil`
  (Farben, Schriften, Logo-Regeln, Raster-Hinweise, Bildsprache, Tonalitaet,
  Verbote) + JSON-Laden/Speichern unter `markenprofile/<marke>.json`.
- `markenprofile/{urban_loft,ameron,althoff_hotels,althoff_collection}.json`
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
  gekennzeichneter Platzhalter im passenden Seitenverhaeltnis.
- `praesentationsassistent/bildbearbeitung.py` -- zentrierter Bildzuschnitt
  auf ein Ziel-Seitenverhaeltnis (Pillow), unabhaengig testbar.
- `praesentationsassistent/kommentare.py` -- liest Kommentare aus
  bestehenden `.pptx`-Dateien direkt aus dem OOXML (python-pptx hat dafuer
  keine High-Level-API, siehe Docstring im Modul). Nur Lesen, keine Annahme
  ueber Bedeutung.
- `praesentationsassistent/cli.py` -- Unterbefehle `erzeugen` (Test-PPTX aus
  TESTDATEN) und `kommentare-lesen <datei.pptx>`.
- `praesentationsassistent/langdock.py` -- leerer Erweiterungspunkt fuer die
  spaetere LLM-Anbindung (`generiere_text()` wirft bewusst `NotImplementedError`).

## Was bewusst noch fehlt

1. **Langdock-Zugangsdaten** -- erst damit wird `langdock.py` implementiert.
2. Keine automatische PDF-Regel-Extraktion zur Laufzeit (die vier Markenprofile
   wurden einmalig von Hand aus den PDFs extrahiert, siehe oben).
3. Keine Marken-Checkliste-Logik, kein Dramaturgie-/Text-Generierungs-Workflow.

## Bekannte Luecke in der Testsuite

Beide Kommentarformate sind implementiert und mit je einer handgebauten
Testdatei abgedeckt (siehe `tests/test_kommentare.py`):

- **Legacy** (`ppt/comments/commentN.xml`, `ppt/commentAuthors.xml`).
- **Modern** (PowerPoint fuer Microsoft 365, Namespace `p188` laut
  [MS-PPTX]: `ppt/comments/modernComment_*.xml`, `ppt/authors.xml`) inkl.
  Antworten, Status "resolved" und Verankerung an Folienelementen. Die
  fruehere Annahme, moderne PowerPoint-Kommentare laegen unter
  `commentThreads` (das ist Excels Format), war falsch -- dadurch wurden
  Kommentare aus aktuellem PowerPoint bisher gar nicht gefunden.

Offen: Abgleich mit einer **echten**, in PowerPoint 365 kommentierten Datei.
Die Testdatei folgt dem dokumentierten Schema, ist aber handgebaut.

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
