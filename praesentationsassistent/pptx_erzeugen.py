"""PPTX-Erzeugung-Grundgeruest.

Baut aus einer Liste einfacher Folien-Beschreibungen eine valide .pptx:
Text steht als echtes Textfeld (kein Text-als-Bild), Bilder als eigene
Bild-Objekte (mittig zugeschnitten, siehe bildbearbeitung.py), und fehlende
Bilder werden als klar gekennzeichneter Platzhalter dargestellt statt leer
gelassen zu werden.

Markenregel-Anwendung: wird ``erzeuge_praesentation`` ein ``Markenprofil``
mitgegeben, setzt sie Titel-/Fliesstext-Schriftart sowie die Titelfarbe aus
der Marke. Raster-Hinweise und Logo-Mindestmasse sind bislang nicht in
eigenen Platzierungs-/Pruefregeln umgesetzt.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

from .bildbearbeitung import bild_datei_zuschneiden
from .markenprofil import Markenprofil

_GRAU_FUELLUNG = RGBColor(0xBF, 0xBF, 0xBF)
_GRAU_RAND = RGBColor(0x80, 0x80, 0x80)


def _hex_zu_rgbcolor(hex_wert: str) -> RGBColor | None:
    """Wandelt ``"#RRGGBB"`` in ``RGBColor`` um, oder ``None`` wenn kein
    brauchbarer Hex-Wert vorliegt (z.B. Veredelungsfarben ohne Hex im CD-
    Manual -- siehe Luecken in den Markenprofilen)."""
    wert = (hex_wert or "").lstrip("#")
    if len(wert) != 6:
        return None
    try:
        return RGBColor(int(wert[0:2], 16), int(wert[2:4], 16), int(wert[4:6], 16))
    except ValueError:
        return None


def _schriftname_fuer_pptx(wert: str) -> str:
    """Extrahiert aus einem Schriften-Eintrag im Markenprofil den reinen,
    bei PowerPoint eintragbaren Font-Familiennamen.

    Die Markenprofile speichern bei manchen Marken bewusst den vollen
    Dokumentations-Text inkl. Seitenzitat statt nur den Font-Namen (Prinzip
    "nichts wird geraten" -- die Herkunft soll nachvollziehbar bleiben),
    z.B. ``"Full Sans (Light, Book, Bold) -- Hauptschrift ... (S.27)"``.
    Als ``font.name`` waere das ein nicht existierender Font und PowerPoint
    wuerde kommentarlos auf eine Ersatzschrift zurueckfallen. Diese Funktion
    schneidet deshalb nur fuer die Darstellung den Namen vor der ersten
    Erklaerung/Klammer ab -- die gespeicherten Markenprofil-Dateien selbst
    bleiben unveraendert (volle Zitate bleiben dort erhalten)."""
    name = wert
    for trenner in (" -- ", " ("):
        name = name.split(trenner, 1)[0]
    # Mehrere gleichwertige Schriftschnitte, mit "/" aufgezaehlt (z.B.
    # "Cheswick Grotesque Regular / Cheswick Grotesque Semibold"): nur den
    # ersten als Standard-Darstellungsschnitt verwenden.
    name = name.split(" / ", 1)[0]
    return name.strip()


@dataclass
class FolienBeschreibung:
    """Einfache, markenunabhaengige Beschreibung einer Folie."""

    titel: str
    text: str = ""
    bild_pfad: Path | None = None
    # Hinweis, der im Platzhalter erscheint, falls kein Bild vorhanden ist
    # (z.B. ein Hinweis auf den Ordner, in dem das Bild erwartet wird).
    bild_ordner_hinweis: str = "Bilder"


def _bildfeld_hinzufuegen(
    slide,
    links: Emu,
    oben: Emu,
    breite: Emu,
    hoehe: Emu,
    bild_pfad: Path | None,
    ordner_hinweis: str,
) -> None:
    """Fuegt an der angegebenen Stelle ein Bild ein -- oder, falls keines
    vorhanden ist, einen klar gekennzeichneten Platzhalter im passenden
    Seitenverhaeltnis."""
    ziel_seitenverhaeltnis = breite / hoehe

    if bild_pfad is not None and Path(bild_pfad).exists():
        puffer = bild_datei_zuschneiden(Path(bild_pfad), ziel_seitenverhaeltnis)
        slide.shapes.add_picture(puffer, links, oben, width=breite, height=hoehe)
        return

    form = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, links, oben, breite, hoehe)
    form.fill.solid()
    form.fill.fore_color.rgb = _GRAU_FUELLUNG
    form.line.color.rgb = _GRAU_RAND
    form.shadow.inherit = False
    tf = form.text_frame
    tf.word_wrap = True
    tf.text = f"Platzhalter: {ordner_hinweis}"
    for absatz in tf.paragraphs:
        absatz.alignment = PP_ALIGN.CENTER
        for lauf in absatz.runs:
            lauf.font.size = Pt(14)
            lauf.font.color.rgb = RGBColor(0x40, 0x40, 0x40)


def erzeuge_praesentation(
    folien: list[FolienBeschreibung],
    ausgabe_pfad: Path | str,
    markenprofil: Markenprofil | None = None,
) -> Path:
    """Baut aus ``folien`` eine valide .pptx-Datei unter ``ausgabe_pfad``.

    Jede Folie bekommt einen Titel (Textfeld), optional einen Fliesstext
    (Textfeld) und rechts ein Bild bzw. einen Platzhalter.

    Wird ein ``markenprofil`` mitgegeben, werden Titel/Fliesstext mit der
    dort hinterlegten Haupt-/Nebenschrift gesetzt und der Titel in der
    ersten Markenfarbe mit gueltigem Hex-Wert eingefaerbt (Farben ohne Hex
    im CD-Manual -- z.B. reine Veredelungsangaben -- werden dabei
    uebersprungen statt geraten). Ohne Markenprofil bleibt es beim
    PowerPoint-Standard-Erscheinungsbild wie bisher.
    """
    prs = Presentation()
    leeres_layout = prs.slide_layouts[6]  # "Leer" im Standard-Template

    hauptschrift = ""
    nebenschrift = ""
    if markenprofil:
        if markenprofil.schriften.hauptschrift:
            hauptschrift = _schriftname_fuer_pptx(markenprofil.schriften.hauptschrift)
        if markenprofil.schriften.nebenschrift:
            nebenschrift = _schriftname_fuer_pptx(markenprofil.schriften.nebenschrift)
    titelfarbe = None
    if markenprofil:
        for farbe in markenprofil.farben:
            titelfarbe = _hex_zu_rgbcolor(farbe.hex_wert)
            if titelfarbe is not None:
                break

    for beschreibung in folien:
        slide = prs.slides.add_slide(leeres_layout)

        titel_feld = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.3), Inches(9.0), Inches(1.0)
        )
        titel_feld.text_frame.text = beschreibung.titel
        titel_absatz = titel_feld.text_frame.paragraphs[0]
        titel_absatz.font.size = Pt(28)
        titel_absatz.font.bold = True
        if hauptschrift:
            titel_absatz.font.name = hauptschrift
        if titelfarbe is not None:
            titel_absatz.font.color.rgb = titelfarbe

        if beschreibung.text:
            text_feld = slide.shapes.add_textbox(
                Inches(0.5), Inches(1.5), Inches(4.5), Inches(5.0)
            )
            text_feld.text_frame.word_wrap = True
            text_feld.text_frame.text = beschreibung.text
            if nebenschrift:
                for absatz in text_feld.text_frame.paragraphs:
                    absatz.font.name = nebenschrift

        _bildfeld_hinzufuegen(
            slide,
            Inches(5.3),
            Inches(1.5),
            Inches(4.2),
            Inches(4.2),
            beschreibung.bild_pfad,
            beschreibung.bild_ordner_hinweis,
        )

    ausgabe_pfad = Path(ausgabe_pfad)
    ausgabe_pfad.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(ausgabe_pfad))
    return ausgabe_pfad
