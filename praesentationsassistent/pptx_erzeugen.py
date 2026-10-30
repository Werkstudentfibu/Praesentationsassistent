"""PPTX-Erzeugung-Grundgeruest.

Baut aus einer Liste einfacher Folien-Beschreibungen eine valide .pptx:
Text steht als echtes Textfeld (kein Text-als-Bild), Bilder als eigene
Bild-Objekte (mittig zugeschnitten, siehe bildbearbeitung.py), und fehlende
Bilder werden als klar gekennzeichneter Platzhalter dargestellt statt leer
gelassen zu werden.

Die eigentliche Markenregel-Anwendung (Farben/Schriften/Raster aus einem
Markenprofil) ist bewusst noch NICHT eingebaut -- das ist der naechste
Schritt, sobald Aarons Regel-Datei vorliegt.
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

_GRAU_FUELLUNG = RGBColor(0xBF, 0xBF, 0xBF)
_GRAU_RAND = RGBColor(0x80, 0x80, 0x80)


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
    folien: list[FolienBeschreibung], ausgabe_pfad: Path | str
) -> Path:
    """Baut aus ``folien`` eine valide .pptx-Datei unter ``ausgabe_pfad``.

    Jede Folie bekommt einen Titel (Textfeld), optional einen Fliesstext
    (Textfeld) und rechts ein Bild bzw. einen Platzhalter.
    """
    prs = Presentation()
    leeres_layout = prs.slide_layouts[6]  # "Leer" im Standard-Template

    for beschreibung in folien:
        slide = prs.slides.add_slide(leeres_layout)

        titel_feld = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.3), Inches(9.0), Inches(1.0)
        )
        titel_feld.text_frame.text = beschreibung.titel
        titel_absatz = titel_feld.text_frame.paragraphs[0]
        titel_absatz.font.size = Pt(28)
        titel_absatz.font.bold = True

        if beschreibung.text:
            text_feld = slide.shapes.add_textbox(
                Inches(0.5), Inches(1.5), Inches(4.5), Inches(5.0)
            )
            text_feld.text_frame.word_wrap = True
            text_feld.text_frame.text = beschreibung.text

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
