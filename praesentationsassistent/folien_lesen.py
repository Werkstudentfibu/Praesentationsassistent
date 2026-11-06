"""Liest die reale Folien-/Element-Struktur einer bestehenden .pptx aus --
fuer die Ueberarbeitungs-Oberflaeche (web/), die echte Folien statt
erfundener Platzhalterdaten anzeigen soll.

Deckt genau das ab, was ``pptx_erzeugen.erzeuge_praesentation`` tatsaechlich
baut (Titel-Textfeld, optionales Fliesstext-Textfeld, Bild-oder-Platzhalter-
Rechteck) -- bei einer von Hand in PowerPoint gebauten Datei mit anderer
Shape-Reihenfolge waere die Titel-/Text-Zuordnung nur eine grobe Vermutung;
hier trifft sie zu, weil wir die Herkunft der Datei kennen.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


@dataclass
class ElementBeschreibung:
    """Ein einzelnes, klickbares Element einer Folie (Prozent-Koordinaten
    relativ zur Folienbreite/-hoehe, wie von der Oberflaeche gebraucht)."""

    id: str
    typ: str  # "title" | "text" | "image" | "sonstiges"
    label: str
    x: float
    y: float
    w: float
    h: float
    text: str = ""


@dataclass
class FolieBeschreibung:
    titel: str
    elemente: list[ElementBeschreibung] = field(default_factory=list)


def _ist_platzhalter_rechteck(shape) -> bool:
    """Das von ``_bildfeld_hinzufuegen`` angelegte graue Platzhalter-
    Rechteck, falls kein echtes Bild vorhanden war -- erkennbar daran, dass
    es ein Auto-Shape mit Text ist, der mit "Platzhalter:" beginnt."""
    if shape.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
        return False
    if not shape.has_text_frame:
        return False
    return shape.text_frame.text.startswith("Platzhalter:")


def lese_folien(pptx_pfad: Path | str) -> list[FolieBeschreibung]:
    """Liest alle Folien einer .pptx mit ihren Elementen (Position in
    Prozent, Typ, aktueller Text) fuer die Ueberarbeitungs-Oberflaeche."""
    prs = Presentation(str(pptx_pfad))
    breite = prs.slide_width
    hoehe = prs.slide_height

    ergebnisse: list[FolieBeschreibung] = []
    for s_idx, slide in enumerate(prs.slides):
        elemente: list[ElementBeschreibung] = []
        titel_text = ""
        text_gesehen = False

        for sh_idx, shape in enumerate(slide.shapes):
            x = (shape.left or 0) / breite * 100
            y = (shape.top or 0) / hoehe * 100
            w = (shape.width or 0) / breite * 100
            h = (shape.height or 0) / hoehe * 100
            elem_id = f"s{s_idx}-e{sh_idx}"

            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                elemente.append(
                    ElementBeschreibung(elem_id, "image", "Bild", x, y, w, h)
                )
            elif _ist_platzhalter_rechteck(shape):
                elemente.append(
                    ElementBeschreibung(
                        elem_id, "image", "Bild (Platzhalter)", x, y, w, h
                    )
                )
            elif shape.has_text_frame and shape.text_frame.text:
                text = shape.text_frame.text
                if not text_gesehen:
                    text_gesehen = True
                    titel_text = text
                    elemente.append(
                        ElementBeschreibung(elem_id, "title", "Titel", x, y, w, h, text)
                    )
                else:
                    elemente.append(
                        ElementBeschreibung(
                            elem_id, "text", "Textblock", x, y, w, h, text
                        )
                    )
            else:
                elemente.append(
                    ElementBeschreibung(
                        elem_id, "sonstiges", f"Element {sh_idx + 1}", x, y, w, h
                    )
                )

        ergebnisse.append(
            FolieBeschreibung(titel=titel_text or f"Folie {s_idx + 1}", elemente=elemente)
        )
    return ergebnisse
