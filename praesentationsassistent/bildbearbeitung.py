"""Bildzuschnitt fuer die PPTX-Erzeugung.

Eigenstaendige, unabhaengig testbare Funktionen (Abschnitt 4 "Technik der
PPTX" im Projektauftrag): ein Bild wird mittig auf ein Ziel-Seitenverhaeltnis
zugeschnitten, ohne es zu verzerren (kein Strecken/Stauchen).
"""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

from PIL import Image


def zentriert_zuschneiden(bild: Image.Image, ziel_seitenverhaeltnis: float) -> Image.Image:
    """Schneidet ``bild`` mittig so zu, dass es exakt ``ziel_seitenverhaeltnis``
    (Breite/Hoehe) hat -- ohne Verzerrung. Gibt ein NEUES Image-Objekt zurueck,
    das uebergebene Bild bleibt unveraendert.

    Ist das Bild zu breit fuer das Ziel, wird links/rechts beschnitten; ist es
    zu hoch, wird oben/unten beschnitten.
    """
    if ziel_seitenverhaeltnis <= 0:
        raise ValueError("ziel_seitenverhaeltnis muss positiv sein")
    breite, hoehe = bild.size
    ist_verhaeltnis = breite / hoehe

    if abs(ist_verhaeltnis - ziel_seitenverhaeltnis) < 1e-6:
        return bild.copy()

    if ist_verhaeltnis > ziel_seitenverhaeltnis:
        # Bild ist (relativ) zu breit -> links/rechts beschneiden
        neue_breite = max(1, round(hoehe * ziel_seitenverhaeltnis))
        links = (breite - neue_breite) // 2
        box = (links, 0, links + neue_breite, hoehe)
    else:
        # Bild ist (relativ) zu hoch -> oben/unten beschneiden
        neue_hoehe = max(1, round(breite / ziel_seitenverhaeltnis))
        oben = (hoehe - neue_hoehe) // 2
        box = (0, oben, breite, oben + neue_hoehe)

    return bild.crop(box)


def bild_datei_zuschneiden(pfad: Path, ziel_seitenverhaeltnis: float) -> BytesIO:
    """Oeffnet die Bilddatei ``pfad``, schneidet sie mittig auf
    ``ziel_seitenverhaeltnis`` zu und gibt das Ergebnis als PNG-Bytestrom
    zurueck (fertig zur Einbettung in eine .pptx via ``add_picture``)."""
    with Image.open(pfad) as bild:
        bild = bild.convert("RGB")
        zugeschnitten = zentriert_zuschneiden(bild, ziel_seitenverhaeltnis)
        puffer = BytesIO()
        zugeschnitten.save(puffer, format="PNG")
        puffer.seek(0)
        return puffer
