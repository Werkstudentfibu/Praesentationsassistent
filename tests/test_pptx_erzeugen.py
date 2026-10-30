"""Tests fuer die PPTX-Erzeugung: valide Datei, echte Textfelder, Bild
bzw. Platzhalter."""

from __future__ import annotations

from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from praesentationsassistent.pptx_erzeugen import FolienBeschreibung, erzeuge_praesentation


def _alle_texte(slide) -> str:
    teile = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            teile.append(shape.text_frame.text)
    return "\n".join(teile)


def test_erzeugt_valide_pptx_mit_echten_textfeldern(tmp_path: Path):
    folien = [
        FolienBeschreibung(
            titel="TESTDATEN Titel",
            text="TESTDATEN Fließtext",
            bild_pfad=None,
            bild_ordner_hinweis="Bilder/x",
        )
    ]
    ziel = tmp_path / "ausgabe.pptx"

    pfad = erzeuge_praesentation(folien, ziel)

    assert pfad.exists()
    prs = Presentation(str(pfad))
    assert len(prs.slides) == 1
    gesamt = _alle_texte(prs.slides[0])
    assert "TESTDATEN Titel" in gesamt
    assert "TESTDATEN Fließtext" in gesamt


def test_mehrere_folien_werden_in_reihenfolge_angelegt(tmp_path: Path):
    folien = [
        FolienBeschreibung(titel="Folie A"),
        FolienBeschreibung(titel="Folie B"),
        FolienBeschreibung(titel="Folie C"),
    ]
    ziel = tmp_path / "mehrere.pptx"

    pfad = erzeuge_praesentation(folien, ziel)

    prs = Presentation(str(pfad))
    assert len(prs.slides) == 3
    titel = [_alle_texte(s).splitlines()[0] for s in prs.slides]
    assert titel == ["Folie A", "Folie B", "Folie C"]


def test_fehlendes_bild_erzeugt_gekennzeichneten_platzhalter(tmp_path: Path):
    folien = [
        FolienBeschreibung(
            titel="T", text="", bild_pfad=None, bild_ordner_hinweis="Bilder/fehlt_hier"
        )
    ]
    ziel = tmp_path / "platzhalter.pptx"

    pfad = erzeuge_praesentation(folien, ziel)

    prs = Presentation(str(pfad))
    gesamt = _alle_texte(prs.slides[0])
    assert "Platzhalter" in gesamt
    assert "Bilder/fehlt_hier" in gesamt
    # Es darf kein Bild-Shape eingefuegt worden sein, nur der Platzhalter-Rahmen.
    bilder = [
        s for s in prs.slides[0].shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE
    ]
    assert bilder == []


def test_vorhandenes_bild_wird_als_bild_shape_eingefuegt(tmp_path: Path):
    bild_pfad = tmp_path / "bild.png"
    Image.new("RGB", (600, 300), color=(200, 10, 10)).save(bild_pfad)
    folien = [
        FolienBeschreibung(
            titel="T", text="", bild_pfad=bild_pfad, bild_ordner_hinweis="egal"
        )
    ]
    ziel = tmp_path / "mit_bild.pptx"

    pfad = erzeuge_praesentation(folien, ziel)

    prs = Presentation(str(pfad))
    bilder = [
        s for s in prs.slides[0].shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE
    ]
    assert len(bilder) == 1
