"""Tests fuer das Markenprofil-Datenmodell.

Verwendet ausschliesslich die klar fiktive Marke "Testmarke Nord" mit frei
erfundenen TESTDATEN -- niemals echte Markennamen (Althoff/AMERON/URBAN LOFT).
"""

from __future__ import annotations

from pathlib import Path

from praesentationsassistent.markenprofil import (
    Bildsprache,
    Farbe,
    Logoregeln,
    Markenprofil,
    Schriften,
    lade_markenprofil,
    leeres_profil,
    markenprofil_pfad,
    speichere_markenprofil,
)

_MARKE = "Testmarke Nord"


def _beispielprofil() -> Markenprofil:
    return Markenprofil(
        name=_MARKE,
        farben=[
            Farbe(name="TESTDATEN Blau", hex_wert="#123456"),
            Farbe(name="TESTDATEN Gelb", hex_wert="#FFD700"),
        ],
        schriften=Schriften(hauptschrift="TESTDATEN Sans", nebenschrift="TESTDATEN Serif"),
        logoregeln=Logoregeln(
            mindestgroesse_mm=20.0,
            schutzraum_mm=5.0,
            erlaubte_versionen=["logo_testdaten_farbig.svg", "logo_testdaten_weiss.svg"],
        ),
        raster_hinweise="TESTDATEN: 12-Spalten-Raster, Randabstand 10 mm",
        bildsprache=Bildsprache(
            keine_stockbilder=True, keine_ki_bilder=True, hinweise="TESTDATEN: nur eigene Fotos"
        ),
        tonalitaet="TESTDATEN: locker, aber nicht albern",
        verbote=["TESTDATEN: kein Comic Sans", "TESTDATEN: kein Vollbild-Logo"],
    )


def test_speichern_legt_datei_am_erwarteten_pfad_an(tmp_path: Path):
    profil = _beispielprofil()
    pfad = speichere_markenprofil(profil, tmp_path)
    assert pfad.exists()
    assert pfad == markenprofil_pfad(_MARKE, tmp_path)
    assert pfad.suffix == ".json"


def test_speichern_und_laden_roundtrip(tmp_path: Path):
    profil = _beispielprofil()
    speichere_markenprofil(profil, tmp_path)
    geladen = lade_markenprofil(_MARKE, tmp_path)
    assert geladen == profil


def test_leeres_profil_enthaelt_keine_erfundenen_inhalte():
    profil = leeres_profil(_MARKE)
    assert profil.name == _MARKE
    assert profil.farben == []
    assert profil.verbote == []
    assert profil.tonalitaet == ""
    assert profil.schriften == Schriften()
    assert profil.logoregeln == Logoregeln()
    assert profil.bildsprache == Bildsprache()


def test_leeres_profil_kann_gespeichert_und_wieder_geladen_werden(tmp_path: Path):
    profil = leeres_profil(_MARKE)
    speichere_markenprofil(profil, tmp_path)
    geladen = lade_markenprofil(_MARKE, tmp_path)
    assert geladen == profil


def test_dateiname_wird_aus_markenname_abgeleitet(tmp_path: Path):
    pfad = markenprofil_pfad("Testmarke Nord", tmp_path)
    assert pfad.name == "testmarke_nord.json"


def test_weitere_schriften_roundtrip(tmp_path: Path):
    """Manche CD-Manuals definieren mehr als zwei Schriftrollen (Befund bei
    AMERON: zusaetzliche Schrift fuer Bildunterschriften)."""
    profil = Markenprofil(
        name=_MARKE,
        schriften=Schriften(
            hauptschrift="TESTDATEN Sans",
            nebenschrift="TESTDATEN Serif",
            weitere_schriften=["TESTDATEN Mono (Bildunterschriften)"],
        ),
    )
    speichere_markenprofil(profil, tmp_path)
    geladen = lade_markenprofil(_MARKE, tmp_path)
    assert geladen.schriften.weitere_schriften == ["TESTDATEN Mono (Bildunterschriften)"]
