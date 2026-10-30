"""Tests fuer den Bildzuschnitt (zentriert, ohne Verzerrung)."""

from __future__ import annotations

from pathlib import Path

from PIL import Image

from praesentationsassistent.bildbearbeitung import (
    bild_datei_zuschneiden,
    zentriert_zuschneiden,
)


def test_zu_breites_bild_wird_seitlich_beschnitten():
    bild = Image.new("RGB", (800, 400), color=(10, 20, 30))  # Seitenverhaeltnis 2:1
    ergebnis = zentriert_zuschneiden(bild, ziel_seitenverhaeltnis=1.0)  # 1:1
    assert ergebnis.size == (400, 400)


def test_zu_hohes_bild_wird_oben_unten_beschnitten():
    bild = Image.new("RGB", (400, 800), color=(10, 20, 30))  # Seitenverhaeltnis 1:2
    ergebnis = zentriert_zuschneiden(bild, ziel_seitenverhaeltnis=1.0)  # 1:1
    assert ergebnis.size == (400, 400)


def test_passendes_seitenverhaeltnis_bleibt_unveraendert():
    bild = Image.new("RGB", (600, 400), color=(1, 2, 3))  # 3:2
    ergebnis = zentriert_zuschneiden(bild, ziel_seitenverhaeltnis=1.5)
    assert ergebnis.size == (600, 400)


def test_zuschnitt_ist_zentriert_nicht_am_rand():
    # Links die Haelfte rot, rechts die Haelfte blau -- nach dem Zuschnitt
    # auf ein quadratisches Format sollten (bei einem 2:1-Bild) beide
    # Farben noch vorkommen, weil der Zuschnitt mittig erfolgt.
    bild = Image.new("RGB", (800, 400))
    for x in range(800):
        farbe = (255, 0, 0) if x < 400 else (0, 0, 255)
        for y in range(400):
            bild.putpixel((x, y), farbe)
    ergebnis = zentriert_zuschneiden(bild, ziel_seitenverhaeltnis=1.0)
    farben = {ergebnis.getpixel((x, 200)) for x in (0, ergebnis.width - 1)}
    assert (255, 0, 0) in farben
    assert (0, 0, 255) in farben


def test_original_bild_bleibt_unveraendert():
    bild = Image.new("RGB", (800, 400), color=(10, 20, 30))
    original_groesse = bild.size
    zentriert_zuschneiden(bild, ziel_seitenverhaeltnis=1.0)
    assert bild.size == original_groesse


def test_bild_datei_zuschneiden_liefert_korrektes_seitenverhaeltnis(tmp_path: Path):
    quelle = tmp_path / "testbild.png"
    Image.new("RGB", (900, 300), color=(5, 5, 5)).save(quelle)

    puffer = bild_datei_zuschneiden(quelle, ziel_seitenverhaeltnis=1.0)

    ergebnis = Image.open(puffer)
    assert ergebnis.format == "PNG"
    assert ergebnis.size[0] == ergebnis.size[1]
