"""Tests fuer die Marken-Beispielpraesentation -- Langdock komplett
gemockt, kein echter Netzwerkzugriff/API-Key. Markenprofil ist die klar
fiktive "Testmarke Nord", niemals eine echte Marke."""

from __future__ import annotations

import json
from pathlib import Path
from unittest.mock import patch

import pytest
from pptx import Presentation

from praesentationsassistent.beispiel_erzeugen import (
    BeispielErzeugenFehler,
    erzeuge_markenbeispiel,
)
from praesentationsassistent.markenprofil import Farbe, Markenprofil, speichere_markenprofil

_MARKE = "Testmarke Nord"


def _profil_speichern(ordner: Path) -> None:
    profil = Markenprofil(name=_MARKE, farben=[Farbe(name="TESTDATEN Blau", hex_wert="#123456")])
    speichere_markenprofil(profil, ordner)


def test_erzeugt_pptx_aus_gueltiger_langdock_antwort(tmp_path: Path):
    _profil_speichern(tmp_path)
    gueltige_antwort = json.dumps(
        [
            {"titel": "TESTDATEN Einleitung", "text": "TESTDATEN Satz eins."},
            {"titel": "TESTDATEN Abschluss", "text": "TESTDATEN Satz zwei."},
        ]
    )
    ziel = tmp_path / "beispiel.pptx"

    with patch(
        "praesentationsassistent.beispiel_erzeugen.generiere_text",
        return_value=gueltige_antwort,
    ):
        pfad = erzeuge_markenbeispiel(_MARKE, "TESTDATEN Thema", ziel, tmp_path)

    assert pfad.exists()
    prs = Presentation(str(pfad))
    assert len(prs.slides) == 2


def test_ungueltiges_json_wirft_klaren_fehler(tmp_path: Path):
    _profil_speichern(tmp_path)
    ziel = tmp_path / "beispiel.pptx"

    with patch(
        "praesentationsassistent.beispiel_erzeugen.generiere_text",
        return_value="das ist kein JSON",
    ):
        with pytest.raises(BeispielErzeugenFehler, match="kein gueltiges JSON"):
            erzeuge_markenbeispiel(_MARKE, "TESTDATEN Thema", ziel, tmp_path)


def test_leere_liste_wirft_klaren_fehler(tmp_path: Path):
    _profil_speichern(tmp_path)
    ziel = tmp_path / "beispiel.pptx"

    with patch(
        "praesentationsassistent.beispiel_erzeugen.generiere_text", return_value="[]"
    ):
        with pytest.raises(BeispielErzeugenFehler, match="erwartete Format"):
            erzeuge_markenbeispiel(_MARKE, "TESTDATEN Thema", ziel, tmp_path)
