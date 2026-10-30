"""Zentrale Konfiguration ueber Umgebungsvariablen.

Vorbild: ``rechnung/config.py`` aus der Rechnungsautomatisierung (selbes
Praefix-Muster, selbe Idee: ohne Konfiguration bleibt alles lokal und mit
sinnvollen Standardwerten lauffaehig).
"""

from __future__ import annotations

import os
from pathlib import Path


def markenprofile_ordner() -> Path:
    """Ordner, in dem Markenprofile als ``<marke>.json`` liegen.

    Ueberschreibbar per Umgebungsvariable ``PRAESENTATION_MARKENPROFILE_ORDNER``
    (z.B. fuer Tests oder einen abweichenden Installationsort). Standard: der
    Ordner ``markenprofile`` neben diesem Paket.
    """
    wert = os.environ.get("PRAESENTATION_MARKENPROFILE_ORDNER")
    if wert:
        return Path(wert)
    return Path(__file__).resolve().parent.parent / "markenprofile"
