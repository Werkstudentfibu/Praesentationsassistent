"""Markenprofile: Datenmodell fuer die Visual-Identity-Regeln einer Marke.

WICHTIG: Dieses Modul enthaelt bewusst KEINE Inhalte fuer die echten Marken
(Althoff Hotels, Althoff Collection, AMERON, URBAN LOFT). Es stellt nur die
Datenstruktur sowie Lade-/Speicherfunktionen bereit. Echte Markenprofile
werden erst angelegt, wenn Aaron die verbindliche Regel-Datei liefert --
siehe README.md. Fuer Tests wird ausschliesslich die klar fiktive Marke
"Testmarke Nord" mit frei erfundenen TESTDATEN verwendet, damit nichts mit
echten Markenwerten verwechselt werden kann.
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path


@dataclass
class Farbe:
    """Eine benannte Farbe der Markenpalette."""

    name: str
    hex_wert: str  # z.B. "#1A1A1A"


@dataclass
class Schriften:
    """Haupt- und Nebenschrift der Marke, plus optionale weitere Schriften
    (z.B. eine dritte Schrift fuer Bildunterschriften/Infotexte -- manche
    CD-Manuals definieren mehr als zwei Schriftrollen, Befund beim AMERON-
    Profil: Bodoni Sans/TStar Pro als Haupt-/Nebenschrift, zusaetzlich
    TStar Mono Rounded fuer Infos/Bildunterschriften)."""

    hauptschrift: str = ""
    nebenschrift: str = ""
    weitere_schriften: list[str] = field(default_factory=list)


@dataclass
class Logoregeln:
    """Regeln fuer die Logo-Verwendung."""

    mindestgroesse_mm: float | None = None
    schutzraum_mm: float | None = None
    # Pfade/Dateinamen der erlaubten Logo-Versionen (z.B. "logo_farbig.svg")
    erlaubte_versionen: list[str] = field(default_factory=list)


@dataclass
class Bildsprache:
    """Regeln zur Bildsprache."""

    keine_stockbilder: bool = False
    keine_ki_bilder: bool = False
    hinweise: str = ""


@dataclass
class Markenprofil:
    """Vollstaendiges Visual-Identity-Profil einer Marke.

    Noch ohne Inhalte -- reine Struktur, die spaeter aus Aarons Regel-Datei
    befuellt wird.
    """

    name: str
    farben: list[Farbe] = field(default_factory=list)
    schriften: Schriften = field(default_factory=Schriften)
    logoregeln: Logoregeln = field(default_factory=Logoregeln)
    raster_hinweise: str = ""
    bildsprache: Bildsprache = field(default_factory=Bildsprache)
    tonalitaet: str = ""
    verbote: list[str] = field(default_factory=list)


def leeres_profil(name: str) -> Markenprofil:
    """Erzeugt ein leeres Markenprofil-Geruest fuer ``name`` -- ohne jegliche
    inhaltliche Annahme. Wird spaeter aus Aarons Regel-Datei befuellt."""
    return Markenprofil(name=name)


def _dateiname(marke: str) -> str:
    """Leitet aus dem Markennamen einen sicheren Dateinamen ab (z.B.
    "Testmarke Nord" -> "testmarke_nord.json")."""
    sicher = "".join(c.lower() if c.isalnum() else "_" for c in marke).strip("_")
    while "__" in sicher:
        sicher = sicher.replace("__", "_")
    return f"{sicher}.json"


def markenprofil_pfad(marke: str, ordner: Path) -> Path:
    """Pfad, unter dem das Markenprofil ``marke`` in ``ordner`` liegt/liegen soll."""
    return Path(ordner) / _dateiname(marke)


def speichere_markenprofil(profil: Markenprofil, ordner: Path) -> Path:
    """Speichert ``profil`` als JSON unter ``ordner/<marke>.json``."""
    ordner = Path(ordner)
    ordner.mkdir(parents=True, exist_ok=True)
    pfad = markenprofil_pfad(profil.name, ordner)
    pfad.write_text(
        json.dumps(asdict(profil), indent=2, ensure_ascii=False), encoding="utf-8"
    )
    return pfad


def lade_markenprofil(marke: str, ordner: Path) -> Markenprofil:
    """Laedt das Markenprofil ``marke`` aus ``ordner/<marke>.json``."""
    pfad = markenprofil_pfad(marke, Path(ordner))
    daten = json.loads(pfad.read_text(encoding="utf-8"))
    return _aus_dict(daten)


def _aus_dict(daten: dict) -> Markenprofil:
    return Markenprofil(
        name=daten["name"],
        farben=[Farbe(**f) for f in daten.get("farben", [])],
        schriften=Schriften(**daten.get("schriften", {})),
        logoregeln=Logoregeln(**daten.get("logoregeln", {})),
        raster_hinweise=daten.get("raster_hinweise", ""),
        bildsprache=Bildsprache(**daten.get("bildsprache", {})),
        tonalitaet=daten.get("tonalitaet", ""),
        verbote=daten.get("verbote", []),
    )
