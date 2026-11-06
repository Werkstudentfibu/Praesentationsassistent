"""Erzeugt eine Beispiel-Praesentation fuer eine echte Marke: echte Farben,
echte Schriften aus dem Markenprofil, Gliederung per Langdock generiert.

Nichts wird erfunden, was nicht gegeben ist: der Langdock-Prompt verlangt
ausdruecklich allgemeine Gliederungs-/Ueberschriftstexte zum genannten
Thema, KEINE konkreten Zahlen/Fakten/Statistiken -- die hat niemand
mitgegeben, also darf das Modell auch keine nennen. Das ist bewusst noch
keine echte Dokument-gestuetzte Inhaltserzeugung (Abschnitt 3 der
Spezifikation) -- nur ein visueller Marken-Abgleich zum Anschauen.
"""

from __future__ import annotations

import json
from pathlib import Path

from .langdock import LangdockFehler, generiere_text
from .markenprofil import Markenprofil, lade_markenprofil
from .pptx_erzeugen import FolienBeschreibung, erzeuge_praesentation

_STANDARD_MARKENPROFILE_ORDNER = Path(__file__).resolve().parent / "markenprofile"


class BeispielErzeugenFehler(Exception):
    """Fehler beim Erzeugen der Marken-Beispielpraesentation."""


def _gliederung_anfordern(marke: str, thema: str, profil: Markenprofil) -> list[dict]:
    system = (
        f"Du hilfst, die Gliederung einer PowerPoint-Praesentation fuer die "
        f"Marke {marke} zu entwerfen. Antworte AUSSCHLIESSLICH mit einem "
        f"JSON-Array, jedes Element mit den Feldern \"titel\" und \"text\" "
        f"(1-2 kurze Saetze). Erzeuge 3 bis 4 Folien. Nenne KEINE konkreten "
        f"Zahlen, Statistiken oder Fakten -- die wurden dir nicht gegeben, "
        f"erfinde also keine. Halte dich nur an allgemeine, thematische "
        f"Gliederungs- und Ueberschriftstexte."
    )
    if profil.tonalitaet:
        system += f" Tonalitaet der Marke: {profil.tonalitaet}"

    antwort = generiere_text(f"Thema der Praesentation: {thema}", system=system)
    try:
        gliederung = json.loads(antwort)
    except json.JSONDecodeError as exc:
        raise BeispielErzeugenFehler(
            f"Langdock-Antwort war kein gueltiges JSON: {antwort[:500]!r}"
        ) from exc
    if not isinstance(gliederung, list) or not gliederung:
        raise BeispielErzeugenFehler(
            f"Langdock-Antwort hatte nicht das erwartete Format (Liste): {antwort[:500]!r}"
        )
    return gliederung


def erzeuge_markenbeispiel(
    marke: str,
    thema: str,
    ausgabe_pfad: Path | str,
    markenprofile_ordner: Path | str = _STANDARD_MARKENPROFILE_ORDNER,
) -> Path:
    """Laedt das echte Markenprofil ``marke``, laesst per Langdock eine
    thematische Gliederung entwerfen und baut daraus eine .pptx mit den
    echten Marken-Schriften/Farben. Wirft ``BeispielErzeugenFehler`` bzw.
    ``LangdockFehler``, statt bei Problemen stillschweigend etwas zu raten.
    """
    profil = lade_markenprofil(marke, Path(markenprofile_ordner))

    try:
        gliederung = _gliederung_anfordern(marke, thema, profil)
    except LangdockFehler:
        raise

    folien = [
        FolienBeschreibung(
            titel=str(eintrag.get("titel", "")).strip() or "(ohne Titel)",
            text=str(eintrag.get("text", "")).strip(),
            bild_ordner_hinweis=f"Bilder/{marke}",
        )
        for eintrag in gliederung
    ]
    return erzeuge_praesentation(folien, ausgabe_pfad, markenprofil=profil)
