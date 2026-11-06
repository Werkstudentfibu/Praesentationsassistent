"""Lokale Ueberarbeitungs-Oberflaeche: zeigt echte, bereits erzeugte
Praesentationen folienweise an, laesst pro Element oder pro ganzer Folie
einen Aenderungswunsch eintragen und schreibt diese beim Klick auf
"Aenderungen umsetzen" als echte PowerPoint-Kommentare in die Datei
(``kommentare.schreibe_kommentare``).

Setzt Aenderungen NICHT automatisch um (kein KI-Schreibzugriff auf Layout/
Text in dieser Stufe) -- das entspricht Abschnitt 5 der Spezifikation nur
in der Eingabe-Haelfte (Kommentare sammeln); das automatische Umsetzen
gegen das Markenprofil ist ein spaeterer, noch nicht gebauter Schritt.
"""

from __future__ import annotations

from dataclasses import asdict
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
from pptx import Presentation
from pydantic import BaseModel

from ..folien_lesen import lese_folien
from ..kommentare import NeuerKommentar, schreibe_kommentare

_STATIC_ORDNER = Path(__file__).resolve().parent / "static"


class AenderungEingabe(BaseModel):
    folie: int  # 1-basiert
    text: str
    elementLabel: str
    x: float | None = None  # Prozent der Folienbreite, falls Element-bezogen
    y: float | None = None


class KommentareEingabe(BaseModel):
    aenderungen: list[AenderungEingabe]


def _pptx_pfad_aufloesen(ordner: Path, dateiname: str) -> Path:
    """Loest ``dateiname`` sicher innerhalb von ``ordner`` auf -- lehnt
    Pfad-Ausbrueche (z.B. "../") und Nicht-.pptx-Dateien ab."""
    if not dateiname.lower().endswith(".pptx"):
        raise HTTPException(400, "Nur .pptx-Dateien.")
    kandidat = (ordner / dateiname).resolve()
    if kandidat.parent != ordner.resolve() or not kandidat.exists():
        raise HTTPException(404, f"Datei nicht gefunden: {dateiname}")
    return kandidat


def erzeuge_app(ordner: Path | str) -> FastAPI:
    """Baut die FastAPI-App; ``ordner`` ist der Ordner mit den zu
    ueberarbeitenden .pptx-Dateien (z.B. "ausgabe")."""
    ordner = Path(ordner)
    ordner.mkdir(parents=True, exist_ok=True)
    app = FastAPI(title="Praesentationsassistent - Ueberarbeiten")

    @app.get("/", response_class=HTMLResponse)
    def index() -> str:
        return (_STATIC_ORDNER / "ueberarbeiten.html").read_text(encoding="utf-8")

    @app.get("/api/praesentationen")
    def liste() -> list[str]:
        return sorted(p.name for p in ordner.glob("*.pptx"))

    @app.get("/api/praesentationen/{dateiname}")
    def folien(dateiname: str) -> list[dict]:
        pfad = _pptx_pfad_aufloesen(ordner, dateiname)
        return [asdict(f) for f in lese_folien(pfad)]

    @app.post("/api/praesentationen/{dateiname}/kommentare")
    def kommentare_schreiben(dateiname: str, eingabe: KommentareEingabe) -> dict:
        pfad = _pptx_pfad_aufloesen(ordner, dateiname)
        prs = Presentation(str(pfad))
        breite, hoehe = prs.slide_width, prs.slide_height

        neue_kommentare = []
        for a in eingabe.aenderungen:
            x_emu = round(a.x / 100 * breite) if a.x is not None else None
            y_emu = round(a.y / 100 * hoehe) if a.y is not None else None
            neue_kommentare.append(
                NeuerKommentar(
                    folie=a.folie,
                    text=f"[{a.elementLabel}] {a.text}",
                    x_emu=x_emu,
                    y_emu=y_emu,
                )
            )
        try:
            schreibe_kommentare(pfad, neue_kommentare)
        except ValueError as exc:
            raise HTTPException(400, str(exc)) from exc
        return {"geschrieben": len(neue_kommentare)}

    return app
