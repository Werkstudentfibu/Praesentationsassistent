"""Kommandozeile: erzeugen / kommentare-lesen.

Kein LLM-Aufruf, kein Langdock-Anschluss -- siehe praesentationsassistent/langdock.py.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .beispiel_erzeugen import BeispielErzeugenFehler, erzeuge_markenbeispiel
from .kommentare import lese_kommentare
from .langdock import LangdockFehler
from .pptx_erzeugen import FolienBeschreibung, erzeuge_praesentation

# Frei erfundene TESTDATEN fuer den Befehl "erzeugen" -- keine echten
# Markeninhalte. Die eigentliche Dokument- und Markenregel-gestuetzte
# Folienerzeugung folgt spaeter.
_TESTFOLIEN = [
    FolienBeschreibung(
        titel="Testfolie 1 (TESTDATEN)",
        text=(
            "Dies ist ein frei erfundener Beispieltext fuer das Grundgeruest "
            "des Praesentationsassistenten. Kein echter Markeninhalt."
        ),
        bild_pfad=None,
        bild_ordner_hinweis="Bilder/testfolie1",
    ),
    FolienBeschreibung(
        titel="Testfolie 2 (TESTDATEN)",
        text="Zweite Testfolie ohne Bild -- zeigt den Platzhalter.",
        bild_pfad=None,
        bild_ordner_hinweis="Bilder/testfolie2",
    ),
]


def _cmd_erzeugen(args: argparse.Namespace) -> int:
    pfad = erzeuge_praesentation(_TESTFOLIEN, Path(args.ausgabe))
    print(f"Test-Praesentation erzeugt: {pfad}")
    return 0


def _cmd_erzeugen_mit_marke(args: argparse.Namespace) -> int:
    try:
        pfad = erzeuge_markenbeispiel(args.marke, args.thema, Path(args.ausgabe))
    except (LangdockFehler, BeispielErzeugenFehler) as exc:
        print(f"Fehler: {exc}")
        return 1
    print(f"Marken-Beispielpraesentation erzeugt: {pfad}")
    return 0


def _cmd_ueberarbeiten(args: argparse.Namespace) -> int:
    import uvicorn

    from .web.server import erzeuge_app

    app = erzeuge_app(args.ordner)
    print(f"Oeffne im Browser: http://127.0.0.1:{args.port}")
    uvicorn.run(app, host="127.0.0.1", port=args.port, log_level="warning")
    return 0


def _cmd_kommentare_lesen(args: argparse.Namespace) -> int:
    datei = Path(args.datei)
    if not datei.exists():
        print(f"Fehler: Datei nicht gefunden: {datei}")
        return 1
    kommentare = lese_kommentare(datei)
    if not kommentare:
        print("Keine Kommentare gefunden.")
        return 0
    for k in kommentare:
        autor = f" [{k.autor}]" if k.autor else ""
        bezug = f" ({k.element_bezug})" if k.element_bezug else ""
        antwort = "  -> Antwort" if k.ist_antwort else ""
        status = f" [status: {k.status}]" if k.status else ""
        print(
            f"Folie {k.folie}{antwort}{autor}{bezug}: {k.text}  "
            f"<quelle: {k.quelle}>{status}"
        )
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="praesentationsassistent",
        description="Visual-Identity-Praesentationsassistent (Grundgeruest).",
    )
    sub = parser.add_subparsers(dest="befehl", required=True)

    p_erzeugen = sub.add_parser(
        "erzeugen", help="Erzeugt eine Test-Praesentation aus TESTDATEN."
    )
    p_erzeugen.add_argument(
        "--ausgabe",
        default="test_praesentation.pptx",
        help="Ausgabedatei (.pptx), Standard: test_praesentation.pptx",
    )
    p_erzeugen.set_defaults(func=_cmd_erzeugen)

    p_erzeugen_mit_marke = sub.add_parser(
        "erzeugen-mit-marke",
        help=(
            "Erzeugt eine Beispiel-Praesentation mit echten Marken-Farben/"
            "Schriften (braucht LANGDOCK_API_KEY lokal gesetzt)."
        ),
    )
    p_erzeugen_mit_marke.add_argument(
        "marke",
        help="Markenname, z.B. 'URBAN LOFT', 'AMERON', 'Althoff Hotels', 'Althoff Collection'",
    )
    p_erzeugen_mit_marke.add_argument(
        "thema", help="Thema der Beispiel-Praesentation, z.B. 'Neueroeffnung Hotel X'"
    )
    p_erzeugen_mit_marke.add_argument(
        "--ausgabe",
        default="beispiel_mit_marke.pptx",
        help="Ausgabedatei (.pptx), Standard: beispiel_mit_marke.pptx",
    )
    p_erzeugen_mit_marke.set_defaults(func=_cmd_erzeugen_mit_marke)

    p_ueberarbeiten = sub.add_parser(
        "ueberarbeiten",
        help="Startet die lokale Ueberarbeitungs-Oberflaeche (Browser) fuer erzeugte Praesentationen.",
    )
    p_ueberarbeiten.add_argument(
        "--ordner",
        default="ausgabe",
        help="Ordner mit den .pptx-Dateien zum Ueberarbeiten, Standard: ausgabe",
    )
    p_ueberarbeiten.add_argument("--port", type=int, default=8765)
    p_ueberarbeiten.set_defaults(func=_cmd_ueberarbeiten)

    p_kommentare = sub.add_parser(
        "kommentare-lesen", help="Listet gefundene Kommentare aus einer vorhandenen .pptx."
    )
    p_kommentare.add_argument("datei", help="Pfad zur .pptx-Datei")
    p_kommentare.set_defaults(func=_cmd_kommentare_lesen)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
