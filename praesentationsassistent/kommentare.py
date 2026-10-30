"""Kommentare aus bestehenden .pptx-Dateien auslesen (Grundgeruest).

Rechercheergebnis (Stand 2026, python-pptx 1.0.2): python-pptx bietet KEINE
High-Level-API fuer Folien-Kommentare.

  - ``slide.notes_slide`` / ``slide.has_notes_slide`` sind die Referenten-
    NOTIZEN unterhalb der Folie (python-pptx-Begriff "Notes Slide") -- das
    ist etwas anderes als Kommentare und hat mit dem hier gesuchten
    Feedback-Workflow nichts zu tun.
  - ``presentation.core_properties.comments`` ist nur ein einzelnes
    Freitextfeld in den Dokument-Metadaten (wie "Autor" oder "Thema"),
    keine Liste von Kommentaren pro Folie.

Echte Folien-Kommentare stecken ausschliesslich im zugrunde liegenden OOXML-
Paket (eine .pptx ist ein ZIP-Container). Es gibt dabei ZWEI verschiedene,
inkompatible Formate, je nachdem mit welcher PowerPoint-Version die
Kommentare angelegt wurden:

1. "Legacy Comments" (vor PowerPoint 2016): ``ppt/slides/_rels/slideN.xml.rels``
   enthaelt eine Relationship vom Typ
   ``http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments``
   auf ``ppt/comments/commentN.xml``. Dort stehen die Kommentare als
   ``<p:cmLst><p:cm authorId=".." idx=".." dt="..">
       <p:pos x=".." y=".."/><p:text>...</p:text>
     </p:cm></p:cmLst>``. Die Autorennamen liegen separat in
   ``ppt/commentAuthors.xml``. Dieses Format ist gut dokumentiert (ECMA-376/
   ISO-29500) und wird unten vollstaendig unterstuetzt + getestet.

2. "Moderne/Threaded Comments" (PowerPoint 2016+ / Microsoft 365): liegen
   unter ``ppt/commentThreads/commentThreadN.xml`` bzw.
   ``ppt/commentAuthors/commentAuthorsLst.xml`` und werden ueber eine
   Relationship referenziert, deren Typ "commentThread" enthaelt. Das Schema
   ist deutlich schlechter oeffentlich dokumentiert als das aeltere Format
   und wurde von Microsoft in den letzten Jahren mehrfach erweitert. Wir
   lesen es daher nur BEST-EFFORT (generisches Einsammeln aller Text-Knoten
   unterhalb einer passenden Relationship) -- das wurde NICHT gegen eine
   echte, in aktuellem PowerPoint/Microsoft 365 erzeugte Beispieldatei
   getestet. Bekannte Luecke, siehe README.md und Testsuite.

Diese Funktion trifft bewusst keine Annahme ueber die Bedeutung der
Kommentare (z.B. "erledigt"/"offen", wer antworten soll) -- sie liest nur
Rohtext + erkennbare Herkunft.
"""

from __future__ import annotations

import posixpath
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from pptx import Presentation

_NS_LEGACY = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_PACKAGE_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
_LEGACY_COMMENTS_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)


@dataclass
class Kommentar:
    """Ein einzelner, roh ausgelesener Kommentar."""

    folie: int  # 1-basierte Foliennummer in Praesentations-Reihenfolge
    text: str
    element_bezug: str | None = None  # Positions-/Element-Bezug, falls erkennbar
    autor: str | None = None
    quelle: str = "legacy"  # "legacy" oder "modern" (siehe Modul-Docstring)


def _rels_pfad(teil_pfad: str) -> str:
    ordner, datei = teil_pfad.rsplit("/", 1)
    return f"{ordner}/_rels/{datei}.rels"


def _ziel_aufloesen(teil_pfad: str, ziel: str) -> str:
    """Loest ein relatives Relationship-Target (z.B. "../comments/comment1.xml")
    gegen den Ordner von ``teil_pfad`` auf."""
    ordner = teil_pfad.rsplit("/", 1)[0]
    return posixpath.normpath(posixpath.join(ordner, ziel))


def _lade_autoren_legacy(zf: zipfile.ZipFile) -> dict[str, str]:
    pfad = "ppt/commentAuthors.xml"
    if pfad not in zf.namelist():
        return {}
    baum = ET.fromstring(zf.read(pfad))
    autoren: dict[str, str] = {}
    for autor in baum.findall(f"{{{_NS_LEGACY}}}cmAuthor"):
        autoren[autor.get("id", "")] = autor.get("name", "")
    return autoren


def _lese_legacy_kommentare(
    zf: zipfile.ZipFile, slide_pfad: str, slide_nummer: int, autoren: dict[str, str]
) -> list[Kommentar]:
    rels_pfad = _rels_pfad(slide_pfad)
    if rels_pfad not in zf.namelist():
        return []

    rels_baum = ET.fromstring(zf.read(rels_pfad))
    ergebnisse: list[Kommentar] = []
    for rel in rels_baum.findall(f"{{{_NS_PACKAGE_RELS}}}Relationship"):
        if rel.get("Type") != _LEGACY_COMMENTS_REL_TYPE:
            continue
        teil_pfad = _ziel_aufloesen(slide_pfad, rel.get("Target", ""))
        if teil_pfad not in zf.namelist():
            continue
        kommentar_baum = ET.fromstring(zf.read(teil_pfad))
        for cm in kommentar_baum.findall(f"{{{_NS_LEGACY}}}cm"):
            text_el = cm.find(f"{{{_NS_LEGACY}}}text")
            text = (text_el.text or "") if text_el is not None else ""
            pos_el = cm.find(f"{{{_NS_LEGACY}}}pos")
            bezug = (
                f"x={pos_el.get('x')}, y={pos_el.get('y')}"
                if pos_el is not None
                else None
            )
            autor_id = cm.get("authorId")
            ergebnisse.append(
                Kommentar(
                    folie=slide_nummer,
                    text=text,
                    element_bezug=bezug,
                    autor=autoren.get(autor_id) if autor_id is not None else None,
                    quelle="legacy",
                )
            )
    return ergebnisse


def _lese_moderne_kommentare(
    zf: zipfile.ZipFile, slide_pfad: str, slide_nummer: int
) -> list[Kommentar]:
    """Best-effort fuer Threaded Comments (PowerPoint 2016+). Siehe
    Modul-Docstring: Schema weniger sicher dokumentiert, daher generisches
    Einsammeln von Textknoten statt fest erwarteter Elementnamen."""
    rels_pfad = _rels_pfad(slide_pfad)
    if rels_pfad not in zf.namelist():
        return []

    rels_baum = ET.fromstring(zf.read(rels_pfad))
    ergebnisse: list[Kommentar] = []
    for rel in rels_baum.findall(f"{{{_NS_PACKAGE_RELS}}}Relationship"):
        typ = rel.get("Type", "")
        if "commentThread" not in typ and "threadedcomment" not in typ.lower():
            continue
        teil_pfad = _ziel_aufloesen(slide_pfad, rel.get("Target", ""))
        if teil_pfad not in zf.namelist():
            continue
        teil_baum = ET.fromstring(zf.read(teil_pfad))
        for knoten in teil_baum.iter():
            lokaler_tag = knoten.tag.rsplit("}", 1)[-1]
            if lokaler_tag == "text" and (knoten.text or "").strip():
                ergebnisse.append(
                    Kommentar(
                        folie=slide_nummer,
                        text=knoten.text.strip(),
                        element_bezug=None,
                        autor=None,
                        quelle="modern",
                    )
                )
    return ergebnisse


def lese_kommentare(pptx_pfad: Path | str) -> list[Kommentar]:
    """Liest alle Kommentare aus einer vorhandenen .pptx-Datei.

    Deckt das aeltere ("legacy") Format vollstaendig ab und das modernere
    Threaded-Comments-Format best-effort (siehe Modul-Docstring fuer
    Einschraenkungen). Trifft keine Annahme ueber die Bedeutung der
    Kommentare, liefert nur Rohtext + erkennbare Herkunft.

    Die Foliennummer folgt der Reihenfolge, in der die Folien in der
    Praesentation tatsaechlich erscheinen (ueber python-pptx aufgeloest) --
    nicht der (ggf. abweichenden) Nummerierung der internen slideN.xml-
    Dateinamen.
    """
    pptx_pfad = Path(pptx_pfad)
    prs = Presentation(str(pptx_pfad))

    ergebnisse: list[Kommentar] = []
    with zipfile.ZipFile(pptx_pfad) as zf:
        autoren = _lade_autoren_legacy(zf)
        for index, slide in enumerate(prs.slides, start=1):
            slide_pfad = str(slide.part.partname).lstrip("/")
            ergebnisse.extend(
                _lese_legacy_kommentare(zf, slide_pfad, index, autoren)
            )
            ergebnisse.extend(_lese_moderne_kommentare(zf, slide_pfad, index))
    return ergebnisse
