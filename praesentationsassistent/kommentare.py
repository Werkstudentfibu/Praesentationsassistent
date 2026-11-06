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

1. "Legacy Comments" (vor PowerPoint 2021/Microsoft 365):
   ``ppt/slides/_rels/slideN.xml.rels`` enthaelt eine Relationship vom Typ
   ``http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments``
   auf ``ppt/comments/commentN.xml``. Dort stehen die Kommentare als
   ``<p:cmLst><p:cm authorId=".." idx=".." dt="..">
       <p:pos x=".." y=".."/><p:text>...</p:text>
     </p:cm></p:cmLst>``. Die Autorennamen liegen separat in
   ``ppt/commentAuthors.xml``. Dokumentiert in ECMA-376/ISO-29500.

2. "Moderne Kommentare" (PowerPoint fuer Microsoft 365 / PowerPoint 2021+,
   Standard bei allen aktuellen Installationen). Dokumentiert in Microsofts
   Open-Specification [MS-PPTX] (Namespace ``p188``):

   - Die Folien-Relationship hat den Typ
     ``http://schemas.microsoft.com/office/2018/10/relationships/comments``
     und zeigt auf ``ppt/comments/modernComment_<...>.xml``.
   - Dort: ``<p188:cmLst><p188:cm id="{GUID}" authorId="{GUID}"
     created=".." status="resolved"?> ... <p188:replyLst><p188:reply ..>
     <p188:txBody>..</p188:txBody></p188:reply></p188:replyLst>
     <p188:txBody><a:p><a:r><a:t>Text</a:t></a:r></a:p></p188:txBody>
     </p188:cm></p188:cmLst>`` -- der Text ist also DrawingML-Fliesstext
     (``a:p``/``a:t``), nicht ein einfaches ``text``-Element.
   - Verankerung an einem Folienelement steht in Monikern wie
     ``<ac:deMkLst><pc:docMk/><pc:sldMk ../><ac:spMk id="3" ../></ac:deMkLst>``.
   - Autoren stehen in ``ppt/authors.xml`` (``p188:authorLst``/``p188:author``),
     referenziert aus ``ppt/_rels/presentation.xml.rels`` mit Typ
     ``http://schemas.microsoft.com/office/2018/10/relationships/authors``.

   (Hinweis: Die frueher hier vermuteten "commentThreads" sind das Format
   fuer Excel-Kommentare, nicht fuer PowerPoint.)

Diese Funktion trifft bewusst keine Annahme ueber die Bedeutung der
Kommentare (z.B. wer antworten soll) -- sie liest nur Rohtext, erkennbare
Herkunft und die von PowerPoint selbst gespeicherten Rohdaten (Status
"resolved", Antwort-Beziehung).

Ausserdem (unten, ``schreibe_kommentare``): Kommentare im Legacy-Format in
eine vorhandene .pptx SCHREIBEN -- genutzt von der Ueberarbeitungs-
Oberflaeche (web/). PowerPoint 365 zeigt diese ganz normal an.
"""

from __future__ import annotations

import posixpath
import re
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

from pptx import Presentation

_NS_LEGACY = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_PACKAGE_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS_P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
_NS_DRAWINGML = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"

_LEGACY_COMMENTS_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
_MODERN_COMMENTS_REL_TYPE = (
    "http://schemas.microsoft.com/office/2018/10/relationships/comments"
)
_MODERN_AUTHORS_REL_TYPE = (
    "http://schemas.microsoft.com/office/2018/10/relationships/authors"
)

_COMMENTS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.comments+xml"
)
_AUTOR_NAME = "Praesentationsassistent (Ueberarbeitungs-Oberflaeche)"
_AUTOR_INITIALEN = "PA"

# Moniker, die nur Dokument/Folie bezeichnen -- kein konkretes Element.
_MONIKER_OHNE_ELEMENTBEZUG = {"docMk", "sldMk", "sldMkLst", "deMkLst", "txMkLst"}


@dataclass
class Kommentar:
    """Ein einzelner, roh ausgelesener Kommentar."""

    folie: int  # 1-basierte Foliennummer in Praesentations-Reihenfolge
    text: str
    element_bezug: str | None = None  # Positions-/Element-Bezug, falls erkennbar
    autor: str | None = None
    quelle: str = "legacy"  # "legacy" oder "modern" (siehe Modul-Docstring)
    # Nur modernes Format: Rohwert des status-Attributs (z.B. "resolved"),
    # None = nicht gesetzt (in PowerPoint "offen").
    status: str | None = None
    # Nur modernes Format: True, wenn dies eine Antwort im Thread ist.
    ist_antwort: bool = False


def _lokaler_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _rels_pfad(teil_pfad: str) -> str:
    ordner, datei = teil_pfad.rsplit("/", 1)
    return f"{ordner}/_rels/{datei}.rels"


def _ziel_aufloesen(teil_pfad: str, ziel: str) -> str:
    """Loest ein relatives Relationship-Target (z.B. "../comments/comment1.xml")
    gegen den Ordner von ``teil_pfad`` auf. Absolute Targets ("/ppt/...")
    werden relativ zur Paketwurzel interpretiert."""
    if ziel.startswith("/"):
        return posixpath.normpath(ziel.lstrip("/"))
    ordner = teil_pfad.rsplit("/", 1)[0]
    return posixpath.normpath(posixpath.join(ordner, ziel))


def _relationships(zf: zipfile.ZipFile, teil_pfad: str) -> list[ET.Element]:
    rels_pfad = _rels_pfad(teil_pfad)
    if rels_pfad not in zf.namelist():
        return []
    rels_baum = ET.fromstring(zf.read(rels_pfad))
    return rels_baum.findall(f"{{{_NS_PACKAGE_RELS}}}Relationship")


# --------------------------------------------------------------------------
# Legacy-Format
# --------------------------------------------------------------------------


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
    ergebnisse: list[Kommentar] = []
    for rel in _relationships(zf, slide_pfad):
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


# --------------------------------------------------------------------------
# Modernes Format (p188)
# --------------------------------------------------------------------------


def _lade_autoren_modern(zf: zipfile.ZipFile) -> dict[str, str]:
    """Liest ``ppt/authors.xml`` (bzw. das ueber die Praesentations-
    Relationship referenzierte Autoren-Teil). Schluessel ist die Autor-GUID."""
    pfade: list[str] = []
    for rel in _relationships(zf, "ppt/presentation.xml"):
        if rel.get("Type") == _MODERN_AUTHORS_REL_TYPE:
            pfade.append(_ziel_aufloesen("ppt/presentation.xml", rel.get("Target", "")))
    if not pfade:
        pfade = ["ppt/authors.xml"]

    autoren: dict[str, str] = {}
    for pfad in pfade:
        if pfad not in zf.namelist():
            continue
        baum = ET.fromstring(zf.read(pfad))
        for autor in baum.iter(f"{{{_NS_P188}}}author"):
            autoren[autor.get("id", "")] = autor.get("name", "")
    return autoren


def _txbody_text(element: ET.Element) -> str:
    """Setzt den Text eines ``p188:txBody`` zusammen: Absaetze (``a:p``)
    werden mit Zeilenumbruch getrennt, Text-Laeufe (``a:t``) aneinander-
    gehaengt, ``a:br`` wird zum Zeilenumbruch."""
    tx_body = element.find(f"{{{_NS_P188}}}txBody")
    if tx_body is None:
        return ""
    absaetze: list[str] = []
    for absatz in tx_body.findall(f"{{{_NS_DRAWINGML}}}p"):
        teile: list[str] = []
        for knoten in absatz.iter():
            name = _lokaler_name(knoten.tag)
            if name == "t":
                teile.append(knoten.text or "")
            elif name == "br":
                teile.append("\n")
        absaetze.append("".join(teile))
    return "\n".join(absaetze).strip()


def _element_bezug_modern(cm: ET.Element) -> str | None:
    """Beschreibt die Verankerung eines modernen Kommentars: Folienelement
    (aus den Monikern, z.B. ``spMk id="3"`` -> "spMk id=3") und/oder Position."""
    teile: list[str] = []
    for knoten in cm.iter():
        name = _lokaler_name(knoten.tag)
        if (
            name.endswith("Mk")
            and name not in _MONIKER_OHNE_ELEMENTBEZUG
            and knoten.get("id") is not None
        ):
            teile.append(f"{name} id={knoten.get('id')}")
    pos_el = cm.find(f"{{{_NS_P188}}}pos")
    if pos_el is not None:
        teile.append(f"x={pos_el.get('x')}, y={pos_el.get('y')}")
    return "; ".join(teile) if teile else None


def _lese_moderne_kommentare(
    zf: zipfile.ZipFile, slide_pfad: str, slide_nummer: int, autoren: dict[str, str]
) -> list[Kommentar]:
    ergebnisse: list[Kommentar] = []
    for rel in _relationships(zf, slide_pfad):
        if rel.get("Type") != _MODERN_COMMENTS_REL_TYPE:
            continue
        teil_pfad = _ziel_aufloesen(slide_pfad, rel.get("Target", ""))
        if teil_pfad not in zf.namelist():
            continue
        baum = ET.fromstring(zf.read(teil_pfad))
        for cm in baum.findall(f"{{{_NS_P188}}}cm"):
            status = cm.get("status")
            bezug = _element_bezug_modern(cm)
            ergebnisse.append(
                Kommentar(
                    folie=slide_nummer,
                    text=_txbody_text(cm),
                    element_bezug=bezug,
                    autor=autoren.get(cm.get("authorId", "")),
                    quelle="modern",
                    status=status,
                )
            )
            reply_lst = cm.find(f"{{{_NS_P188}}}replyLst")
            if reply_lst is None:
                continue
            for reply in reply_lst.findall(f"{{{_NS_P188}}}reply"):
                ergebnisse.append(
                    Kommentar(
                        folie=slide_nummer,
                        text=_txbody_text(reply),
                        element_bezug=bezug,
                        autor=autoren.get(reply.get("authorId", "")),
                        quelle="modern",
                        status=status,
                        ist_antwort=True,
                    )
                )
    return ergebnisse


# --------------------------------------------------------------------------
# Kommentare LESEN -- oeffentliche Funktion
# --------------------------------------------------------------------------


def lese_kommentare(pptx_pfad: Path | str) -> list[Kommentar]:
    """Liest alle Kommentare aus einer vorhandenen .pptx-Datei.

    Deckt das aeltere ("legacy") und das moderne Kommentarformat ab (siehe
    Modul-Docstring). Antworten im modernen Format werden als eigene
    Eintraege direkt nach ihrem Ausgangskommentar geliefert
    (``ist_antwort=True``).

    Die Foliennummer folgt der Reihenfolge, in der die Folien in der
    Praesentation tatsaechlich erscheinen (ueber python-pptx aufgeloest) --
    nicht der (ggf. abweichenden) Nummerierung der internen slideN.xml-
    Dateinamen.
    """
    pptx_pfad = Path(pptx_pfad)
    prs = Presentation(str(pptx_pfad))

    ergebnisse: list[Kommentar] = []
    with zipfile.ZipFile(pptx_pfad) as zf:
        autoren_legacy = _lade_autoren_legacy(zf)
        autoren_modern = _lade_autoren_modern(zf)
        for index, slide in enumerate(prs.slides, start=1):
            slide_pfad = str(slide.part.partname).lstrip("/")
            ergebnisse.extend(
                _lese_legacy_kommentare(zf, slide_pfad, index, autoren_legacy)
            )
            ergebnisse.extend(
                _lese_moderne_kommentare(zf, slide_pfad, index, autoren_modern)
            )
    return ergebnisse


# ---------------------------------------------------------------------------
# Kommentare SCHREIBEN (legacy-Format) -- fuer die Ueberarbeitungs-
# Oberflaeche (web/): eine echte, in PowerPoint sichtbare Rueckmeldung statt
# einer separaten, nicht-standardkonformen Begleitdatei.
# ---------------------------------------------------------------------------


@dataclass
class NeuerKommentar:
    """Ein neu anzulegender Kommentar (von einem Aufrufer wie der
    Ueberarbeitungs-Oberflaeche uebergeben)."""

    folie: int  # 1-basiert, wie Kommentar.folie
    text: str
    element_bezug: str | None = None
    x_emu: int | None = None
    y_emu: int | None = None


def _naechste_rid(bestehende_ids: list[str]) -> str:
    zahlen = [int(m.group(1)) for i in bestehende_ids if (m := re.fullmatch(r"rId(\d+)", i))]
    return f"rId{max(zahlen, default=0) + 1}"


def _xml_parsen(inhalte: dict[str, bytes], pfad: str) -> ET.Element:
    return ET.fromstring(inhalte[pfad])


def _xml_schreiben(inhalte: dict[str, bytes], pfad: str, baum: ET.Element) -> None:
    inhalte[pfad] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + ET.tostring(
        baum, encoding="unicode"
    ).encode("utf-8")


def _content_type_sicherstellen(inhalte: dict[str, bytes], partname: str, content_type: str) -> None:
    pfad = "[Content_Types].xml"
    baum = _xml_parsen(inhalte, pfad)
    for override in baum.findall(f"{{{_NS_CONTENT_TYPES}}}Override"):
        if override.get("PartName") == partname:
            return
    override = ET.SubElement(baum, f"{{{_NS_CONTENT_TYPES}}}Override")
    override.set("PartName", partname)
    override.set("ContentType", content_type)
    _xml_schreiben(inhalte, pfad, baum)


def _autor_id_sicherstellen(inhalte: dict[str, bytes], namen: set[str]) -> str:
    """Sorgt dafuer, dass ``ppt/commentAuthors.xml`` existiert und einen
    Eintrag fuer diese Oberflaeche hat; gibt dessen ``id`` zurueck."""
    pfad = "ppt/commentAuthors.xml"
    neu = pfad not in namen
    baum = (
        ET.Element(f"{{{_NS_LEGACY}}}cmAuthorLst")
        if neu
        else _xml_parsen(inhalte, pfad)
    )

    for autor in baum.findall(f"{{{_NS_LEGACY}}}cmAuthor"):
        if autor.get("name") == _AUTOR_NAME:
            return autor.get("id", "0")

    vorhandene_ids = [int(a.get("id", "0")) for a in baum.findall(f"{{{_NS_LEGACY}}}cmAuthor")]
    autor_id = str(max(vorhandene_ids, default=-1) + 1)
    neuer_autor = ET.SubElement(baum, f"{{{_NS_LEGACY}}}cmAuthor")
    neuer_autor.set("id", autor_id)
    neuer_autor.set("name", _AUTOR_NAME)
    neuer_autor.set("initials", _AUTOR_INITIALEN)
    neuer_autor.set("lastIdx", "0")
    neuer_autor.set("clrIdx", "0")
    _xml_schreiben(inhalte, pfad, baum)

    if neu:
        namen.add(pfad)
        _content_type_sicherstellen(
            inhalte,
            "/ppt/commentAuthors.xml",
            "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml",
        )
        praes_rels_pfad = "ppt/_rels/presentation.xml.rels"
        rels_baum = _xml_parsen(inhalte, praes_rels_pfad)
        bestehende_ids = [
            r.get("Id", "") for r in rels_baum.findall(f"{{{_NS_PACKAGE_RELS}}}Relationship")
        ]
        rel = ET.SubElement(rels_baum, f"{{{_NS_PACKAGE_RELS}}}Relationship")
        rel.set("Id", _naechste_rid(bestehende_ids))
        rel.set(
            "Type",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors",
        )
        rel.set("Target", "commentAuthors.xml")
        _xml_schreiben(inhalte, praes_rels_pfad, rels_baum)

    return autor_id


def _naechster_freier_kommentarteil(namen: set[str]) -> str:
    n = 1
    while f"ppt/comments/comment{n}.xml" in namen:
        n += 1
    return f"ppt/comments/comment{n}.xml"


def _naechster_freier_idx(inhalte: dict[str, bytes], namen: set[str]) -> int:
    hoechster = 0
    for name in namen:
        if not re.fullmatch(r"ppt/comments/comment\d+\.xml", name):
            continue
        baum = _xml_parsen(inhalte, name)
        for cm in baum.findall(f"{{{_NS_LEGACY}}}cm"):
            hoechster = max(hoechster, int(cm.get("idx", "0")))
    return hoechster + 1


def _kommentarteil_fuer_slide_sicherstellen(
    inhalte: dict[str, bytes], namen: set[str], slide_pfad: str
) -> str:
    """Gibt den Teil-Pfad der (ggf. neu angelegten) ``commentN.xml`` fuer
    diese Folie zurueck -- wiederverwendet einen vorhandenen, falls die
    Folie schon Kommentare hat."""
    rels_pfad = _rels_pfad(slide_pfad)
    if rels_pfad in namen:
        rels_baum = _xml_parsen(inhalte, rels_pfad)
        for rel in rels_baum.findall(f"{{{_NS_PACKAGE_RELS}}}Relationship"):
            if rel.get("Type") == _LEGACY_COMMENTS_REL_TYPE:
                return _ziel_aufloesen(slide_pfad, rel.get("Target", ""))
    else:
        rels_baum = ET.Element(f"{{{_NS_PACKAGE_RELS}}}Relationships")

    kommentarteil_pfad = _naechster_freier_kommentarteil(namen)
    namen.add(kommentarteil_pfad)
    leerer_baum = ET.Element(f"{{{_NS_LEGACY}}}cmLst")
    _xml_schreiben(inhalte, kommentarteil_pfad, leerer_baum)
    _content_type_sicherstellen(
        inhalte, "/" + kommentarteil_pfad, _COMMENTS_CONTENT_TYPE
    )

    bestehende_ids = [
        r.get("Id", "") for r in rels_baum.findall(f"{{{_NS_PACKAGE_RELS}}}Relationship")
    ]
    rel = ET.SubElement(rels_baum, f"{{{_NS_PACKAGE_RELS}}}Relationship")
    rel.set("Id", _naechste_rid(bestehende_ids))
    rel.set("Type", _LEGACY_COMMENTS_REL_TYPE)
    ziel = posixpath.relpath(kommentarteil_pfad, slide_pfad.rsplit("/", 1)[0])
    rel.set("Target", ziel)
    namen.add(rels_pfad)
    _xml_schreiben(inhalte, rels_pfad, rels_baum)

    return kommentarteil_pfad


def schreibe_kommentare(
    pptx_pfad: Path | str, neue_kommentare: list[NeuerKommentar]
) -> None:
    """Haengt ``neue_kommentare`` als echte (legacy) PowerPoint-Kommentare an
    die vorhandene .pptx an -- in PowerPoint normal sichtbar und mit
    ``lese_kommentare`` wieder auslesbar. Aendert nur die Kommentar-
    bezogenen Teile des Pakets, alles andere (Folien, Bilder, Layouts, ...)
    bleibt unveraendert erhalten.
    """
    if not neue_kommentare:
        return

    pptx_pfad = Path(pptx_pfad)
    prs = Presentation(str(pptx_pfad))
    slide_pfade = {i: str(s.part.partname).lstrip("/") for i, s in enumerate(prs.slides, start=1)}

    ET.register_namespace("", _NS_PACKAGE_RELS)
    ET.register_namespace("p", _NS_LEGACY)

    with zipfile.ZipFile(pptx_pfad) as zf:
        namen = set(zf.namelist())
        inhalte: dict[str, bytes] = {n: zf.read(n) for n in zf.namelist()}
        info_liste = zf.infolist()

    autor_id = _autor_id_sicherstellen(inhalte, namen)
    naechster_idx = _naechster_freier_idx(inhalte, namen)
    jetzt = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"

    kommentarteil_je_folie: dict[int, str] = {}
    for eintrag in neue_kommentare:
        if eintrag.folie not in slide_pfade:
            raise ValueError(f"Keine Folie Nr. {eintrag.folie} in dieser Praesentation.")
        if eintrag.folie not in kommentarteil_je_folie:
            kommentarteil_je_folie[eintrag.folie] = _kommentarteil_fuer_slide_sicherstellen(
                inhalte, namen, slide_pfade[eintrag.folie]
            )
        kommentarteil_pfad = kommentarteil_je_folie[eintrag.folie]

        baum = _xml_parsen(inhalte, kommentarteil_pfad)
        cm = ET.SubElement(baum, f"{{{_NS_LEGACY}}}cm")
        cm.set("authorId", autor_id)
        cm.set("idx", str(naechster_idx))
        cm.set("dt", jetzt)
        naechster_idx += 1
        if eintrag.x_emu is not None and eintrag.y_emu is not None:
            pos = ET.SubElement(cm, f"{{{_NS_LEGACY}}}pos")
            pos.set("x", str(eintrag.x_emu))
            pos.set("y", str(eintrag.y_emu))
        text_el = ET.SubElement(cm, f"{{{_NS_LEGACY}}}text")
        text_el.text = eintrag.text
        _xml_schreiben(inhalte, kommentarteil_pfad, baum)

    neuer_pfad = pptx_pfad.with_suffix(".tmp_kommentare.pptx")
    with zipfile.ZipFile(neuer_pfad, "w", zipfile.ZIP_DEFLATED) as zf_neu:
        geschrieben: set[str] = set()
        for info in info_liste:
            if info.filename in geschrieben:
                continue
            geschrieben.add(info.filename)
            zf_neu.writestr(info, inhalte.get(info.filename, b""))
        for pfad, daten in inhalte.items():
            if pfad not in geschrieben:
                zf_neu.writestr(pfad, daten)
    neuer_pfad.replace(pptx_pfad)
