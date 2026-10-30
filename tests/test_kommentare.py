"""Tests fuer das Kommentar-Auslesen.

python-pptx kann Kommentare nicht selbst anlegen (keine High-Level-API,
siehe praesentationsassistent/kommentare.py). Um das Lesen trotzdem echt zu
testen, bauen wir hier von Hand die zusaetzlichen OOXML-Teile fuer das
aeltere ("legacy") Kommentarformat in eine bestehende .pptx ein: Content-
Types-Override, eine neue Relationship auf der ersten Folie, sowie die
Teile ``ppt/comments/comment1.xml`` und ``ppt/commentAuthors.xml``.

Das modernere Threaded-Comments-Format (PowerPoint 2016+) wird dagegen NICHT
mit einer selbstgebauten Testdatei abgedeckt -- das Schema ist nicht so
zuverlaessig dokumentiert, dass ein handgebautes Beispiel wirklich dem
echten Format entspraeche. Das ist eine bekannte, bewusst akzeptierte Luecke
der Testsuite (siehe README.md); die Leseunterstuetzung dafuer im Code ist
best-effort und erst mit einer echten, in aktuellem PowerPoint erzeugten
Beispieldatei zuverlaessig pruefbar.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

from praesentationsassistent.kommentare import lese_kommentare
from praesentationsassistent.pptx_erzeugen import FolienBeschreibung, erzeuge_praesentation

_COMMENT_AUTHORS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    '<p:cmAuthor id="0" name="TESTDATEN Pruefer" initials="TP" lastIdx="1" clrIdx="0"/>'
    "</p:cmAuthorLst>"
)

_COMMENT1_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    '<p:cm authorId="0" idx="1" dt="2026-01-01T10:00:00.000">'
    '<p:pos x="1000" y="2000"/>'
    "<p:text>TESTDATEN Kommentar: Logo bitte groesser</p:text>"
    "</p:cm>"
    "</p:cmLst>"
)

_CONTENT_TYPE_OVERRIDES = (
    '<Override PartName="/ppt/commentAuthors.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"/>'
    '<Override PartName="/ppt/comments/comment1.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/>'
)

_RELATIONSHIP = (
    '<Relationship Id="rIdTestKommentar" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
    'Target="../comments/comment1.xml"/>'
)


def _baue_pptx_mit_legacy_kommentar(basis_pptx: Path, ziel_pptx: Path) -> None:
    """Kopiert ``basis_pptx`` nach ``ziel_pptx`` und fuegt dabei die OOXML-
    Teile fuer genau einen Legacy-Kommentar auf der ersten Folie ein."""
    with zipfile.ZipFile(basis_pptx) as quelle:
        inhalte = {name: quelle.read(name) for name in quelle.namelist()}

    # 1. Content-Types: die zwei neuen Teile bekannt machen.
    content_types = inhalte["[Content_Types].xml"].decode("utf-8")
    assert "</Types>" in content_types
    content_types = content_types.replace(
        "</Types>", _CONTENT_TYPE_OVERRIDES + "</Types>"
    )
    inhalte["[Content_Types].xml"] = content_types.encode("utf-8")

    # 2. Rels der ersten Folie: Relationship auf den Kommentar-Teil ergaenzen.
    rels_pfad = "ppt/slides/_rels/slide1.xml.rels"
    rels_xml = inhalte[rels_pfad].decode("utf-8")
    assert "</Relationships>" in rels_xml
    rels_xml = rels_xml.replace("</Relationships>", _RELATIONSHIP + "</Relationships>")
    inhalte[rels_pfad] = rels_xml.encode("utf-8")

    # 3. Die neuen Teile selbst.
    inhalte["ppt/commentAuthors.xml"] = _COMMENT_AUTHORS_XML.encode("utf-8")
    inhalte["ppt/comments/comment1.xml"] = _COMMENT1_XML.encode("utf-8")

    with zipfile.ZipFile(ziel_pptx, "w", zipfile.ZIP_DEFLATED) as ziel:
        for name, daten in inhalte.items():
            ziel.writestr(name, daten)


def test_legacy_kommentar_wird_ausgelesen(tmp_path: Path):
    basis = tmp_path / "basis.pptx"
    erzeuge_praesentation(
        [FolienBeschreibung(titel="TESTDATEN Folie 1")], basis
    )

    mit_kommentar = tmp_path / "mit_kommentar.pptx"
    _baue_pptx_mit_legacy_kommentar(basis, mit_kommentar)

    kommentare = lese_kommentare(mit_kommentar)

    assert len(kommentare) == 1
    k = kommentare[0]
    assert k.folie == 1
    assert k.text == "TESTDATEN Kommentar: Logo bitte groesser"
    assert k.autor == "TESTDATEN Pruefer"
    assert k.quelle == "legacy"
    assert k.element_bezug == "x=1000, y=2000"


def test_pptx_ohne_kommentare_liefert_leere_liste(tmp_path: Path):
    ohne_kommentar = tmp_path / "ohne.pptx"
    erzeuge_praesentation([FolienBeschreibung(titel="Ohne Kommentar")], ohne_kommentar)

    assert lese_kommentare(ohne_kommentar) == []
