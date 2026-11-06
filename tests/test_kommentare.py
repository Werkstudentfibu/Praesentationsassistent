"""Tests fuer das Kommentar-Auslesen.

python-pptx kann Kommentare nicht selbst anlegen (keine High-Level-API,
siehe praesentationsassistent/kommentare.py). Um das Lesen trotzdem echt zu
testen, bauen wir hier von Hand die zusaetzlichen OOXML-Teile fuer das
aeltere ("legacy") Kommentarformat in eine bestehende .pptx ein: Content-
Types-Override, eine neue Relationship auf der ersten Folie, sowie die
Teile ``ppt/comments/comment1.xml`` und ``ppt/commentAuthors.xml``.

Fuer das moderne Kommentarformat (PowerPoint fuer Microsoft 365, Namespace
``p188``, siehe [MS-PPTX]) gibt es ebenfalls eine handgebaute Testdatei nach
dem dokumentierten Schema: ``ppt/comments/modernComment_*.xml`` mit Antwort
und Element-Verankerung, ``ppt/authors.xml`` und die zugehoerigen
Relationships. Das ersetzt NICHT den Abgleich mit einer echten, in
PowerPoint 365 kommentierten Datei -- der steht noch aus (siehe README.md).
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


# --------------------------------------------------------------------------
# Modernes Format (PowerPoint fuer Microsoft 365)
# --------------------------------------------------------------------------

_NS_P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
_NS_AC = "http://schemas.microsoft.com/office/drawing/2013/main/command"
_NS_PC = "http://schemas.microsoft.com/office/powerpoint/2013/main/command"
_AUTOR_1 = "{11111111-1111-1111-1111-111111111111}"
_AUTOR_2 = "{22222222-2222-2222-2222-222222222222}"

_AUTHORS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<p188:authorLst xmlns:p188="{_NS_P188}">'
    f'<p188:author id="{_AUTOR_1}" name="TESTDATEN Pruefer" initials="TP" '
    'userId="TESTDATEN Pruefer" providerId="None"/>'
    f'<p188:author id="{_AUTOR_2}" name="TESTDATEN Antwort" initials="TA" '
    'userId="TESTDATEN Antwort" providerId="None"/>'
    "</p188:authorLst>"
)

_MODERN_COMMENT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p188:cmLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    f'xmlns:p188="{_NS_P188}">'
    # Kommentar 1: an einer Form verankert, mit Antwort, zwei Absaetze
    f'<p188:cm id="{{AAAAAAAA-0000-0000-0000-000000000001}}" authorId="{_AUTOR_1}" '
    'created="2026-01-01T10:00:00.000">'
    f'<ac:deMkLst xmlns:ac="{_NS_AC}">'
    f'<pc:docMk xmlns:pc="{_NS_PC}"/>'
    f'<pc:sldMk xmlns:pc="{_NS_PC}" cId="0" sldId="257"/>'
    '<ac:spMk id="3" creationId="{BBBBBBBB-0000-0000-0000-000000000001}"/>'
    "</ac:deMkLst>"
    "<p188:replyLst>"
    f'<p188:reply id="{{AAAAAAAA-0000-0000-0000-000000000002}}" authorId="{_AUTOR_2}" '
    'created="2026-01-01T11:00:00.000">'
    "<p188:txBody><a:bodyPr/><a:lstStyle/>"
    "<a:p><a:r><a:t>TESTDATEN Antwort: erledigt</a:t></a:r></a:p>"
    "</p188:txBody></p188:reply>"
    "</p188:replyLst>"
    "<p188:txBody><a:bodyPr/><a:lstStyle/>"
    "<a:p><a:r><a:t>TESTDATEN Kommentar: Logo </a:t></a:r>"
    "<a:r><a:t>bitte groesser</a:t></a:r></a:p>"
    "<a:p><a:r><a:t>Zweite Zeile</a:t></a:r></a:p>"
    "</p188:txBody>"
    "</p188:cm>"
    # Kommentar 2: per Position auf der Folie, als erledigt markiert
    f'<p188:cm id="{{AAAAAAAA-0000-0000-0000-000000000003}}" authorId="{_AUTOR_2}" '
    'created="2026-01-02T09:00:00.000" status="resolved">'
    f'<pc:sldMkLst xmlns:pc="{_NS_PC}"><pc:docMk/><pc:sldMk cId="0" sldId="257"/>'
    "</pc:sldMkLst>"
    '<p188:pos x="500" y="700"/>'
    "<p188:txBody><a:bodyPr/><a:lstStyle/>"
    "<a:p><a:r><a:t>TESTDATEN Farbe pruefen</a:t></a:r></a:p>"
    "</p188:txBody>"
    "</p188:cm>"
    "</p188:cmLst>"
)

_MODERN_KOMMENTAR_TEIL = "ppt/comments/modernComment_101_TEST0001.xml"


def _baue_pptx_mit_modernem_kommentar(basis_pptx: Path, ziel_pptx: Path) -> None:
    """Wie ``_baue_pptx_mit_legacy_kommentar``, aber mit den OOXML-Teilen des
    modernen Kommentarformats auf der ZWEITEN Folie."""
    with zipfile.ZipFile(basis_pptx) as quelle:
        inhalte = {name: quelle.read(name) for name in quelle.namelist()}

    def _ergaenze(teil: str, ende_tag: str, einschub: str) -> None:
        xml = inhalte[teil].decode("utf-8")
        assert ende_tag in xml
        inhalte[teil] = xml.replace(ende_tag, einschub + ende_tag).encode("utf-8")

    _ergaenze(
        "[Content_Types].xml",
        "</Types>",
        '<Override PartName="/ppt/authors.xml" '
        'ContentType="application/vnd.ms-powerpoint.authors+xml"/>'
        f'<Override PartName="/{_MODERN_KOMMENTAR_TEIL}" '
        'ContentType="application/vnd.ms-powerpoint.comments+xml"/>',
    )
    _ergaenze(
        "ppt/_rels/presentation.xml.rels",
        "</Relationships>",
        '<Relationship Id="rIdTestAutoren" '
        'Type="http://schemas.microsoft.com/office/2018/10/relationships/authors" '
        'Target="authors.xml"/>',
    )
    _ergaenze(
        "ppt/slides/_rels/slide2.xml.rels",
        "</Relationships>",
        '<Relationship Id="rIdTestModern" '
        'Type="http://schemas.microsoft.com/office/2018/10/relationships/comments" '
        'Target="../comments/modernComment_101_TEST0001.xml"/>',
    )

    inhalte["ppt/authors.xml"] = _AUTHORS_XML.encode("utf-8")
    inhalte[_MODERN_KOMMENTAR_TEIL] = _MODERN_COMMENT_XML.encode("utf-8")

    with zipfile.ZipFile(ziel_pptx, "w", zipfile.ZIP_DEFLATED) as ziel:
        for name, daten in inhalte.items():
            ziel.writestr(name, daten)


def test_moderner_kommentar_wird_ausgelesen(tmp_path: Path):
    basis = tmp_path / "basis.pptx"
    erzeuge_praesentation(
        [
            FolienBeschreibung(titel="TESTDATEN Folie 1"),
            FolienBeschreibung(titel="TESTDATEN Folie 2"),
        ],
        basis,
    )
    mit_kommentar = tmp_path / "modern.pptx"
    _baue_pptx_mit_modernem_kommentar(basis, mit_kommentar)

    kommentare = lese_kommentare(mit_kommentar)

    assert len(kommentare) == 3
    haupt, antwort, erledigt = kommentare

    assert haupt.folie == 2
    assert haupt.quelle == "modern"
    assert haupt.text == "TESTDATEN Kommentar: Logo bitte groesser\nZweite Zeile"
    assert haupt.autor == "TESTDATEN Pruefer"
    assert haupt.element_bezug == "spMk id=3"
    assert haupt.status is None
    assert haupt.ist_antwort is False

    assert antwort.folie == 2
    assert antwort.text == "TESTDATEN Antwort: erledigt"
    assert antwort.autor == "TESTDATEN Antwort"
    assert antwort.ist_antwort is True
    assert antwort.element_bezug == "spMk id=3"

    assert erledigt.folie == 2
    assert erledigt.text == "TESTDATEN Farbe pruefen"
    assert erledigt.autor == "TESTDATEN Antwort"
    assert erledigt.element_bezug == "x=500, y=700"
    assert erledigt.status == "resolved"
    assert erledigt.ist_antwort is False
