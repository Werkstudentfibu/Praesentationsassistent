"""Tests fuer die Langdock-Anbindung -- komplett gemockt, kein echter
Netzwerkzugriff und kein echter API-Key in der Testsuite."""

from __future__ import annotations

import json
from unittest.mock import MagicMock, patch

import pytest

from praesentationsassistent.langdock import LangdockFehler, generiere_text


def _antwort(inhalt: str) -> MagicMock:
    mock = MagicMock()
    mock.read.return_value = json.dumps(
        {"choices": [{"message": {"content": inhalt}}]}
    ).encode("utf-8")
    mock.__enter__.return_value = mock
    mock.__exit__.return_value = False
    return mock


def test_ohne_api_key_bricht_klar_ab(monkeypatch):
    monkeypatch.delenv("LANGDOCK_API_KEY", raising=False)
    with pytest.raises(LangdockFehler, match="LANGDOCK_API_KEY"):
        generiere_text("Testprompt")


def test_erfolgreicher_aufruf_liefert_text(monkeypatch):
    monkeypatch.setenv("LANGDOCK_API_KEY", "TESTDATEN-key-nicht-echt")
    with patch("praesentationsassistent.langdock.urllib.request.urlopen") as m:
        m.return_value = _antwort("TESTDATEN-Antworttext")
        ergebnis = generiere_text("Testprompt")
    assert ergebnis == "TESTDATEN-Antworttext"


def test_request_enthaelt_bearer_token_und_body(monkeypatch):
    monkeypatch.setenv("LANGDOCK_API_KEY", "TESTDATEN-key-nicht-echt")
    with patch("praesentationsassistent.langdock.urllib.request.urlopen") as m:
        m.return_value = _antwort("egal")
        generiere_text("Mein Prompt", system="Mein System-Prompt")

    gesendete_anfrage = m.call_args[0][0]
    assert gesendete_anfrage.get_header("Authorization") == "Bearer TESTDATEN-key-nicht-echt"
    body = json.loads(gesendete_anfrage.data.decode("utf-8"))
    assert body["messages"] == [
        {"role": "system", "content": "Mein System-Prompt"},
        {"role": "user", "content": "Mein Prompt"},
    ]
    assert "model" in body


def test_http_fehler_wird_zu_langdockfehler(monkeypatch):
    import urllib.error

    monkeypatch.setenv("LANGDOCK_API_KEY", "TESTDATEN-key-nicht-echt")
    fehler = urllib.error.HTTPError(
        "https://api.langdock.com/openai/eu/v1/chat/completions",
        401,
        "Unauthorized",
        {},
        None,
    )
    fehler.read = lambda: b'{"error": "invalid key"}'
    with patch("praesentationsassistent.langdock.urllib.request.urlopen", side_effect=fehler):
        with pytest.raises(LangdockFehler, match="401"):
            generiere_text("Testprompt")


def test_dedizierte_basis_url_nutzt_api_public_pfad(monkeypatch):
    monkeypatch.setenv("LANGDOCK_API_KEY", "TESTDATEN-key-nicht-echt")
    monkeypatch.setenv("LANGDOCK_BASE_URL", "https://langdock.testfirma-nicht-echt.de")
    with patch("praesentationsassistent.langdock.urllib.request.urlopen") as m:
        m.return_value = _antwort("egal")
        generiere_text("Testprompt")

    gesendete_anfrage = m.call_args[0][0]
    assert gesendete_anfrage.full_url == (
        "https://langdock.testfirma-nicht-echt.de/api/public/v1/chat/completions"
    )
