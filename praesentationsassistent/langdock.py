"""Langdock-Anbindung (LLM-Textgenerierung) ueber die Chat-Completions-API.

Der API-Key bleibt bewusst lokal bei jedem Nutzer (Aarons Vorgabe 30.10.2026:
"jeder Benutzer soll sein eigenes Langdock-Account benutzen", Key wird nie
geteilt/zentral gespeichert) -- diese Funktion liest ihn nur aus einer
Umgebungsvariable (gesetzt ueber eine lokale ``.env``-Datei auf dem jeweiligen
Rechner, analog zu ``rechnung.env`` beim Rechnungsprogramm).

Quelle fuer das API-Format: oeffentliche Langdock-Entwicklerdokumentation
(docs.langdock.com, Stand 30.10.2026) -- Chat-Completions-Endpunkt ist zum
OpenAI-Format kompatibel. UNGETESTET gegen einen echten Key (keiner verfuegbar
in dieser Umgebung) -- Format anhand der Doku nachgebaut, mit Mock-Tests
gegen das erwartete Request/Response-Schema abgesichert. Sobald ein echter
Key verfuegbar ist, sollte ein echter Aufruf das bestaetigen.

Kein Fakten-Erfinden (Projektprinzip): der Aufrufer darf dem Prompt nur
Fakten aus den Quelldokumenten mitgeben -- das entscheidet der Aufrufer,
nicht dieses Modul.
"""

from __future__ import annotations

import json
import os
import urllib.error
import urllib.request

_STANDARD_BASIS_URL = "https://api.langdock.com"
_STANDARD_PFAD = "/openai/eu/v1/chat/completions"
_STANDARD_MODELL = "gpt-5.4"
_TIMEOUT_S = 60


class LangdockFehler(Exception):
    """Fehler bei einem Langdock-Aufruf (fehlender Key, HTTP-Fehler, o.ae.)."""


def _basis_url() -> str:
    """Basis-URL -- Langdock Cloud per Default, oder eine dedizierte
    Deployment-URL aus ``LANGDOCK_BASE_URL`` (dann OHNE ``/openai/eu``-Pfad,
    die dedizierte Variante haengt stattdessen ``/api/public`` an, siehe
    Doku "Dedicated Deployment")."""
    override = os.environ.get("LANGDOCK_BASE_URL", "").rstrip("/")
    return override or _STANDARD_BASIS_URL


def _completions_url() -> str:
    basis = _basis_url()
    if basis == _STANDARD_BASIS_URL:
        return basis + _STANDARD_PFAD
    # Dedizierte Deployment-URL: laut Doku "<domain>/api/public" + gleicher
    # Chat-Completions-Unterpfad ohne die Cloud-spezifische "/openai/eu"-Region.
    return basis + "/api/public/v1/chat/completions"


def generiere_text(prompt: str, system: str | None = None) -> str:
    """Erzeugt Text ueber Langdock (Chat-Completions, OpenAI-kompatibles
    Format). Liest den API-Key aus der Umgebungsvariable
    ``LANGDOCK_API_KEY`` -- wird die Variable nicht gesetzt, bricht die
    Funktion klar ab, statt einen falschen/leeren Key zu versuchen.

    Raises:
        LangdockFehler: Key fehlt, oder die Anfrage schlaegt fehl
            (Netzwerk, HTTP-Fehlerstatus, unerwartetes Antwortformat).
    """
    key = os.environ.get("LANGDOCK_API_KEY", "").strip()
    if not key:
        raise LangdockFehler(
            "LANGDOCK_API_KEY ist nicht gesetzt. Bitte den eigenen Langdock-"
            "API-Key lokal in der .env-Datei eintragen (bleibt nur auf "
            "diesem Rechner, wird nirgends hochgeladen)."
        )

    nachrichten = []
    if system:
        nachrichten.append({"role": "system", "content": system})
    nachrichten.append({"role": "user", "content": prompt})

    modell = os.environ.get("LANGDOCK_MODEL", "").strip() or _STANDARD_MODELL
    body = json.dumps({"model": modell, "messages": nachrichten}).encode("utf-8")

    anfrage = urllib.request.Request(
        _completions_url(),
        data=body,
        method="POST",
        headers={
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
        },
    )

    try:
        with urllib.request.urlopen(anfrage, timeout=_TIMEOUT_S) as antwort:
            daten = json.loads(antwort.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise LangdockFehler(
            f"Langdock antwortete mit Fehler {exc.code}: {detail[:500]}"
        ) from exc
    except urllib.error.URLError as exc:
        raise LangdockFehler(f"Langdock war nicht erreichbar: {exc}") from exc

    try:
        return daten["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise LangdockFehler(
            f"Unerwartetes Antwortformat von Langdock: {daten!r}"
        ) from exc
