"""Erweiterungspunkt fuer die spaetere Langdock-Anbindung (LLM-Textgenerierung).

Bewusst NICHT implementiert: Aaron liefert die Langdock-Zugangsdaten erst
spaeter. Bis dahin ist dies ein leerer, aber stabiler Erweiterungspunkt --
andere Module (z.B. eine spaetere Dramaturgie-/Text-Generierung) koennen
schon heute gegen diese Funktionssignatur entwickelt werden, ohne dass sich
spaeter die Aufrufstelle aendert.

Kein Fakten-Erfinden: sobald Langdock angebunden ist, muss der Aufrufer dem
LLM ausschliesslich Fakten aus den Quelldokumenten mitgeben (Prompt-Bau) --
das ist hier bewusst noch nicht entschieden/implementiert.
"""

from __future__ import annotations


def generiere_text(prompt: str) -> str:
    """Erzeugt Text ueber Langdock. Noch nicht implementiert.

    Raises:
        NotImplementedError: immer, bis der Langdock-Zugang konfiguriert ist.
    """
    raise NotImplementedError("Langdock-Zugang noch nicht konfiguriert")
