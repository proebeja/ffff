"""Entscheidungsprotokoll und begründungspflichtiger Override.

Zwei Aufgaben:
1. Jede Mapping-Entscheidung wird eingefroren (Reproduzierbarkeit).
2. Ein manueller Override einer *abgeleiteten* Klasse verlangt eine
   Begründung und wird protokolliert (Abnahmekriterium des Handovers).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Optional

from ..core.model import Klasse, MappedAccount, Quelle


@dataclass
class ProtokollEintrag:
    zeitpunkt: str
    konto: str
    aktion: str          # "map" | "override"
    hgb_pfad: str
    klasse: str
    quelle: str
    regel_id: Optional[str]
    begruendung: str


class Entscheidungsprotokoll:
    def __init__(self) -> None:
        self.eintraege: list[ProtokollEintrag] = []

    def protokolliere_map(self, m: MappedAccount) -> None:
        self.eintraege.append(ProtokollEintrag(
            zeitpunkt=_jetzt(), konto=m.konto, aktion="map",
            hgb_pfad=m.hgb_pfad, klasse=m.klasse.value, quelle=m.quelle.value,
            regel_id=m.regel_id, begruendung=m.begruendung,
        ))

    def override_klasse(self, m: MappedAccount, neue_klasse: Klasse,
                        begruendung: str) -> None:
        """Setzt die Klasse manuell. Begründung ist Pflicht (Audit-Spur)."""
        if not begruendung or not begruendung.strip():
            raise ValueError(
                "Override einer abgeleiteten Klasse verlangt eine Begründung "
                "(Audit-Spur, Abnahmekriterium)."
            )
        m.override_von = m.klasse
        m.override_begruendung = begruendung.strip()
        m.klasse = neue_klasse
        m.quelle = Quelle.OVERRIDE
        self.eintraege.append(ProtokollEintrag(
            zeitpunkt=_jetzt(), konto=m.konto, aktion="override",
            hgb_pfad=m.hgb_pfad, klasse=neue_klasse.value,
            quelle=Quelle.OVERRIDE.value, regel_id=None,
            begruendung=begruendung.strip(),
        ))


def _jetzt() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")
