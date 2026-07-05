"""KI-Urteilsschicht — Schnittstelle, in der ersten Scheibe bewusst inaktiv.

Der Hook sitzt an genau der Stelle, wo die Kaskade sonst "Review" setzt
(Handover Abschnitt 3/4). Ist ein Provider registriert, liefert er einen
begründeten Vorschlag mit Konfidenz und einem Pfad *aus der kanonischen Liste*
— immer als geflaggter Vorschlag (``review=True``), nie als stille
Endentscheidung. In v1 gibt es keinen Provider; ``vorschlagen`` gibt None
zurück und das Konto geht regulär in die Review-Queue.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Optional

from ..core.model import Account


@dataclass(frozen=True)
class KIVorschlag:
    hgb_pfad: str          # muss aus der kanonischen Pfad-Liste stammen
    konfidenz: float       # 0..1
    begruendung: str


# Signatur eines optionalen KI-Providers.
KIProvider = Callable[[Account, list[str]], Optional[KIVorschlag]]


class KIUrteilsschicht:
    """Gebundener Fallback. In v1 ohne Provider -> immer None."""

    def __init__(self, provider: Optional[KIProvider] = None):
        self._provider = provider

    @property
    def aktiv(self) -> bool:
        return self._provider is not None

    def vorschlagen(self, account: Account,
                    kanonische_pfade: list[str]) -> Optional[KIVorschlag]:
        """Best-guess für ein sonst ungelöstes Konto. None, wenn kein Provider
        registriert ist (Standard in v1) oder der Vorschlag keinen kanonischen
        Pfad trifft (Bindung an die kanonische Liste)."""
        if self._provider is None:
            return None
        vorschlag = self._provider(account, kanonische_pfade)
        if vorschlag is None:
            return None
        if vorschlag.hgb_pfad not in kanonische_pfade:
            # Bindung verletzt -> verworfen, Konto bleibt Review.
            return None
        return vorschlag
