"""Reader-Interface. Kapselt die Formaterkennung, damit die Engine
formatagnostisch bleibt: sie sieht nur ``NormalizedLedger``."""

from __future__ import annotations

import hashlib
from abc import ABC, abstractmethod

from ..core.model import NormalizedLedger


class Reader(ABC):
    """Basisklasse aller Reader."""

    #: kurzer Formatname für Protokoll/Detection
    name: str = "base"

    @classmethod
    @abstractmethod
    def kann_lesen(cls, pfad: str) -> bool:
        """Schnelltest, ob dieser Reader die Datei versteht (für detect)."""

    @abstractmethod
    def lesen(self, pfad: str) -> NormalizedLedger:
        """Datei -> normalisierte Kontenliste."""


def fingerprint(pfad: str) -> str:
    """SHA-256 der Rohdatei (Reproduzierbarkeit, Architektur-Spec Abschnitt 6)."""
    h = hashlib.sha256()
    with open(pfad, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()[:16]


def parse_deutsche_zahl(wert) -> float:
    """Robustes Parsen deutscher/DATEV-Zahlformate.

    Behandelt: bereits-numerische Werte; nachgestelltes '-' (DATEV/Haben);
    Tausenderpunkt + Dezimalkomma ('1.999.512,89'); leere Strings.
    """
    if wert is None or wert == "":
        return 0.0
    if isinstance(wert, (int, float)):
        return float(wert)
    s = str(wert).strip()
    if not s:
        return 0.0
    neg = False
    if s.endswith("-"):          # DATEV: nachgestelltes Minus = Haben/negativ
        neg = True
        s = s[:-1].strip()
    if s.startswith("-"):
        neg = True
        s = s[1:].strip()
    # deutsches Format: Punkt = Tausender, Komma = Dezimal
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        # kein Komma: Punkte könnten Tausender ODER Dezimal sein.
        # Heuristik: genau ein Punkt mit 1-2 Nachkommastellen -> Dezimal.
        if s.count(".") == 1 and len(s.split(".")[1]) in (1, 2):
            pass  # als Dezimalpunkt belassen
        else:
            s = s.replace(".", "")
    s = s.replace(" ", "").replace("\xa0", "")
    try:
        v = float(s)
    except ValueError:
        return 0.0
    return -v if neg else v
