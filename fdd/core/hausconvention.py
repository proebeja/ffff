"""Lädt und kapselt ``hausconvention.json``.

Die Datei ist **Konfiguration, kein Code** (Handover Abschnitt 2). Dieses Modul
liest sie, validiert grob und bietet typisierte Zugriffe. Es hartkodiert keine
Regel — jede Regel, jeder SKR-Bereich, jede Schwelle stammt aus der Datei.
"""

from __future__ import annotations

import json
import os
import unicodedata
from dataclasses import dataclass
from typing import Any, Optional

_DEFAULT_PATH = os.path.join(os.path.dirname(__file__), "..", "config", "hausconvention.json")


def normalisiere(text: str) -> str:
    """Normalisierung laut ``matching_logik``: lowercase + Umlaute (ae/oe/ue/ss).

    Umlaute werden zu ae/oe/ue aufgelöst, ß zu ss, damit Regeln, die schon in
    aufgelöster Schreibweise vorliegen ('rueckstell'), auf 'Rückstellung'
    matchen. Danach wird evtl. verbliebene Kombinationsdiakritik entfernt.
    """
    if text is None:
        return ""
    t = str(text).lower()
    t = (t.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue")
           .replace("ß", "ss"))
    # verbliebene Diakritika (z.B. aus decomposed input) entfernen
    t = unicodedata.normalize("NFKD", t)
    t = "".join(c for c in t if not unicodedata.combining(c))
    return t


@dataclass
class ReklassRegel:
    hgb_pfad: str
    klasse: str
    na_de: str
    na_en: str
    split_regelset: Optional[str] = None
    hinweis: str = ""
    # v2.5: WC-Seite (OA/OL). Keine eigene Klasse, sondern die Seite des
    # Working Capital — bei gemischten Positionen als ``seite_operating``.
    # None für FA/ND/EQ/DT (nicht Teil des Working Capital).
    seite: Optional[str] = None


@dataclass
class Typ1Regel:
    id: str
    hgb_pfad: str
    prioritaet: int
    alle: list[str]
    eines: list[str]
    keines: list[str]
    kontotyp: Optional[str] = None
    hinweis: str = ""


@dataclass
class Typ2Regel:
    id: str
    klasse: str
    prioritaet: int
    eines: list[str]
    alle: list[str]
    keines: list[str]
    review: bool = False
    hinweis: str = ""
    # v2.3-Marker (nur Status/Anzeige, keine Logik in dieser Scheibe):
    pflichtfrage: Optional[str] = None      # "aufriss" | "pension"
    verhaltenspruefung: bool = False
    gekoppelt_mit: Optional[str] = None
    standardfrage: str = ""                 # v2.5


class Hausconvention:
    """Objektschnittstelle auf die Konfigurationsdatei."""

    def __init__(self, data: dict[str, Any]):
        self._d = data
        self.version: str = data.get("version", "?")
        self._woerterbuch: dict[str, str] = data.get("pfad_woerterbuch_de_en", {})
        self._reklass = [self._parse_reklass(r) for r in data.get("reklassifizierung", [])]
        self._typ1 = [self._parse_typ1(r)
                      for r in data.get("mapping_regeln_hgb", {}).get("regeln", [])]
        self._typ2: dict[str, list[Typ2Regel]] = {
            regelset: [self._parse_typ2(r) for r in regeln]
            for regelset, regeln in data.get("klassifizierungs_regeln_mixed", {}).items()
            if not regelset.startswith("_")
        }
        skr = data.get("skr03_default_bereiche", {})
        self._skr_bereiche: list[tuple[int, int, str]] = [
            (int(a), int(b), pfad) for a, b, pfad in skr.get("bereiche", [])
        ]
        self._tech_ab: int = skr.get("technische_konten", {}).get("ab", 9000)
        self.wesentlichkeit: dict[str, Any] = data.get("wesentlichkeit", {})

    # ---- Fabrik -----------------------------------------------------------
    @classmethod
    def laden(cls, pfad: str = _DEFAULT_PATH) -> "Hausconvention":
        with open(pfad, encoding="utf-8") as f:
            return cls(json.load(f))

    # ---- Parser -----------------------------------------------------------
    @staticmethod
    def _parse_reklass(r: dict) -> ReklassRegel:
        return ReklassRegel(
            hgb_pfad=r["hgb_pfad"], klasse=r["klasse"],
            na_de=r.get("na_de", ""), na_en=r.get("na_en", ""),
            split_regelset=r.get("split_regelset"), hinweis=r.get("hinweis", ""),
            seite=r.get("seite") or r.get("seite_operating"),
        )

    @staticmethod
    def _parse_typ1(r: dict) -> Typ1Regel:
        return Typ1Regel(
            id=r["id"], hgb_pfad=r["hgb_pfad"], prioritaet=r.get("prioritaet", 0),
            alle=[normalisiere(x) for x in r.get("wenn_bezeichnung_alle", [])],
            eines=[normalisiere(x) for x in r.get("wenn_bezeichnung_eines", [])],
            keines=[normalisiere(x) for x in r.get("wenn_bezeichnung_keines", [])],
            kontotyp=r.get("gilt_fuer_kontotyp"), hinweis=r.get("hinweis", ""),
        )

    @staticmethod
    def _parse_typ2(r: dict) -> Typ2Regel:
        return Typ2Regel(
            id=r["id"], klasse=r["klasse"], prioritaet=r.get("prioritaet", 0),
            eines=[normalisiere(x) for x in r.get("wenn_eines", [])],
            alle=[normalisiere(x) for x in r.get("wenn_alle", [])],
            keines=[normalisiere(x) for x in r.get("wenn_keines", [])],
            review=r.get("review", False), hinweis=r.get("hinweis", ""),
            pflichtfrage=r.get("pflichtfrage"),
            verhaltenspruefung=r.get("verhaltenspruefung", False),
            gekoppelt_mit=r.get("gekoppelt_mit"),
            standardfrage=r.get("standardfrage", ""),
        )

    # ---- Zugriffe ---------------------------------------------------------
    @property
    def typ1_regeln(self) -> list[Typ1Regel]:
        return self._typ1

    def typ2_regeln(self, regelset: str) -> list[Typ2Regel]:
        return self._typ2.get(regelset, [])

    @property
    def reklass_regeln(self) -> list[ReklassRegel]:
        return self._reklass

    @property
    def skr_bereiche(self) -> list[tuple[int, int, str]]:
        return self._skr_bereiche

    @property
    def tech_ab(self) -> int:
        return self._tech_ab

    def uebersetze_pfad(self, hgb_pfad: str) -> str:
        """HGB-Pfad DE -> EN, Segment für Segment über das Pfad-Wörterbuch.

        Unbekannte Segmente (z.B. tiefere MIXED-Unterzeilen) bleiben stehen,
        damit nie Information verloren geht.
        """
        segmente = [s for s in hgb_pfad.split("/") if s]
        uebersetzt = [self._woerterbuch.get(s, s) for s in segmente]
        return "/" + "/".join(uebersetzt)

    def skr_pfad(self, konto: str) -> Optional[str]:
        """SKR03-Default: Kontonummer-Bereich -> Pfad. None außerhalb aller
        Bereiche oder bei nicht-numerischen Konten (z.B. SAP ``H035310000``)."""
        nummer = self._konto_nummer(konto)
        if nummer is None:
            return None
        for a, b, pfad in self._skr_bereiche:
            if a <= nummer <= b:
                return pfad
        return None

    def ist_technisch(self, konto: str) -> bool:
        nummer = self._konto_nummer(konto)
        return nummer is not None and nummer >= self._tech_ab

    @staticmethod
    def _konto_nummer(konto: str) -> Optional[int]:
        """Führende Ziffernfolge als Integer; None wenn keine (SAP-Alnum)."""
        s = str(konto).strip()
        ziffern = ""
        for ch in s:
            if ch.isdigit():
                ziffern += ch
            else:
                break
        return int(ziffern) if ziffern else None
