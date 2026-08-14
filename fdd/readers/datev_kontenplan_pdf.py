"""Reader für den DATEV-Kontenplan (Sachkonten) als PDF.

Der Kontenplan ist **keine Strukturquelle**: er trägt keine HGB-Zuordnung.
Zwei Dinge liefert er trotzdem, und beide sind belastbar:

1. **Kontobezeichnungen** — sauberer und vollständiger als die SuSa-Spalte.
2. **DATEV-Funktionsbezeichnungen** — die Spalte markiert Geldkonten sowie
   die Sammelkonten Debitor und Kreditor. Das ist eine Zuordnung des
   Buchhaltungssystems selbst, kein Namensraten, und damit deterministisch.
   Sie greift nur dort, wo der Kontennachweis schweigt (Konten, die es 2023
   noch nicht gab), und steht in der Kaskade deshalb hinter ihm.

Der Kontenplan zeigt den Stand der bebuchten Konten zum Ausgabezeitpunkt und
deckt daher nicht alle Konten der SuSa ab; ``abdeckung`` weist das aus.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

_UV = "/Aktiva/B Umlaufvermoegen"
_FORD = f"{_UV}/II Forderungen und sonstige Vermoegensgegenstaende"

#: DATEV-Funktionsbezeichnung -> kanonischer HGB-Pfad. Bewusst kurz: nur die
#: Funktionen, die das System eindeutig vergibt.
FUNKTION_PFAD: dict[str, str] = {
    "Geldkonto": f"{_UV}/IV Kassenbestand und Guthaben bei Kreditinstituten",
    "Sammelkonto Debitor": f"{_FORD}/Forderungen aus Lieferungen und Leistungen",
    "Sammelkonto Kreditor": ("/Passiva/C Verbindlichkeiten/"
                             "Verbindlichkeiten aus Lieferungen und Leistungen"),
}

_ZEILE = re.compile(r"^(?P<von>\d{1,5})\s+(?P<vsub>\d)\s+(?P<bis>\d{1,5})\s+"
                    r"(?P<bsub>\d)\s+(?P<ski>[A-Z])\s*(?P<rest>.*)$")
_FUNKTION = re.compile(r"\s+(?P<zf>\d{1,3})\s+(?P<ski2>[A-Z])\s+"
                       r"(?P<funktion>Geldkonto|Sammelkonto Debitor|Sammelkonto Kreditor)\s*$")


@dataclass
class KontenplanEintrag:
    konto: str
    bezeichnung: str
    funktion: Optional[str] = None

    @property
    def hgb_pfad(self) -> Optional[str]:
        return FUNKTION_PFAD.get(self.funktion or "")


@dataclass
class Kontenplan:
    eintraege: dict[str, KontenplanEintrag] = field(default_factory=dict)
    quelle_datei: str = ""

    def bezeichnung(self, konto: str) -> Optional[str]:
        e = self.eintraege.get(konto)
        return e.bezeichnung if e and e.bezeichnung else None

    def hgb_pfad(self, konto: str) -> Optional[str]:
        e = self.eintraege.get(konto)
        return e.hgb_pfad if e else None

    def mit_funktion(self) -> dict[str, KontenplanEintrag]:
        return {k: e for k, e in self.eintraege.items() if e.funktion}

    def abdeckung(self, konten: set[str]) -> tuple[int, int]:
        """(abgedeckt, gesamt) bezogen auf die übergebenen SuSa-Konten."""
        return sum(1 for k in konten if k in self.eintraege), len(konten)


def lies_kontenplan(pfad: str) -> Kontenplan:
    import pdfplumber

    plan = Kontenplan(quelle_datei=pfad)
    with pdfplumber.open(pfad) as pdf:
        for seite in pdf.pages:
            for roh in (seite.extract_text() or "").split("\n"):
                m = _ZEILE.match(roh.strip())
                if not m:
                    continue
                # Nur Einzelkonten, keine Bereichszeilen (von != bis).
                if m.group("von") != m.group("bis") or m.group("vsub") != m.group("bsub"):
                    continue
                rest = m.group("rest").strip()
                funktion = None
                mf = _FUNKTION.search(rest)
                if mf:
                    funktion = mf.group("funktion")
                    rest = rest[:mf.start()].strip()
                # Am Zeilenende steht das SKI-Kennzeichen der zweiten Spalte.
                rest = re.sub(r"\s+[A-Z]$", "", rest).strip()
                konto = f"{m.group('von')} {m.group('vsub')}"
                plan.eintraege[konto] = KontenplanEintrag(
                    konto=konto, bezeichnung=rest, funktion=funktion)
    return plan


def wende_kontenplan_an(ledger, plan: Kontenplan):
    """Setzt den ``fs_pfad`` aus der DATEV-Funktionsbezeichnung — aber nur für
    Konten, die noch keinen tragen. Der Kontennachweis behält damit Vorrang;
    der Kontenplan schließt nur dessen Lücken."""
    from dataclasses import replace

    from ..core.model import NormalizedLedger

    neue, warnungen, gesetzt = [], list(ledger.warnungen), 0
    for a in ledger.accounts:
        pfad = plan.hgb_pfad(a.konto)
        # Konten, die der Reader als technisch oder als ungeklärt gemeldet hat,
        # bleiben unangetastet — der Funktionscode darf eine offene Frage nicht
        # zuschütten.
        if a.fs_pfad is None and pfad and a.kontotyp not in ("technisch", "strittig"):
            neue.append(replace(a, fs_pfad=pfad,
                                kontotyp=("bilanz_passiv" if pfad.startswith("/Passiva")
                                          else "bilanz_aktiv")))
            gesetzt += 1
        else:
            neue.append(a)
    if gesetzt:
        warnungen.append(
            f"{gesetzt} Konto(en) ohne Eintrag im Kontennachweis über die "
            "DATEV-Funktionsbezeichnung des Kontenplans zugeordnet "
            "(Geldkonto / Sammelkonto Debitor / Sammelkonto Kreditor).")
    return NormalizedLedger(
        accounts=neue, perioden=list(ledger.perioden), entity=ledger.entity,
        quelle_datei=ledger.quelle_datei,
        hat_kontennachweis=ledger.hat_kontennachweis,
        fingerprint=ledger.fingerprint, warnungen=warnungen)
