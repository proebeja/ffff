"""Reconciliation SuSa gegen Kontennachweis (Vorstufe zu Schicht 3).

Stellt je HGB-Position die Summe aus der SuSa der Positionssumme des
Kontennachweises gegenüber und macht die Differenz sichtbar. Zusätzlich
werden die beiden Mengen-Differenzen ausgewiesen:

* Konten, die nur der Kontennachweis kennt (fehlten in der SuSa) — genau die
  Lücke, durch die bisher ganze Positionen unbemerkt fehlten.
* Konten, die nur die SuSa kennt (im Abschluss nicht nachgewiesen).

Die Differenz ist bewusst *keine* zu beseitigende Fehlgröße, sondern
bewahrenswerte Information (Abschlussbuchungen, Umgliederungen) — sie wird
ausgewiesen, nicht wegdefiniert.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from ..core.model import MappedAccount, NormalizedLedger
from ..readers.kontennachweis import Kontennachweis


@dataclass
class ReconZeile:
    hgb_pfad: str
    susa: dict[str, float]
    kn: dict[str, float]

    def differenz(self, p: str) -> float:
        return self.susa.get(p, 0.0) - self.kn.get(p, 0.0)

    def hat_differenz(self, perioden: list[str], schwelle: float = 0.005) -> bool:
        return any(abs(self.differenz(p)) > schwelle for p in perioden)


@dataclass
class Reconciliation:
    zeilen: list[ReconZeile]
    perioden: list[str]
    nur_im_kn: list[str] = field(default_factory=list)
    nur_in_susa: list[str] = field(default_factory=list)
    #: Konto -> (Bezeichnung, Saldo je Periode) für die Mengen-Differenzen
    details: dict[str, tuple[str, dict[str, float]]] = field(default_factory=dict)

    def gesamtdifferenz(self, p: str) -> float:
        return sum(z.differenz(p) for z in self.zeilen)

    @property
    def zeilen_mit_differenz(self) -> list[ReconZeile]:
        return [z for z in self.zeilen if z.hat_differenz(self.perioden)]


def reconcile(mapped: list[MappedAccount], kn: Kontennachweis,
              perioden: list[str],
              susa_konten: set[str] | None = None) -> Reconciliation:
    """Vergleicht die gemappten SuSa-Salden je HGB-Position mit den
    Positionssummen des Kontennachweises.

    ``susa_konten`` sind die Kontonummern, die tatsächlich aus der SuSa kamen
    (ohne die aus dem Kontennachweis ergänzten) — nur so lässt sich die
    Mengen-Differenz sauber ausweisen.
    """
    kn_pos = kn.positionen()
    kn_konten = set(kn.konten)
    susa_konten = susa_konten if susa_konten is not None else {m.konto for m in mapped}

    # SuSa-Summen je Position — nur echte SuSa-Konten, sonst spiegelte sich
    # der Kontennachweis selbst und die Differenz wäre per Konstruktion null.
    susa_pos: dict[str, dict[str, float]] = {}
    for m in mapped:
        if m.konto not in susa_konten:
            continue
        ziel = susa_pos.setdefault(m.hgb_pfad, {p: 0.0 for p in perioden})
        for p in perioden:
            ziel[p] += m.saldo(p)

    alle_pfade = sorted(set(susa_pos) | set(kn_pos))
    zeilen = [
        ReconZeile(
            hgb_pfad=pfad,
            susa=susa_pos.get(pfad, {p: 0.0 for p in perioden}),
            kn=kn_pos.get(pfad, {p: 0.0 for p in perioden}),
        )
        for pfad in alle_pfade
    ]

    details: dict[str, tuple[str, dict[str, float]]] = {}
    for konto in sorted(kn_konten - susa_konten):
        kk = kn.konten[konto]
        details[konto] = (kk.bezeichnung, {p: kk.salden.get(p, 0.0) for p in perioden})
    by_konto = {m.konto: m for m in mapped}
    for konto in sorted(susa_konten - kn_konten):
        m = by_konto.get(konto)
        if m is not None:
            details[konto] = (m.bezeichnung, {p: m.saldo(p) for p in perioden})

    return Reconciliation(
        zeilen=zeilen, perioden=list(perioden),
        nur_im_kn=sorted(kn_konten - susa_konten),
        nur_in_susa=sorted(susa_konten - kn_konten),
        details=details,
    )
