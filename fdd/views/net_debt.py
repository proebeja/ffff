"""Net-Debt-View (Vorstufe zu Schicht 6, Lead NA / Net Debt).

Rein aus dem Datenmodell abgeleitet: filtert Klasse = ND, gruppiert nach
NA-Zeile. Keine von Hand gelegten Verknüpfungen. Zwei Blöcke:

  A) Direkte Net-Debt-Positionen (eindeutige HGB-Position -> ND)
  B) Umgliederung aus NWC in ND (thereof-ND der drei gemischten Positionen)

Jede Zahl bleibt bis zum Einzelkonto rückverfolgbar (die Gruppen tragen ihre
Konten-Referenzen).
"""

from __future__ import annotations

from dataclasses import dataclass, field

from ..core.model import Klasse, MappedAccount


@dataclass
class NetDebtZeile:
    na_de: str
    na_en: str
    betraege: dict[str, float]              # periode -> Summe
    konten: list[MappedAccount] = field(default_factory=list)
    aus_mixed: bool = False


@dataclass
class NetDebtView:
    perioden: list[str]
    direkt: list[NetDebtZeile]
    umgliederung: list[NetDebtZeile]        # thereof-ND aus gemischten Positionen
    entity: str = ""

    def subtotal(self, periode: str) -> float:
        return sum(z.betraege.get(periode, 0.0)
                   for z in (self.direkt + self.umgliederung))

    @property
    def alle_zeilen(self) -> list[NetDebtZeile]:
        return self.direkt + self.umgliederung


def baue_net_debt(mapped: list[MappedAccount], perioden: list[str],
                  entity: str = "") -> NetDebtView:
    direkt: dict[str, NetDebtZeile] = {}
    mixed: dict[str, NetDebtZeile] = {}

    for m in mapped:
        if m.klasse != Klasse.ND:
            continue
        ziel = mixed if m.aus_mixed else direkt
        # Ein Konto mit Seitenwechsel trägt je Periode zu einer anderen
        # NA-Zeile bei; es erscheint deshalb in beiden und liefert je Periode
        # nur dorthin, wo es in dieser Periode steht.
        for periode in perioden:
            na_de, na_en = m.na_in(periode)
            zeile = ziel.get(na_de)
            if zeile is None:
                zeile = NetDebtZeile(na_de=na_de, na_en=na_en,
                                     betraege={q: 0.0 for q in perioden},
                                     aus_mixed=m.aus_mixed)
                ziel[na_de] = zeile
            zeile.betraege[periode] += m.saldo(periode)
            if m not in zeile.konten:
                zeile.konten.append(m)

    ordnung = _na_reihenfolge()

    def sortiere(d: dict[str, NetDebtZeile]) -> list[NetDebtZeile]:
        return sorted(d.values(),
                      key=lambda z: (ordnung.get(z.na_de, 999), z.na_de))

    return NetDebtView(
        perioden=list(perioden), direkt=sortiere(direkt),
        umgliederung=sortiere(mixed), entity=entity,
    )


def _na_reihenfolge() -> dict[str, int]:
    """Anzeigereihenfolge der ND-Zeilen, an das Namur-Databook angelehnt."""
    return {
        "Liquide Mittel": 0,
        "Wertpapiere": 1,
        "Ausleihungen": 2,
        "Anleihen": 3,
        "Verbindlichkeiten ggue. Kreditinstituten": 4,
        "Pensionsrueckstellungen": 5,
        "Steuerrueckstellungen": 6,
        "Forderungen ggue. verbundenen Unternehmen": 7,
        "Verbindlichkeiten ggue. verbundenen Unternehmen": 8,
        "Sonstige Vermoegensgegenstaende": 20,
        "Sonstige Rueckstellungen": 21,
        "Sonstige Verbindlichkeiten": 22,
    }
