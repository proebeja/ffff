"""Lead NA (Net Assets) und Lead PL (GuV) — die beiden Übersichts-Leads.

Beide folgen derselben Mechanik wie Net-Debt- und WC-Lead: rein aus dem
Datenmodell abgeleitet, nach Klasse gefiltert, nach Net-Asset-Zeile gruppiert.
Keine eigene Rechenlogik, keine von Hand gelegten Verknüpfungen — jede Zeile
zieht später im Export aus genau einem Aufriss.

Lead NA ist die Net-Asset-Brücke: Anlagevermögen + Working Capital + Net Debt
+ latente Steuern = Net Assets. Da die Bilanz aufgeht, muss Net Assets plus
Eigenkapital null ergeben; genau das prüft die zweite Kontrollzeile im Export.

Lead PL bildet die GuV in HGB-Reihenfolge ab. Vorzeichen wie im Mastersheet
(Soll positiv, Haben negativ): Erträge stehen negativ, Aufwendungen positiv,
die Summe ist damit das Ergebnis mit umgekehrtem Vorzeichen.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from ..core.model import Klasse, MappedAccount


@dataclass
class LeadZeile:
    na_de: str
    na_en: str
    klasse: str
    betraege: dict[str, float]
    konten: list[MappedAccount] = field(default_factory=list)


@dataclass
class LeadBlock:
    """Ein Abschnitt des Leads mit eigener Zwischensumme."""

    titel: str
    zeilen: list[LeadZeile]

    def summe(self, p: str) -> float:
        return sum(z.betraege.get(p, 0.0) for z in self.zeilen)


@dataclass
class LeadView:
    perioden: list[str]
    bloecke: list[LeadBlock]
    entity: str = ""
    #: Zeilen, die nicht in einen Block laufen (Lead NA: das Eigenkapital als
    #: Gegenprobe; Lead PL: leer).
    nachrichtlich: list[LeadZeile] = field(default_factory=list)

    def gesamt(self, p: str) -> float:
        return sum(b.summe(p) for b in self.bloecke)


# Reihenfolge der Net-Asset-Zeilen innerhalb eines Blocks.
_NA_ORDNUNG = {
    "Immaterielle Vermoegensgegenstaende": 0, "Sachanlagen": 1, "Finanzanlagen": 2,
    "Vorraete": 10, "Forderungen aus L+L": 11, "Geleistete Anzahlungen": 12,
    "Verbindlichkeiten aus L+L": 13, "Erhaltene Anzahlungen": 14,
    "Sonstige Vermoegensgegenstaende": 20, "Aktive Rechnungsabgrenzung": 21,
    "Sonstige Verbindlichkeiten": 22, "Sonstige Rueckstellungen": 23,
    "Passive Rechnungsabgrenzung": 24,
    "Liquide Mittel": 30, "Wertpapiere": 31, "Ausleihungen": 32, "Anleihen": 33,
    "Verbindlichkeiten ggue. Kreditinstituten": 34, "Pensionsrueckstellungen": 35,
    "Steuerrueckstellungen": 36,
    "Forderungen ggue. verbundenen Unternehmen": 37,
    "Verbindlichkeiten ggue. verbundenen Unternehmen": 38,
    "Aktive latente Steuern": 40, "Passive latente Steuern": 41,
}

#: HGB-Reihenfolge der GuV-Positionen (§ 275 Abs. 2 GKV).
_PL_ORDNUNG = {
    "Umsatzerloese": 0,
    "Bestandsveraenderung": 1,
    "Andere aktivierte Eigenleistungen": 2,
    "Sonstige betriebliche Ertraege": 3,
    "Materialaufwand": 4,
    "Loehne und Gehaelter": 5,
    "Soziale Abgaben und Altersversorgung": 6,
    "Abschreibungen": 7,
    "Sonstige betriebliche Aufwendungen": 8,
    "Ertraege aus Beteiligungen": 9,
    "Sonstige Zinsen und aehnliche Ertraege": 10,
    "Zinsen und aehnliche Aufwendungen": 11,
    "Steuern vom Einkommen und vom Ertrag": 12,
    "Sonstige Steuern": 13,
    "Ertraege aus Verlustuebernahme": 14,
    "Aufgrund Gewinnabfuehrungsvertrag abgefuehrte Gewinne": 15,
}


def _gruppiere(mapped: list[MappedAccount], klassen: tuple[Klasse, ...],
               perioden: list[str]) -> list[LeadZeile]:
    zeilen: dict[str, LeadZeile] = {}
    for m in mapped:
        if m.klasse not in klassen:
            continue
        if not m.na_de or m.na_de.startswith("("):
            continue
        for periode in perioden:
            na_de, na_en = m.na_in(periode)
            z = zeilen.get(na_de)
            if z is None:
                z = LeadZeile(na_de=na_de, na_en=na_en, klasse=m.klasse.value,
                              betraege={q: 0.0 for q in perioden})
                zeilen[na_de] = z
            z.betraege[periode] += m.saldo(periode)
            if m not in z.konten:
                z.konten.append(m)
    return list(zeilen.values())


def baue_lead_na(mapped: list[MappedAccount], perioden: list[str],
                 entity: str = "") -> LeadView:
    """Net-Asset-Brücke: Anlagevermögen, Working Capital, Net Debt, latente
    Steuern. Das Eigenkapital läuft nicht in die Summe, sondern dient als
    Gegenprobe (Net Assets + Eigenkapital = 0)."""
    def sortiert(zeilen):
        return sorted(zeilen, key=lambda z: (_NA_ORDNUNG.get(z.na_de, 500), z.na_de))

    bloecke = [
        LeadBlock("Anlagevermögen", sortiert(_gruppiere(mapped, (Klasse.FA,), perioden))),
        LeadBlock("Net Working Capital",
                  sortiert(_gruppiere(mapped, (Klasse.TWC, Klasse.OWC), perioden))),
        LeadBlock("Net Debt", sortiert(_gruppiere(mapped, (Klasse.ND,), perioden))),
        LeadBlock("Latente Steuern", sortiert(_gruppiere(mapped, (Klasse.DT,), perioden))),
    ]
    eigenkapital = sortiert(_gruppiere(mapped, (Klasse.EQ,), perioden))
    return LeadView(perioden=list(perioden),
                    bloecke=[b for b in bloecke if b.zeilen],
                    entity=entity, nachrichtlich=eigenkapital)


def baue_lead_pl(mapped: list[MappedAccount], perioden: list[str],
                 entity: str = "") -> LeadView:
    """GuV in HGB-Reihenfolge, eine Zeile je GuV-Position."""
    zeilen = sorted(_gruppiere(mapped, (Klasse.PL,), perioden),
                    key=lambda z: (_PL_ORDNUNG.get(z.na_de, 500), z.na_de))
    block = LeadBlock("Gewinn- und Verlustrechnung", zeilen)
    return LeadView(perioden=list(perioden), bloecke=[block] if zeilen else [],
                    entity=entity)
