"""Abschlusstreue je Periodenspalte — nicht pauschal fürs ganze Buch.

Die Datenlage ist selten für alle Perioden gleich. Bei AJNS liegt für 2023 ein
Jahresabschluss mit Kontennachweis vor, für 2022 nur dessen Vorjahresspalte
plus ein aggregierter Prüfbericht, für 2024 und den Zwischenstand 2025 gar
nichts. Ein pauschales "abschlusstreu" wäre für die einen Spalten falsch, ein
pauschales "vorläufig" für die anderen — deshalb trägt **jede Spalte ihren
eigenen Status**, und zwar getrennt für Bilanz und GuV: der Kontennachweis
deckt nur die Bilanz ab, die GuV bleibt auch 2023 abgeleitet.
"""

from __future__ import annotations

from dataclasses import dataclass, field

ABSCHLUSSTREU = "abschlusstreu"
AGGREGIERT = "abschlusstreu (nur aggregiert)"
ABGELEITET = "abgeleitet — nicht abschlusstreu"
VORLAEUFIG = "vorläufig — nicht abschlusstreu"


@dataclass
class SpaltenStatus:
    periode: str
    bilanz: str
    guv: str
    quelle: str
    hinweis: str = ""

    @property
    def kurz(self) -> str:
        if self.bilanz == self.guv:
            return self.bilanz
        return f"Bilanz: {self.bilanz} · GuV: {self.guv}"

    @property
    def ist_abschlusstreu(self) -> bool:
        return self.bilanz.startswith(ABSCHLUSSTREU)


@dataclass
class StatusMatrix:
    spalten: list[SpaltenStatus] = field(default_factory=list)

    def fuer(self, periode: str) -> SpaltenStatus | None:
        return next((s for s in self.spalten if s.periode == periode), None)

    @property
    def gemischt(self) -> bool:
        return len({s.kurz for s in self.spalten}) > 1

    def zusammenfassung(self) -> str:
        return " | ".join(f"{s.periode}: {s.kurz}" for s in self.spalten)


def baue_status(perioden: list[str], kontennachweis_perioden: set[str],
                aggregiert_perioden: set[str], quellen: dict[str, str]
                ) -> StatusMatrix:
    """Ordnet jeder Periode ihren Status zu.

    ``kontennachweis_perioden`` sind die Spalten, für die ein Kontennachweis
    die Bilanzstruktur trägt; ``aggregiert_perioden`` jene, für die nur ein
    aggregierter Abschluss vorliegt. Die GuV ist in keinem Fall abschlusstreu,
    solange kein Kontennachweis zur GuV vorliegt."""
    spalten = []
    for p in perioden:
        if p in kontennachweis_perioden:
            bilanz = ABSCHLUSSTREU
        elif p in aggregiert_perioden:
            bilanz = AGGREGIERT
        else:
            bilanz = VORLAEUFIG
        guv = ABGELEITET if bilanz != VORLAEUFIG else VORLAEUFIG
        hinweis = ""
        if bilanz == VORLAEUFIG:
            hinweis = "Kein Abschluss vorhanden — Struktur aus Hausconvention/SKR03."
        elif bilanz == AGGREGIERT:
            hinweis = ("Abstimmung nur auf Gliederungsebene möglich; der "
                       "Bericht weist keine Kontenebene aus.")
        spalten.append(SpaltenStatus(periode=p, bilanz=bilanz, guv=guv,
                                     quelle=quellen.get(p, ""), hinweis=hinweis))
    return StatusMatrix(spalten=spalten)
