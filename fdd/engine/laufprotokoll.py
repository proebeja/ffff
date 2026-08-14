"""Laufprotokoll — misst, wo die Zeit hingeht, bevor irgendetwas umgebaut wird.

Bewusst schlicht: ein Kontextmanager je Phase, ein Zähler für die Aufrufe an
die KI-Schicht und eine Zählung der geschriebenen Zellen. Keine Optimierung,
keine Heuristik — nur die Messung.
"""

from __future__ import annotations

import time
from contextlib import contextmanager
from dataclasses import dataclass, field


@dataclass
class Phase:
    name: str
    sekunden: float
    detail: str = ""


@dataclass
class Laufprotokoll:
    phasen: list[Phase] = field(default_factory=list)
    ki_aufrufe: list[str] = field(default_factory=list)
    zellen: int = 0
    formeln: int = 0
    blaetter: int = 0

    @contextmanager
    def phase(self, name: str):
        t = time.perf_counter()
        merker: dict[str, str] = {}
        try:
            yield merker
        finally:
            self.phasen.append(Phase(name, time.perf_counter() - t,
                                     merker.get("detail", "")))

    def ki_aufruf(self, wofuer: str) -> None:
        self.ki_aufrufe.append(wofuer)

    @property
    def gesamt(self) -> float:
        return sum(p.sekunden for p in self.phasen)

    def zaehle_arbeitsmappe(self, pfad: str) -> None:
        """Geschriebene Zellen und Formeln — die Größe des Ergebnisses, an der
        sich die Ausgabezeit misst."""
        import openpyxl

        wb = openpyxl.load_workbook(pfad)
        zellen = formeln = 0
        for ws in wb.worksheets:
            for reihe in ws.iter_rows():
                for c in reihe:
                    if c.value is None:
                        continue
                    zellen += 1
                    if isinstance(c.value, str) and c.value.startswith("="):
                        formeln += 1
        self.zellen, self.formeln, self.blaetter = zellen, formeln, len(wb.worksheets)
        wb.close()

    def als_text(self) -> list[str]:
        breite = max((len(p.name) for p in self.phasen), default=10)
        zeilen = [f"{'Phase':{breite}}  {'Sekunden':>9}  {'Anteil':>7}  Detail"]
        for p in self.phasen:
            anteil = p.sekunden / self.gesamt if self.gesamt else 0.0
            zeilen.append(f"{p.name:{breite}}  {p.sekunden:>9.2f}  "
                          f"{anteil:>6.1%}  {p.detail}")
        zeilen.append(f"{'GESAMT':{breite}}  {self.gesamt:>9.2f}  {1.0:>6.1%}")
        zeilen.append("")
        zeilen.append(f"Aufrufe an die KI-Schicht: {len(self.ki_aufrufe)}"
                      + (f" ({', '.join(self.ki_aufrufe[:5])})"
                         if self.ki_aufrufe else
                         " — kein Provider registriert, die Schicht ist eine "
                         "Schnittstelle ohne Implementierung"))
        zeilen.append(f"Geschriebene Zellen: {self.zellen:,} davon Formeln: "
                      f"{self.formeln:,} auf {self.blaetter} Blättern")
        return zeilen
