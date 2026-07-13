"""Working-Capital-View (Vorstufe zu Schicht 6, Lead NA / Working Capital).

Gleiche Mechanik wie die Net-Debt-View: rein aus dem Datenmodell abgeleitet,
gefiltert nach Klasse (TWC/OWC) und Net-Asset-Zeile. Keine eigene Rechenlogik.

Zweigeteilt nach Operating Assets / Operating Liabilities (Seite aus dem
HGB-Pfad: /Aktiva vs. /Passiva), innerhalb je getrennt nach TWC und OWC.

Bewusst NICHT enthalten: die normalisierte Referenz / das Target Working
Capital. Dieser Tab zeigt nur das **Ist-Working-Capital je Periode**; die
Referenzbildung braucht die (noch nicht implementierte) Verhaltensprüfung.

Prinzip: die WC-Definition ist über alle Perioden identisch, weil alle
Perioden durch dieselbe Klassifizierung laufen. Genau das macht die spätere
Kaufpreisanpassung (Completion Accounts vs. Referenz-WC) überhaupt rechenbar.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from ..core.model import Klasse, MappedAccount

_WC_KLASSEN = (Klasse.TWC, Klasse.OWC)


@dataclass
class WCZeile:
    na_de: str
    na_en: str
    klasse: str                 # "TWC" | "OWC"
    seite: str                  # "OA" (Operating Assets) | "OL" (Operating Liabilities)
    betraege: dict[str, float]
    konten: list[MappedAccount] = field(default_factory=list)


@dataclass
class WCView:
    perioden: list[str]
    zeilen: list[WCZeile]
    entity: str = ""
    # Konten mit WC-Klasse, die in keiner NA-Zeile landen (Raster-Löcher):
    ohne_na_zeile: list[MappedAccount] = field(default_factory=list)

    def _summe(self, seite: str | None, klasse: str | None, periode: str) -> float:
        return sum(z.betraege.get(periode, 0.0) for z in self.zeilen
                   if (seite is None or z.seite == seite)
                   and (klasse is None or z.klasse == klasse))

    def operating_assets(self, p: str) -> float:
        return self._summe("OA", None, p)

    def operating_liabilities(self, p: str) -> float:
        return self._summe("OL", None, p)

    def twc(self, p: str) -> float:
        return self._summe(None, "TWC", p)

    def owc(self, p: str) -> float:
        return self._summe(None, "OWC", p)

    def net_working_capital(self, p: str) -> float:
        """Operating Assets − Operating Liabilities. In der vorzeichenrichtigen
        Speicherung (Aktiva +, Passiva −) ist das schlicht die Summe aller
        WC-Zeilen."""
        return self._summe(None, None, p)

    def zeilen_fuer(self, seite: str, klasse: str) -> list[WCZeile]:
        return [z for z in self.zeilen if z.seite == seite and z.klasse == klasse]


def _seite(m: MappedAccount) -> str:
    return "OA" if m.hgb_pfad.startswith("/Aktiva") else "OL"


def baue_working_capital(mapped: list[MappedAccount], perioden: list[str],
                         entity: str = "") -> WCView:
    zeilen: dict[tuple[str, str, str], WCZeile] = {}
    ohne: list[MappedAccount] = []

    for m in mapped:
        if m.klasse not in _WC_KLASSEN:
            continue
        if not m.na_de or m.na_de.startswith("("):
            # WC-Klasse, aber keine echte NA-Zeile -> fällt durchs Raster
            ohne.append(m)
            continue
        seite = _seite(m)
        key = (seite, m.klasse.value, m.na_de)
        z = zeilen.get(key)
        if z is None:
            z = WCZeile(na_de=m.na_de, na_en=m.na_en, klasse=m.klasse.value,
                        seite=seite, betraege={p: 0.0 for p in perioden})
            zeilen[key] = z
        for p in perioden:
            z.betraege[p] += m.saldo(p)
        z.konten.append(m)

    ordnung = _na_reihenfolge()
    geordnet = sorted(
        zeilen.values(),
        key=lambda z: (0 if z.seite == "OA" else 1,      # Assets vor Liabilities
                       0 if z.klasse == "TWC" else 1,     # TWC vor OWC
                       ordnung.get(z.na_de, 500), z.na_de),
    )
    return WCView(perioden=list(perioden), zeilen=geordnet, entity=entity,
                  ohne_na_zeile=ohne)


def _na_reihenfolge() -> dict[str, int]:
    return {
        # Operating Assets — TWC
        "Forderungen aus L+L": 0,
        "Vorraete": 1,
        "Geleistete Anzahlungen": 2,
        # Operating Assets — OWC
        "Sonstige Vermoegensgegenstaende": 10,
        "Aktive Rechnungsabgrenzung": 11,
        # Operating Liabilities — TWC
        "Verbindlichkeiten aus L+L": 20,
        "Erhaltene Anzahlungen": 21,
        # Operating Liabilities — OWC
        "Sonstige Verbindlichkeiten": 30,
        "Sonstige Rueckstellungen": 31,
        "Passive Rechnungsabgrenzung": 32,
    }
