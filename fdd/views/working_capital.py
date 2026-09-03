"""Working-Capital-View (Vorstufe zu Schicht 6, Lead NA / Working Capital).

Gleiche Mechanik wie die Net-Debt-View: rein aus dem Datenmodell abgeleitet,
gefiltert nach Klasse (TWC/OWC) und Net-Asset-Zeile. Keine eigene Rechenlogik.

Primäre Gruppierung ist die klassische FDD-Schnittrichtung: erst der
TWC-Block (alle Trade-Positionen, Assets wie Liabilities), dann der OWC-Block,
darunter NWC = Saldo TWC + Saldo OWC. Die OA/OL-Ableitung (`seite`, aus der
Engine bzw. dem HGB-Pfad) bleibt vollständig im Datenmodell und wird im Lead
je Zeile als Spalte ausgewiesen — sie ist nur nicht mehr die Gliederungsebene.

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

    def zeilen_je_klasse(self, klasse: str) -> list[WCZeile]:
        """Alle Zeilen einer WC-Klasse, Assets vor Liabilities.

        Primäre Schnittrichtung des Leads: erst TWC, dann OWC — die
        OA/OL-Unterscheidung bleibt als Attribut je Zeile erhalten."""
        return [z for z in self.zeilen if z.klasse == klasse]


def _seite(m: MappedAccount) -> str:
    """WC-Seite des Kontos. Kommt aus der Engine (``oa_ol_ableitung``, v2.5);
    der Pfad-Fallback greift nur für Konten, die vor v2.5 gemappt wurden.

    Der Fallback fragt ``bilanzseite`` und nicht den Pfadanfang: bei einem
    Nicht-HGB-Kontenrahmen lautet der Pfad ``/AASB/Aktiva/...``, und ein
    ``startswith("/Aktiva")`` machte daraus stillschweigend eine Passivseite.
    """
    return m.seite or ("OA" if m.bilanzseite == "AKTIVA" else "OL")


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
        for periode in perioden:
            na_de, na_en = m.na_in(periode)
            # Geschlüsselt wird nach Klasse und NA-Zeile, NICHT nach Seite:
            # ein und dieselbe Position darf nur einmal im Lead stehen. Die
            # Seite ist ein Attribut der Position (sie folgt ihrem HGB-Pfad),
            # keine zweite Gliederungsebene — sonst zöge dieselbe Position
            # zweimal aus demselben Aufriss und stünde doppelt im Saldo.
            key = (m.klasse.value, na_de)
            z = zeilen.get(key)
            if z is None:
                z = WCZeile(na_de=na_de, na_en=na_en, klasse=m.klasse.value,
                            seite=_seite(m),
                            betraege={q: 0.0 for q in perioden})
                zeilen[key] = z
            z.betraege[periode] += m.saldo(periode)
            if m not in z.konten:
                z.konten.append(m)

    ordnung = _na_reihenfolge()
    geordnet = sorted(
        zeilen.values(),
        # TWC vor OWC (primäre Gruppierung des Leads), darin Assets vor
        # Liabilities. Die OA/OL-Ableitung bleibt im Datenmodell erhalten.
        key=lambda z: (0 if z.klasse == "TWC" else 1,
                       0 if z.seite == "OA" else 1,
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
