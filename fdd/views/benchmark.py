"""Benchmark: eigene Klassifizierung gegen eine frühere **menschliche**.

Die SuSa trägt in einer Nebenspalte 79 von Hand gesetzte Klassifizierungen aus
einer früheren Bearbeitung, in fremder Pfadsyntax (``/netdebt/cashlike/…``,
``/QoE/…``, ``/Gehalt/MA``, ``/profit and loss/income/sales/B2B``). Sie fließen
**nicht** ins Mapping ein — dieses Modul stellt sie nur daneben, damit sichtbar
wird, wo Werkzeug und Mensch auseinanderlaufen.

Übersetzt wird nur, was die fremde Syntax eindeutig hergibt:

* ``/netdebt/…``            -> ND
* ``/profit and loss/…``    -> PL
* ``/Gehalt/…``, ``/Sozialabgaben …`` -> PL (Personalaufwand)
* ``/QoE/…``                -> PL, aber die QoE-Dimension (Bereinigung der
  Ergebnisqualität) ist in unseren Klassen gar nicht abbildbar; sie ist eine
  zweite Achse neben der Bilanz-/GuV-Zuordnung und wird als solche vermerkt.
* ``/WC-Bereinigung/…``     -> **nicht übersetzbar**: sagt zwar Working
  Capital, lässt aber offen, ob TWC oder OWC.

Alles andere bleibt ausdrücklich unübersetzt, statt geraten zu werden.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from ..core.model import MappedAccount

EINDEUTIG = "eindeutig"
TEILWEISE = "teilweise"
UNUEBERSETZBAR = "nicht uebersetzbar"


@dataclass
class Uebersetzung:
    klasse: str | None
    guete: str
    hinweis: str = ""


def uebersetze(roh: str) -> Uebersetzung:
    """Fremde Klassifizierungssyntax -> kanonische Klasse."""
    t = (roh or "").strip()
    if not t:
        return Uebersetzung(None, UNUEBERSETZBAR, "leer")
    low = t.lower()

    if low.startswith("/netdebt"):
        return Uebersetzung("ND", EINDEUTIG)
    if low.startswith("/profit and loss"):
        return Uebersetzung("PL", EINDEUTIG)
    if low.startswith(("/gehalt", "/sozialabgaben")):
        return Uebersetzung("PL", EINDEUTIG, "Personalaufwand")
    if low.startswith("/qoe"):
        return Uebersetzung(
            "PL", TEILWEISE,
            "QoE ist eine Bereinigungsachse neben der GuV-Zuordnung und in "
            "unseren Klassen nicht abbildbar; verglichen wird nur die "
            "GuV-Eigenschaft.")
    if low.startswith("/wc"):
        return Uebersetzung(
            None, UNUEBERSETZBAR,
            "nennt Working Capital, lässt aber offen ob TWC oder OWC.")
    # Freitext ohne Pfadsyntax: nur eine führende Klassenangabe ist verwertbar.
    m = re.match(r"^(ND|TWC|OWC|FA|EQ|DT|PL)\b", t)
    if m:
        return Uebersetzung(m.group(1), TEILWEISE,
                            "Freitext, führende Klassenangabe ausgewertet.")
    return Uebersetzung(None, UNUEBERSETZBAR, "keine bekannte Syntax.")


@dataclass
class BenchmarkZeile:
    konto: str
    bezeichnung: str
    mensch_roh: str
    mensch_klasse: str | None
    guete: str
    hinweis: str
    eigene_klasse: str
    eigener_pfad: str
    eigene_na: str

    @property
    def vergleichbar(self) -> bool:
        return self.mensch_klasse is not None

    @property
    def abweichung(self) -> bool:
        return self.vergleichbar and self.mensch_klasse != self.eigene_klasse


@dataclass
class BenchmarkView:
    zeilen: list[BenchmarkZeile] = field(default_factory=list)
    entity: str = ""

    @property
    def abweichungen(self) -> list[BenchmarkZeile]:
        return [z for z in self.zeilen if z.abweichung]

    @property
    def uebereinstimmungen(self) -> list[BenchmarkZeile]:
        return [z for z in self.zeilen if z.vergleichbar and not z.abweichung]

    @property
    def unuebersetzbar(self) -> list[BenchmarkZeile]:
        return [z for z in self.zeilen if not z.vergleichbar]

    def quote(self) -> float:
        n = len(self.uebereinstimmungen) + len(self.abweichungen)
        return len(self.uebereinstimmungen) / n if n else 0.0


def baue_benchmark(mapped: list[MappedAccount], manuell: dict[str, str],
                   entity: str = "") -> BenchmarkView:
    by = {m.konto: m for m in mapped}
    zeilen: list[BenchmarkZeile] = []
    for konto, roh in sorted(manuell.items()):
        m = by.get(konto)
        if m is None:
            continue
        u = uebersetze(roh)
        zeilen.append(BenchmarkZeile(
            konto=konto, bezeichnung=m.bezeichnung, mensch_roh=roh,
            mensch_klasse=u.klasse, guete=u.guete, hinweis=u.hinweis,
            eigene_klasse=m.klasse.value, eigener_pfad=m.hgb_pfad,
            eigene_na=m.na_de))
    return BenchmarkView(zeilen=zeilen, entity=entity)
