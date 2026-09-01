"""Setup-Dialog: Datenlage klären, bevor gerechnet wird.

Erste und wichtigste Frage ist der **Kontennachweis**. Er entscheidet, in
welchem Modus das Engagement läuft:

* ``abschlusstreu``   — eine Strukturquelle liegt vor. Das ist entweder ein
  Kontennachweis oder eine im Export **eingebettete FS-Hierarchie** (z.B. der
  SAP-BW-Export, der seine HGB-Gliederung als Baum mitliefert). Beides leistet
  dasselbe: die HGB-Grundgliederung folgt dem Abschluss statt einem Default
  und ist auf ihn überleitbar. Entscheidend ist die Datenlage, nicht die Frage,
  über welchen Kanal die Struktur ins Tool kam.
* ``vorlaeufig``      — kein Kontennachweis. Die Struktur kommt aus
  Hausconvention und SKR-Default. Das ist ein ausdrücklicher **Default-Modus**:
  das Databook ist vorläufig und **nicht abschlusstreu**; es wird als solches
  markiert, und der Kontennachweis wird aktiv angefordert.

Die Architektur-Spec verlangt, dass das Tool zu Beginn meldet, in welchem der
beiden Fälle es läuft — genau das leistet dieses Modul.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass
class SetupErgebnis:
    modus: str                      # "abschlusstreu" | "vorlaeufig"
    kontennachweis_datei: Optional[str]
    abdeckung: float                # Anteil der Konten mit Struktur aus dem KN
    meldung: str
    anforderung: str = ""           # Datenanforderung an den Mandanten

    @property
    def ist_abschlusstreu(self) -> bool:
        return self.modus == "abschlusstreu"

    @property
    def databook_kennzeichen(self) -> str:
        return ("Abschlusstreu — HGB-Gliederung folgt dem Abschluss"
                if self.ist_abschlusstreu else
                "VORLÄUFIG — NICHT ABSCHLUSSTREU (Default-Modus ohne Kontennachweis)")


#: Frage-Reihenfolge des Setup-Dialogs. Der Kontennachweis steht bewusst
#: zuerst — er hebt Auto-Mapping und Reconciliation zugleich.
FRAGEN: list[str] = [
    "1. Liegt ein Kontennachweis (Einzelkonten je JA-Position) vor? "
    "Falls ja: bitte Datei angeben — er ist die maßgebliche Strukturquelle.",
    "2. Welcher Kontenrahmen liegt zugrunde (SKR03 / SKR04 / eigener ERP-Plan)?",
    "3. Wesentlichkeitsschwellen bestätigen (Default: 2 % Bilanzsumme, "
    "1 % für Net-Debt-Positionen).",
    "4. Deal-Kontext: Stichtag, Konsolidierungskreis, Locked Box oder "
    "Completion Accounts?",
]


def setup(kontennachweis_datei: Optional[str], konten_gesamt: int = 0,
          konten_mit_kn_struktur: int = 0,
          eingebettete_struktur: Optional[str] = None) -> SetupErgebnis:
    """Bestimmt den Modus aus der Datenlage.

    ``eingebettete_struktur`` benennt eine Strukturquelle, die der Export
    selbst mitbringt (SAP-FS-Hierarchie). Sie ist dem Kontennachweis
    gleichwertig — liegt sie vor, ist das Databook abschlusstreu."""
    if not kontennachweis_datei and not eingebettete_struktur:
        return SetupErgebnis(
            modus="vorlaeufig", kontennachweis_datei=None, abdeckung=0.0,
            meldung=("Kein Kontennachweis vorhanden — Default-Modus. Die "
                     "HGB-Struktur stammt aus Hausconvention und SKR-Default. "
                     "Das Databook ist VORLÄUFIG und NICHT ABSCHLUSSTREU."),
            anforderung=("Kontennachweis (Einzelkonten je JA-Position) beim "
                         "Mandanten anfordern — er hebt Auto-Mapping und "
                         "Reconciliation zugleich auf Kontenebene."),
        )
    abdeckung = (konten_mit_kn_struktur / konten_gesamt) if konten_gesamt else 0.0
    quelle = kontennachweis_datei and "Kontennachweis" or eingebettete_struktur
    meldung = (f"{quelle} vorhanden — abschlusstreuer Modus. "
               f"Struktur aus dem Abschluss für {konten_mit_kn_struktur} von "
               f"{konten_gesamt} Konten ({abdeckung:.0%}); für den Rest greift "
               "die übrige Kaskade (Hausconvention/SKR-Default).")
    anforderung = ("" if abdeckung > 0.95 else
                   "Für die nicht nachgewiesenen Konten den Kontennachweis "
                   "vervollständigen lassen.")
    return SetupErgebnis(modus="abschlusstreu",
                         kontennachweis_datei=kontennachweis_datei or eingebettete_struktur,
                         abdeckung=abdeckung, meldung=meldung,
                         anforderung=anforderung)
