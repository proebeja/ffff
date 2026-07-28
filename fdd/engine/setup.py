"""Setup-Dialog: Datenlage klären, bevor gerechnet wird.

Erste und wichtigste Frage ist der **Kontennachweis**. Er entscheidet, in
welchem Modus das Engagement läuft:

* ``abschlusstreu``   — Kontennachweis liegt vor. Die HGB-Grundgliederung
  folgt dem testierten Abschluss und ist auf ihn überleitbar.
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
        return ("Abschlusstreu — HGB-Gliederung folgt dem Kontennachweis"
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
          konten_mit_kn_struktur: int = 0) -> SetupErgebnis:
    """Bestimmt den Modus aus der Datenlage."""
    if not kontennachweis_datei:
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
    meldung = (f"Kontennachweis vorhanden — abschlusstreuer Modus. "
               f"Struktur aus dem Abschluss für {konten_mit_kn_struktur} von "
               f"{konten_gesamt} Konten ({abdeckung:.0%}); für den Rest greift "
               "die übrige Kaskade (Hausconvention/SKR-Default).")
    anforderung = ("" if abdeckung > 0.95 else
                   "Für die nicht nachgewiesenen Konten den Kontennachweis "
                   "vervollständigen lassen.")
    return SetupErgebnis(modus="abschlusstreu",
                         kontennachweis_datei=kontennachweis_datei,
                         abdeckung=abdeckung, meldung=meldung,
                         anforderung=anforderung)
