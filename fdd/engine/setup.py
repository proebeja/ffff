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

Die **zweite** Frage ist der Kontenrahmen. Bis v2.9 war sie eine reine
Protokollangabe — es gab nur ``skr03_default_bereiche``, und jedes Mandat lief
gegen die HGB-Kette. Mit dem ersten Nicht-HGB-Rahmen (AASB) steuert die
Antwort: ``skr03``/``skr04``/``hgb`` lassen die bestehende Kette unverändert
laufen, ``aasb`` schaltet auf die Kaskade des Kontenrahmens um. Eine
unbekannte Antwort ist ein Fehler und kein stiller Rückfall auf HGB.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from ..core.kontenrahmen import Kontenrahmen, lade_rahmen


@dataclass
class SetupErgebnis:
    modus: str                      # "abschlusstreu" | "vorlaeufig"
    kontennachweis_datei: Optional[str]
    abdeckung: float                # Anteil der Konten mit Struktur aus dem KN
    meldung: str
    anforderung: str = ""           # Datenanforderung an den Mandanten
    #: Antwort auf Frage 2. ``None`` heißt HGB-Kette (SKR03/SKR04/eigener Plan).
    kontenrahmen: Optional[Kontenrahmen] = None
    #: Antwort auf Frage 5: Namen der verbundenen Unternehmen und
    #: Gesellschafter. Sie werden zur Laufzeit an die Konzernregel angehängt.
    konzernnamen: list[str] = field(default_factory=list)

    @property
    def ist_abschlusstreu(self) -> bool:
        return self.modus == "abschlusstreu"

    @property
    def rahmen_meldung(self) -> str:
        if self.kontenrahmen is None:
            return ("Kontenrahmen HGB — Zuordnung über Kontennachweis, "
                    "Typ-1-Regeln und SKR-Bereichstabelle.")
        kr = self.kontenrahmen
        return (f"{kr.name} {kr.version} — Zuordnung über den "
                f"Kontonamen ({len(kr.stichwortregeln)} Stichwortregeln, "
                f"{len(kr.bibliothek)} Musterkonten, {len(kr.fs_lines)} FS "
                f"Line Items). Die HGB-Kette greift nicht."
                + (f" {len(self.konzernnamen)} verbundene Unternehmen erfasst."
                   if self.konzernnamen else
                   " ACHTUNG: keine verbundenen Unternehmen erfasst — "
                   "Konzernsalden werden nicht als solche erkannt."))

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
    "2. Welcher Kontenrahmen liegt zugrunde (skr03 / skr04 / hgb / aasb)? "
    "Bei skr03, skr04 und hgb läuft die HGB-Kette; aasb schaltet auf den "
    "Kontenrahmen für australische Abschlüsse um.",
    "3. Wesentlichkeitsschwellen bestätigen (Default: 2 % Bilanzsumme, "
    "1 % für Net-Debt-Positionen).",
    "4. Deal-Kontext: Stichtag, Konsolidierungskreis, Locked Box oder "
    "Completion Accounts?",
    "5. Namen der verbundenen Unternehmen und Gesellschafter? Sie werden an "
    "die Regel für Konzernsalden angehängt — ohne sie liest die Software "
    "'Accrued Interest - Aurora' als Zinsabgrenzung statt als "
    "Konzernforderung.",
]


def setup(kontennachweis_datei: Optional[str], konten_gesamt: int = 0,
          konten_mit_kn_struktur: int = 0,
          eingebettete_struktur: Optional[str] = None,
          kontenrahmen: str = "skr03",
          konzernnamen: Optional[list[str]] = None) -> SetupErgebnis:
    """Bestimmt den Modus aus der Datenlage.

    ``eingebettete_struktur`` benennt eine Strukturquelle, die der Export
    selbst mitbringt (SAP-FS-Hierarchie). Sie ist dem Kontennachweis
    gleichwertig — liegt sie vor, ist das Databook abschlusstreu.

    ``kontenrahmen`` ist die Antwort auf Frage 2 und entscheidet, welche
    Kaskade die Engine fährt; ``konzernnamen`` die auf Frage 5."""
    rahmen = lade_rahmen(kontenrahmen)
    namen = list(konzernnamen or [])
    if not kontennachweis_datei and not eingebettete_struktur:
        return SetupErgebnis(
            modus="vorlaeufig", kontennachweis_datei=None, abdeckung=0.0,
            meldung=("Kein Kontennachweis vorhanden — Default-Modus. Die "
                     "Struktur stammt aus " + ("Hausconvention und "
                     "SKR-Default" if rahmen is None else
                     f"{rahmen.name} {rahmen.version}") + ". "
                     "Das Databook ist VORLÄUFIG und NICHT ABSCHLUSSTREU."),
            anforderung=("Kontennachweis (Einzelkonten je JA-Position) beim "
                         "Mandanten anfordern — er hebt Auto-Mapping und "
                         "Reconciliation zugleich auf Kontenebene."),
            kontenrahmen=rahmen, konzernnamen=namen,
        )
    abdeckung = (konten_mit_kn_struktur / konten_gesamt) if konten_gesamt else 0.0
    quelle = kontennachweis_datei and "Kontennachweis" or eingebettete_struktur
    rest = ("Hausconvention/SKR-Default" if rahmen is None
            else f"{rahmen.name} {rahmen.version}")
    meldung = (f"{quelle} vorhanden — abschlusstreuer Modus. "
               f"Struktur aus dem Abschluss für {konten_mit_kn_struktur} von "
               f"{konten_gesamt} Konten ({abdeckung:.0%}); für den Rest greift "
               f"die übrige Kaskade ({rest}).")
    anforderung = ("" if abdeckung > 0.95 else
                   "Für die nicht nachgewiesenen Konten den Kontennachweis "
                   "vervollständigen lassen.")
    return SetupErgebnis(modus="abschlusstreu",
                         kontennachweis_datei=kontennachweis_datei or eingebettete_struktur,
                         abdeckung=abdeckung, meldung=meldung,
                         anforderung=anforderung, kontenrahmen=rahmen,
                         konzernnamen=namen)
