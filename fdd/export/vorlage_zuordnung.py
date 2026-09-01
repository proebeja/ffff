"""Zuordnung unseres Positionsvokabulars auf das der Dealtool-Vorlage.

Die Vorlage trägt ihr eigenes Vokabular in Ticker 1. Unseres stammt aus
``reklassifizierung`` der Hausconvention. Beide sind an vielen Stellen
identisch — und an einigen eben nicht. Genau dort entsteht der Fehler, vor dem
die Übergabe warnt: ``SUMIFS`` findet nichts, die Positionszeile zeigt stumm
null, während die Kontozeile darunter einen Wert ausweist.

Deshalb ist die Zuordnung hier **explizit und vollständig aufgeschrieben**,
nicht im Befüllcode verstreut:

* Der Regelfall ist die Identität. Wer nicht in ``ABWEICHUNGEN`` steht, wird
  unverändert übernommen.
* Jede Abweichung trägt eine Begründung. Sie ist eine fachliche Entscheidung,
  keine Schreibweisenkorrektur.
* Positionen, für die die Vorlage keine Zeile hat, entstehen über eine
  Dummy-Zeile. Auf der Bilanzseite ergibt sich der Block aus der Klasse, in
  der GuV muss er benannt werden — ``Ertraege aus Verlustuebernahme`` gehört
  ins Finanzergebnis, das steht in keiner Klasse.

Die Klassen sind das zweite Ticker-Kriterium. Unsere Klasse ``OWC`` kennt die
Vorlage nicht; sie trennt in ``OA`` und ``OL``. Diese Trennung leitet
``oa_ol_ableitung`` der Hausconvention ohnehin schon aus der Bilanzseite ab —
wir übernehmen sie hier nur in die Spalte, statt sie neu zu erfinden.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from ..core.model import Klasse, MappedAccount


@dataclass(frozen=True)
class Abweichung:
    """Eine Position, deren Vorlagenname vom unseren abweicht."""

    ziel_na: str
    #: Nur nötig, wenn die Vorlage gar keine passende Position hat und die
    #: Zeile über eine Dummy-Zeile entsteht. Auf der Bilanzseite ergibt sich
    #: der Block aus der Klasse; in der GuV gibt es keine Klasse zum Ableiten.
    block: Optional[str] = None
    grund: str = ""
    #: Beschriftung der neuen Zeile. Unser Positionsvokabular ist durchgehend
    #: transliteriert (``Ertraege``); in einer Berichtszeile hat das nichts zu
    #: suchen.
    label_de: str = ""
    label_en: str = ""


#: Unser ``na_de`` -> Vokabular der Vorlage. Nur echte Abweichungen.
#:
#: Die Bilanzseite steht bewusst nicht darin: dort stimmen beide Vokabulare
#: Zeichen für Zeichen überein. Die GuV der Vorlage ist dagegen keine
#: § 275-Gliederung, sondern eine Kostenartenrechnung — sie fasst zusammen,
#: wo das HGB trennt, und trennt, wo das HGB zusammenfasst.
ABWEICHUNGEN: dict[str, Abweichung] = {
    "Andere aktivierte Eigenleistungen": Abweichung(
        "Aktivierte Eigenleistungen",
        grund="Gleiche Position, kürzerer Name in der Vorlage."),
    "Loehne und Gehaelter": Abweichung(
        "Personalaufwand",
        grund="Die Vorlage führt den Personalaufwand als eine Zeile. Die "
              "Trennung Löhne/Sozialabgaben bleibt im Aufriss und in den "
              "Kontoslots sichtbar."),
    "Soziale Abgaben und Altersversorgung": Abweichung(
        "Personalaufwand",
        grund="Siehe Löhne und Gehälter — beide laufen in dieselbe Zeile."),
    "Ertraege aus Beteiligungen": Abweichung(
        "Beteiligungsergebnis",
        grund="Die Vorlage saldiert Erträge und Aufwendungen aus "
              "Beteiligungen zu einer Ergebniszeile."),
    "Sonstige Zinsen und aehnliche Ertraege": Abweichung(
        "Zinsertraege",
        grund="Gleiche Position, kürzerer Name in der Vorlage."),
    "Zinsen und aehnliche Aufwendungen": Abweichung(
        "Zinsaufwendungen",
        grund="Gleiche Position, kürzerer Name in der Vorlage."),
    "Steuern vom Einkommen und vom Ertrag": Abweichung(
        "Ertragssteuern",
        grund="Gleiche Position, kürzerer Name in der Vorlage."),
    "Ertraege aus Verlustuebernahme": Abweichung(
        "Ertraege aus Verlustuebernahme", block="Finanzergebnis",
        grund="Die Vorlage kennt die Position nicht. Ergebnisabführung "
              "gehört unterhalb des EBIT ins Finanzergebnis.",
        label_de="Erträge aus Verlustübernahme",
        label_en="Income from loss absorption"),
    "Aufgrund Gewinnabfuehrungsvertrag abgefuehrte Gewinne": Abweichung(
        "Abgefuehrte Gewinne", block="Finanzergebnis",
        grund="Die Vorlage kennt die Position nicht. Gegenstück zur "
              "Verlustübernahme, deshalb derselbe Block.",
        label_de="Abgeführte Gewinne",
        label_en="Profits transferred under a profit transfer agreement"),
}

#: Klassen, die keine Position im Lead NA haben. Das Eigenkapital ist in der
#: Net-Asset-Sicht keine Position, sondern das Ergebnis: ``Nettovermögen``
#: ist definitionsgemäß das Eigenkapital. Es steht deshalb nur im Equity Roll
#: Forward, nicht als Zeile im Blockaufbau.
OHNE_POSITION = (Klasse.EQ, Klasse.REVIEW, Klasse.TECH, Klasse.MIXED)


def ziel_klasse(m: MappedAccount) -> str:
    """Ticker 2 der Vorlage. ``OWC`` wird über die Bilanzseite zu ``OA``/``OL``.

    Die Seite ist bereits in ``m.seite`` abgeleitet (v2.5). Fehlt sie, ist das
    ein Befund und kein Anlass, hier eine zweite Ableitung aufzumachen — die
    Klasse bleibt dann stehen und läuft sichtbar ins Leere.
    """
    if m.klasse is Klasse.OWC and m.seite:
        return m.seite
    return m.klasse.value


def ziel_position(na_de: str) -> tuple[str, Optional[str], str]:
    """Unser Positionsname -> (Name in der Vorlage, Block, Begründung)."""
    ab = ABWEICHUNGEN.get(na_de)
    if ab is None:
        return na_de, None, ""
    return ab.ziel_na, ab.block, ab.grund
