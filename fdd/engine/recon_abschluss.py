"""Abstimmung des Databooks gegen den Jahresabschluss — auf zwei Ebenen.

**Warum zwei Ebenen?** Der Abschluss und die SuSa unterscheiden sich aus zwei
verschiedenen Gründen, und die beiden Gründe zeigen sich an verschiedenen
Stellen:

* Die **Saldenspaltung** ist eine reine Darstellungsfrage: Konto 1400 0 steht
  im Abschluss mit 568.275,36 unter den Forderungen und mit 6.753,92 unter den
  sonstigen Verbindlichkeiten, die SuSa führt die Differenz netto. Auf
  Kontoebene stimmt alles; sichtbar wird die Spaltung erst beim Vergleich der
  **Positionssummen gegen die ausgewiesene Bilanz**.
* Eine **Verrechnung zwischen Konten derselben Position** (1789 in 1780) ist
  umgekehrt auf Positionsebene unsichtbar und zeigt sich nur auf **Kontoebene
  gegen den Kontennachweis**.

Deshalb rechnet dieses Modul beides und trennt in jeder Ebene, was erklärt ist
und was als echte Differenz stehen bleibt. Erklärte Abstimmposten sind keine
Fehler — sie werden ausgewiesen, nicht wegdefiniert.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from ..core.model import MappedAccount
from ..readers.datev_ja_pdf import DatevJA

SALDENSPALTUNG = "Saldenspaltung (Abschluss spaltet, SuSa netto)"
VERRECHNUNG = "Verrechnung innerhalb derselben Position"
OFFEN = "offen"


@dataclass
class PositionsZeile:
    """Positionssumme des Databooks gegen die ausgewiesene Bilanz."""

    hgb_pfad: str
    susa: float
    abschluss: float
    erklaert: float
    erklaerung: str = ""

    @property
    def differenz(self) -> float:
        return self.susa - self.abschluss

    @property
    def rest(self) -> float:
        """Was nach Abzug der erklärten Abstimmposten stehen bleibt."""
        return self.differenz - self.erklaert


@dataclass
class KontoZeile:
    """Einzelkonto des Databooks gegen den Kontennachweis (netto)."""

    konto: str
    bezeichnung: str
    hgb_pfad: str
    susa: float
    kontennachweis: float
    art: str
    erklaerung: str = ""

    @property
    def differenz(self) -> float:
        return self.susa - self.kontennachweis


@dataclass
class AbschlussRecon:
    periode: str
    stichtag: str
    positionen: list[PositionsZeile] = field(default_factory=list)
    konten: list[KontoZeile] = field(default_factory=list)
    nur_im_abschluss: list[str] = field(default_factory=list)
    nur_in_susa: list[str] = field(default_factory=list)

    @property
    def erklaerte_posten(self) -> list[KontoZeile]:
        return [k for k in self.konten if k.art != OFFEN]

    @property
    def echte_differenzen(self) -> list[KontoZeile]:
        return [k for k in self.konten if k.art == OFFEN]

    def rest_gesamt(self) -> float:
        return sum(p.rest for p in self.positionen)


def _positionssummen(mapped: list[MappedAccount], periode: str) -> dict[str, float]:
    out: dict[str, float] = {}
    for m in mapped:
        if m.hgb_pfad.startswith("("):
            continue
        out[m.hgb_pfad] = out.get(m.hgb_pfad, 0.0) + m.saldo(periode)
    return out


def reconcile_gegen_kontennachweis(mapped: list[MappedAccount], ja: DatevJA,
                                   periode: str, vorjahr: bool = False
                                   ) -> AbschlussRecon:
    """Databook gegen Kontennachweis und ausgewiesene Bilanz für ein Jahr."""
    by_konto = {m.konto: m for m in mapped}

    # --- Kontoebene: SuSa gegen den netto gerechneten Kontennachweis --------
    je_konto: dict[str, list] = {}
    for e in ja.eintraege:
        je_konto.setdefault(e.konto, []).append(e)

    konten: list[KontoZeile] = []
    diff_je_position: dict[str, float] = {}
    for konto, eintraege in sorted(je_konto.items()):
        netto = sum(e.vorzeichenrichtig(vorjahr) for e in eintraege)
        m = by_konto.get(konto)
        susa = m.saldo(periode) if m else 0.0
        pfad = (m.hgb_pfad if m else next(
            (e.hgb_pfad for e in eintraege if e.hgb_pfad), "(ohne Pfad)")) or "(ohne Pfad)"
        if abs(susa - netto) > 0.005:
            konten.append(KontoZeile(
                konto=konto, bezeichnung=(m.bezeichnung if m else eintraege[0].bezeichnung),
                hgb_pfad=pfad, susa=susa, kontennachweis=netto, art=OFFEN))
            diff_je_position[pfad] = diff_je_position.get(pfad, 0.0) + (susa - netto)

    # Konten, die die SuSa führt, der Kontennachweis aber nicht — nur relevant,
    # wenn sie in einer vom Abschluss ausgewiesenen Position liegen.
    for m in mapped:
        if m.konto in je_konto or m.hgb_pfad.startswith(("(", "/GuV")):
            continue
        if abs(m.saldo(periode)) > 0.005:
            konten.append(KontoZeile(
                konto=m.konto, bezeichnung=m.bezeichnung, hgb_pfad=m.hgb_pfad,
                susa=m.saldo(periode), kontennachweis=0.0, art=OFFEN,
                erklaerung="Konto im Kontennachweis nicht nachgewiesen."))
            diff_je_position[m.hgb_pfad] = (
                diff_je_position.get(m.hgb_pfad, 0.0) + m.saldo(periode))

    # Gleicht sich die Abweichung innerhalb einer Position aus, ist es eine
    # Verrechnung zwischen Konten und keine Differenz der Position.
    for k in konten:
        if abs(diff_je_position.get(k.hgb_pfad, 0.0)) < 0.005:
            k.art = VERRECHNUNG
            k.erklaerung = ("Gegenposten in derselben Position; die "
                            "Positionssumme bleibt unberührt.")

    # --- Positionsebene: SuSa gegen die ausgewiesene (gespaltene) Bilanz ----
    susa_pos = _positionssummen(mapped, periode)
    abschluss_pos: dict[str, float] = {}
    for e in ja.eintraege:
        if e.hgb_pfad:
            abschluss_pos[e.hgb_pfad] = (abschluss_pos.get(e.hgb_pfad, 0.0)
                                         + e.vorzeichenrichtig(vorjahr))

    erklaert: dict[str, float] = {}
    erklaerung: dict[str, str] = {}
    for konto, eintraege in ja.gespaltene_konten().items():
        netto = sum(e.vorzeichenrichtig(vorjahr) for e in eintraege)
        fuehrend = "AKTIVA" if netto >= 0 else "PASSIVA"
        for e in eintraege:
            if not e.hgb_pfad:
                continue
            betrag = e.vorzeichenrichtig(vorjahr)
            gegen = netto - betrag if e.sektion == fuehrend else -betrag
            erklaert[e.hgb_pfad] = erklaert.get(e.hgb_pfad, 0.0) + gegen
            erklaerung[e.hgb_pfad] = SALDENSPALTUNG

    positionen = []
    for pfad in sorted(set(susa_pos) | set(abschluss_pos)):
        if pfad.startswith("/GuV"):
            continue
        s, a = susa_pos.get(pfad, 0.0), abschluss_pos.get(pfad, 0.0)
        e = erklaert.get(pfad, 0.0)
        if abs(s) < 0.005 and abs(a) < 0.005 and abs(e) < 0.005:
            continue
        positionen.append(PositionsZeile(
            hgb_pfad=pfad, susa=s, abschluss=a, erklaert=e,
            erklaerung=erklaerung.get(pfad, "")))

    return AbschlussRecon(
        periode=periode, stichtag=ja.stichtag, positionen=positionen,
        konten=sorted(konten, key=lambda k: (k.art != OFFEN, k.konto)),
        nur_im_abschluss=sorted(k for k in je_konto if k not in by_konto),
        nur_in_susa=sorted(m.konto for m in mapped
                           if m.konto not in je_konto
                           and not m.hgb_pfad.startswith(("(", "/GuV"))
                           and abs(m.saldo(periode)) > 0.005))


# ---- Aggregierte Abstimmung gegen den Prüfbericht ------------------------
@dataclass
class Ueberleitung:
    """Ein erklärter Posten zwischen Databook und Bericht."""

    text: str
    betrag: float


@dataclass
class AggregatZeile:
    label: str
    ebene: str                 # Pfad-Präfix, auf das aggregiert wird
    databook: float
    bericht: float
    ueberleitung: list[Ueberleitung] = field(default_factory=list)

    @property
    def differenz(self) -> float:
        return self.databook - self.bericht

    @property
    def erklaert(self) -> float:
        return sum(u.betrag for u in self.ueberleitung)

    @property
    def rest(self) -> float:
        return self.differenz - self.erklaert


@dataclass
class AggregatRecon:
    periode: str
    quelle: str
    zeilen: list[AggregatZeile] = field(default_factory=list)
    hinweise: list[str] = field(default_factory=list)

    @property
    def mit_differenz(self) -> list[AggregatZeile]:
        return [z for z in self.zeilen if abs(z.differenz) > 0.005]

    @property
    def mit_rest(self) -> list[AggregatZeile]:
        return [z for z in self.zeilen if abs(z.rest) > 0.005]

    def gesamt(self) -> float:
        return sum(z.differenz for z in self.zeilen)

    def rest_gesamt(self) -> float:
        return sum(z.rest for z in self.zeilen)


#: Gliederungsebene des Prüfberichts -> Präfix des kanonischen Pfads.
_EBENEN: list[tuple[str, str]] = [
    ("I. Immaterielle Vermögensgegenstände", "/Aktiva/A Anlagevermoegen/I Immaterielle"),
    ("II. Sachanlagen", "/Aktiva/A Anlagevermoegen/II Sachanlagen"),
    ("III. Finanzanlagen", "/Aktiva/A Anlagevermoegen/III Finanzanlagen"),
    ("I. Vorräte", "/Aktiva/B Umlaufvermoegen/I Vorraete"),
    ("II. Forderungen und sonstige Vermögensgegenstände",
     "/Aktiva/B Umlaufvermoegen/II Forderungen"),
    ("III. Kassenbestand und Guthaben bei Kreditinstituten",
     "/Aktiva/B Umlaufvermoegen/IV Kassenbestand"),
    ("C. Rechnungsabgrenzungsposten", "/Aktiva/C Rechnungsabgrenzungsposten"),
    ("A. Eigenkapital", "/Passiva/A Eigenkapital"),
    ("B. Rückstellungen", "/Passiva/B Rueckstellungen"),
    ("C. Verbindlichkeiten", "/Passiva/C Verbindlichkeiten"),
]


def reconcile_aggregiert(mapped: list[MappedAccount], periode: str,
                         bericht: dict[str, float], quelle: str,
                         hinweise: Optional[list[str]] = None,
                         ueberleitung: Optional[dict[str, list[Ueberleitung]]] = None
                         ) -> AggregatRecon:
    """Databook gegen einen nur aggregiert vorliegenden Abschluss.

    ``bericht`` bildet die Gliederungsebene des Berichts ab (Label -> Betrag,
    in Databook-Vorzeichen). Tiefer geht es nicht: der Prüfbericht weist keine
    Kontenebene aus, deshalb wird auch nur bis hierher abgestimmt."""
    zeilen = []
    for label, praefix in _EBENEN:
        if label not in bericht:
            continue
        db = sum(m.saldo(periode) for m in mapped
                 if m.hgb_pfad.startswith(praefix))
        zeilen.append(AggregatZeile(label=label, ebene=praefix, databook=db,
                                    bericht=bericht[label],
                                    ueberleitung=list((ueberleitung or {}).get(label, []))))
    return AggregatRecon(periode=periode, quelle=quelle, zeilen=zeilen,
                         hinweise=list(hinweise or []))
