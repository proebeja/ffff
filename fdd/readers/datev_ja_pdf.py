"""Reader für den DATEV/soveo-Jahresabschluss als PDF (AJNS New Media GmbH).

Die Datei trägt drei Bestandteile, die hier getrennt abgegriffen werden:

* **Bilanz** (Blatt 1–3) — aggregierte Positionen je Geschäftsjahr und Vorjahr.
  Abstimmziel, keine Strukturquelle.
* **Kontennachweis zur Bilanz** (Blatt 4–7) — Einzelkonten je Bilanzposition.
  Das ist die **Strukturquelle für die Bilanz**.
* Gewinn- und Verlustrechnung (Blatt 8–9) — aggregiert, ohne Kontennachweis.
  Für die GuV gibt es also keine Strukturquelle; sie bleibt abgeleitet.

Zwei Formateigenheiten:

1. Negative Beträge tragen ein **nachgestelltes** Minus ("165.657,50-").
2. Positionsüberschriften sind über mehrere Zeilen umbrochen und dabei
   getrennt ("Geschäftsaus-\\nstattung"). Sie werden zusammengesetzt und
   entsilbt, bevor der Crosswalk greift.

**Saldenspaltung:** Derselbe Kontoschlüssel kann auf beiden Bilanzseiten
stehen — Konto 1400 0 erscheint als Forderung aus L+L *und* unter den
sonstigen Verbindlichkeiten. Der Abschluss spaltet, die SuSa zeigt netto.
Der Reader hält beide Seiten fest (``gespaltene_konten``); für das Mapping
zählt die Seite, auf der der **Nettosaldo** liegt.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

from ..core.hausconvention import normalisiere
from .base import fingerprint
from .kontennachweis import KNKonto, Kontennachweis, _match_ueberschrift

_BETRAG = r"-?[\d.]+,\d{2}-?"
_KONTO_ZEILE = re.compile(
    rf"^(?P<konto>\d{{1,5}})\s+(?P<sub>\d)\s+(?P<bez>.+?)\s+(?P<gj>{_BETRAG})\s+(?P<vj>{_BETRAG})$")
_KONTO_ZEILE_1 = re.compile(
    rf"^(?P<konto>\d{{1,5}})\s+(?P<sub>\d)\s+(?P<bez>.+?)\s+(?P<gj>{_BETRAG})$")
_NUR_BETRAEGE = re.compile(rf"^(?:{_BETRAG})(?:\s+(?:{_BETRAG}))*$")
_BILANZ_ZEILE = re.compile(rf"^(?P<label>.+?)\s+(?P<betraege>(?:{_BETRAG})(?:\s+{_BETRAG})*)$")

_RAUSCHEN = ("blatt", "kontennachweis zur bilanz", "bilanz zum", "berlin",
             "geschaeftsjahr vorjahr", "konto bezeichnung eur", "eur eur",
             "handelsrecht", "uebertrag", "gewinn- und verlustrechnung",
             "ajns new media")


def _betrag(s: str) -> float:
    s = s.strip()
    neg = s.endswith("-")
    s = s.rstrip("-").lstrip("-")
    wert = float(s.replace(".", "").replace(",", "."))
    return -wert if neg else wert


@dataclass
class KNEintrag:
    """Eine Kontozeile des Kontennachweises — genau eine Bilanzseite."""

    konto: str
    bezeichnung: str
    sektion: str                    # "AKTIVA" | "PASSIVA"
    ueberschrift: str
    hgb_pfad: Optional[str]
    gj: float
    vj: float

    def vorzeichenrichtig(self, vorjahr: bool = False) -> float:
        """Databook-Konvention: Aktiva positiv, Passiva negativ."""
        wert = self.vj if vorjahr else self.gj
        return -wert if self.sektion == "PASSIVA" else wert


@dataclass
class BilanzZeile:
    label: str
    sektion: str
    hgb_pfad: Optional[str]
    gj: float
    vj: float


@dataclass
class DatevJA:
    eintraege: list[KNEintrag]
    bilanz: list[BilanzZeile]
    entity: str
    stichtag: str
    quelle_datei: str
    warnungen: list[str] = field(default_factory=list)

    def gespaltene_konten(self) -> dict[str, list[KNEintrag]]:
        """Konten, die der Abschluss auf beide Bilanzseiten aufteilt."""
        je_konto: dict[str, list[KNEintrag]] = {}
        for e in self.eintraege:
            je_konto.setdefault(e.konto, []).append(e)
        return {k: v for k, v in je_konto.items()
                if len({e.sektion for e in v}) > 1}

    def als_kontennachweis(self, periode_gj: str, periode_vj: str) -> Kontennachweis:
        """Strukturquelle für die Engine. Gespaltene Konten bekommen die Seite,
        auf der ihr Nettosaldo liegt — die SuSa führt sie ebenfalls netto."""
        je_konto: dict[str, list[KNEintrag]] = {}
        for e in self.eintraege:
            je_konto.setdefault(e.konto, []).append(e)

        konten: dict[str, KNKonto] = {}
        warnungen = list(self.warnungen)
        for konto, eintraege in je_konto.items():
            netto_gj = sum(e.vorzeichenrichtig() for e in eintraege)
            netto_vj = sum(e.vorzeichenrichtig(vorjahr=True) for e in eintraege)
            if len({e.sektion for e in eintraege}) > 1:
                seite = "AKTIVA" if netto_gj >= 0 else "PASSIVA"
                fuehrend = next(e for e in eintraege if e.sektion == seite)
                warnungen.append(
                    f"Konto {konto} ist im Abschluss gespalten ("
                    + " / ".join(f"{e.sektion} {e.gj:,.2f}" for e in eintraege)
                    + f"); netto {netto_gj:,.2f} -> {seite}.")
            else:
                fuehrend = eintraege[0]
            if fuehrend.hgb_pfad is None:
                continue
            konten[konto] = KNKonto(
                konto=konto, bezeichnung=fuehrend.bezeichnung,
                hgb_pfad=fuehrend.hgb_pfad,
                salden={periode_gj: netto_gj, periode_vj: netto_vj})
        return Kontennachweis(
            konten=konten, perioden=[periode_gj, periode_vj], entity=self.entity,
            quelle_datei=self.quelle_datei,
            fingerprint=fingerprint(self.quelle_datei), warnungen=warnungen)


def lies_datev_ja(pfad: str) -> DatevJA:
    import pdfplumber

    eintraege: list[KNEintrag] = []
    bilanz: list[BilanzZeile] = []
    warnungen: list[str] = []
    entity, stichtag = "Unbekannt", ""
    zustand = {"sektion": "AKTIVA", "ueberschrift": "", "fragmente": []}

    with pdfplumber.open(pfad) as pdf:
        for seite in pdf.pages:
            text = seite.extract_text() or ""
            kopf = normalisiere(text[:200])
            if not entity or entity == "Unbekannt":
                m = re.search(r"^(.+GmbH)\s*$", text, re.M)
                if m:
                    entity = m.group(1).strip()
            m = re.search(r"zum (\d{2}\.\d{2}\.\d{4})", text)
            if m and not stichtag:
                stichtag = m.group(1)
            if "kontennachweis zur bilanz" in kopf:
                # Der Zustand läuft über die Seiten weiter: eine Position kann
                # am Blattende umbrechen und auf dem nächsten Blatt ohne
                # erneute Überschrift fortgesetzt werden.
                _lies_kontennachweis_seite(text, eintraege, zustand, warnungen)
            elif "bilanz zum" in kopf:
                bilanz += _lies_bilanz_seite(text)

    return DatevJA(eintraege=eintraege, bilanz=bilanz, entity=entity,
                   stichtag=stichtag, quelle_datei=pfad, warnungen=warnungen)


def _ist_rauschen(zeile: str) -> bool:
    low = normalisiere(zeile)
    return not low or any(low.startswith(r) for r in _RAUSCHEN)


def _entsilbe(fragmente: list[str]) -> str:
    """Setzt eine über mehrere Zeilen umbrochene Überschrift zusammen und
    entfernt die Trennstriche am Zeilenende."""
    text = ""
    for f in fragmente:
        f = f.strip()
        if text.endswith("-"):
            text = text[:-1] + f
        elif text:
            text += " " + f
        else:
            text = f
    return re.sub(r"\s+", " ", text).strip()


def _lies_kontennachweis_seite(text: str, eintraege: list, zustand: dict,
                               warnungen: list[str]) -> None:
    """Eine Kontennachweis-Seite. ``zustand`` trägt Sektion und laufende
    Positionsüberschrift über den Seitenumbruch hinweg.

    Der Fragmentpuffer und die geltende Überschrift sind bewusst getrennt:
    Eine Position mit mehreren Konten nennt ihre Überschrift nur einmal, also
    muss sie über die Folgekonten stehen bleiben. Umgekehrt darf sich der
    Puffer nicht über Positionen hinweg aufsummieren."""
    for roh in text.split("\n"):
        zeile = roh.strip()
        low = normalisiere(zeile)
        if low in ("aktiva", "passiva"):
            # Nur ein echter Seitenwechsel setzt die Position zurück. Der
            # Sektionsmarker steht auch im Kopf jeder Folgeseite; dort läuft
            # dieselbe Position weiter.
            if low.upper() != zustand["sektion"]:
                zustand.update(sektion=low.upper(), ueberschrift="", fragmente=[])
            continue
        if _ist_rauschen(zeile):
            continue
        if _NUR_BETRAEGE.match(zeile):
            zustand["fragmente"] = []     # Zwischensumme schließt die Position
            continue

        m = _KONTO_ZEILE.match(zeile) or _KONTO_ZEILE_1.match(zeile)
        if m:
            if zustand["fragmente"]:
                zustand["ueberschrift"] = _entsilbe(zustand["fragmente"])
                zustand["fragmente"] = []
            ueberschrift = zustand["ueberschrift"]
            sektion = zustand["sektion"]
            pfad = _match_ueberschrift(ueberschrift, sektion) if ueberschrift else None
            if ueberschrift and pfad is None:
                warnungen.append(
                    f"Überschrift '{ueberschrift[:60]}' ({sektion}) ohne "
                    "Zuordnung im Crosswalk — betroffene Konten fallen zurück.")
            gruppen = m.groupdict()
            eintraege.append(KNEintrag(
                konto=f"{gruppen['konto']} {gruppen['sub']}",
                bezeichnung=gruppen["bez"].strip(), sektion=sektion,
                ueberschrift=ueberschrift, hgb_pfad=pfad,
                gj=_betrag(gruppen["gj"]),
                vj=_betrag(gruppen["vj"]) if gruppen.get("vj") else 0.0))
            continue

        # Zeile mit Betrag, aber ohne Kontonummer: eine Positionszeile, die der
        # Abschluss ohne Kontendetail ausweist (Jahresfehlbetrag, Bilanzverlust,
        # nicht gedeckter Fehlbetrag). Sie schließt die Position ab und darf
        # nicht als Überschriftenfragment weiterwirken.
        mp = _BILANZ_ZEILE.match(zeile)
        if mp:
            zustand["fragmente"] = []
            zustand.setdefault("ohne_konten", []).append(
                (zustand["sektion"], _entsilbe([mp.group("label")])))
            continue
        zustand["fragmente"].append(zeile)


def _lies_bilanz_seite(text: str) -> list[BilanzZeile]:
    """Aggregierte Bilanzpositionen. Zeilen ohne Kontonummer, mit ein oder
    zwei Beträgen; Überschriften ohne Betrag werden übersprungen."""
    zeilen: list[BilanzZeile] = []
    sektion = "AKTIVA"
    vortext: list[str] = []
    for roh in text.split("\n"):
        zeile = roh.strip()
        low = normalisiere(zeile)
        if low in ("aktiva", "passiva"):
            sektion, vortext = low.upper(), []
            continue
        if _ist_rauschen(zeile) or _NUR_BETRAEGE.match(zeile):
            continue
        m = _BILANZ_ZEILE.match(zeile)
        if not m:
            vortext.append(zeile)
            continue
        label = _entsilbe(vortext + [m.group("label")])
        vortext = []
        betraege = [_betrag(b) for b in re.findall(_BETRAG, m.group("betraege"))]
        if len(betraege) < 2:
            betraege.append(0.0)
        zeilen.append(BilanzZeile(
            label=label, sektion=sektion,
            hgb_pfad=_match_ueberschrift(label, sektion),
            gj=betraege[-2], vj=betraege[-1]))
    return zeilen
