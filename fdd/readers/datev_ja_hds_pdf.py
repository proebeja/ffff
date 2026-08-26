"""Reader für den DATEV-Jahresabschluss im HDS-Layout (Brehna UG).

Anders als bei den bisherigen Mandanten gibt es **keine Summen- und
Saldenliste**. Der Jahresabschluss selbst ist Werte- *und* Strukturquelle: er
trägt einen vollständigen Kontennachweis für Aktivseite, Passivseite **und
Gewinn- und Verlustrechnung**. Damit ist auch die GuV abschlusstreu — bei
Kitchenstories war sie es mangels GuV-Kontennachweis nicht.

Layout je Kontennachweis-Seite::

    Pos Konto Bezeichnung          31.12.2025 31.12.2025
    1545 *** in Ausführung befindliche Bauaufträge   997.795,86
    709000 In Ausführung befindliche Bauaufträge     997.795,86
    *** SUMME AKTIVA                              1.427.519,32

Die vierstellige **Positionsnummer** ist die stabile Kennung der HGB-Position
und trägt den Crosswalk — nicht die Anzeigeschrift, die über mehrere Zeilen
umbricht und dabei getrennt wird ("Gutha-\\nben bei Kreditinstituten").

Vorzeichen: der Abschluss druckt die Passivseite positiv und Aufwendungen
negativ. Das Databook führt Soll positiv und Haben negativ, deshalb werden
Passiv- und GuV-Seite gedreht. Der Verlustvortrag steht im GuV-Nachweis, ist
aber ein Eigenkapitalkonto — er wird als solches geführt.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

from .base import fingerprint

_A = "/Aktiva"
_UV = f"{_A}/B Umlaufvermoegen"
_FORD = f"{_UV}/II Forderungen und sonstige Vermoegensgegenstaende"
_P = "/Passiva"
_EK = f"{_P}/A Eigenkapital"
_VB = f"{_P}/C Verbindlichkeiten"
_G = "/GuV"

#: DATEV-Positionsnummer -> kanonischer HGB-Pfad und Kontotyp.
POSITIONEN: dict[str, tuple[str, str]] = {
    # Aktivseite
    "1545": (f"{_UV}/I Vorraete/Unfertige Erzeugnisse und Leistungen", "bilanz_aktiv"),
    "1550": (f"{_UV}/I Vorraete/Fertige Erzeugnisse und Waren", "bilanz_aktiv"),
    "1570": (f"{_UV}/I Vorraete/Geleistete Anzahlungen", "bilanz_aktiv"),
    "1600": (f"{_FORD}/Forderungen aus Lieferungen und Leistungen", "bilanz_aktiv"),
    "1610": (f"{_FORD}/Forderungen gegen verbundene Unternehmen", "bilanz_aktiv"),
    "1640": (f"{_FORD}/Sonstige Vermoegensgegenstaende", "bilanz_aktiv"),
    "1710": (f"{_UV}/IV Kassenbestand und Guthaben bei Kreditinstituten", "bilanz_aktiv"),
    "1800": (f"{_A}/C Rechnungsabgrenzungsposten", "bilanz_aktiv"),
    # Passivseite
    "2020": (f"{_EK}/I Gezeichnetes Kapital", "bilanz_passiv"),
    "2100": (f"{_EK}/IV Gewinnvortrag Verlustvortrag", "bilanz_passiv"),
    "2115": (f"{_EK}/V Jahresueberschuss Jahresfehlbetrag", "bilanz_passiv"),
    "2125": (f"{_EK}/IV Gewinnvortrag Verlustvortrag", "bilanz_passiv"),
    "2350": (f"{_P}/B Rueckstellungen/Sonstige Rueckstellungen", "bilanz_passiv"),
    "2400": (f"{_VB}/Verbindlichkeiten gegenueber Kreditinstituten", "bilanz_passiv"),
    "2440": (f"{_VB}/Erhaltene Anzahlungen auf Bestellungen", "bilanz_passiv"),
    "2460": (f"{_VB}/Verbindlichkeiten aus Lieferungen und Leistungen", "bilanz_passiv"),
    "2480": (f"{_VB}/Verbindlichkeiten gegenueber verbundenen Unternehmen", "bilanz_passiv"),
    "2510": (f"{_VB}/Sonstige Verbindlichkeiten", "bilanz_passiv"),
    "2600": (f"{_P}/D Rechnungsabgrenzungsposten", "bilanz_passiv"),
    # Gewinn- und Verlustrechnung
    "3010": (f"{_G}/Umsatzerloese", "guv"),
    "3025": (f"{_G}/Bestandsveraenderung", "guv"),
    "3040": (f"{_G}/Sonstige betriebliche Ertraege", "guv"),
    "3100": (f"{_G}/Materialaufwand", "guv"),
    "3150": (f"{_G}/Personalaufwand/Loehne und Gehaelter", "guv"),
    "3200": (f"{_G}/Abschreibungen", "guv"),
    "3210": (f"{_G}/Sonstige betriebliche Aufwendungen", "guv"),
    "3290": (f"{_G}/Sonstige Zinsen und aehnliche Ertraege", "guv"),
    "3320": (f"{_G}/Zinsen und aehnliche Aufwendungen", "guv"),
    "3400": (f"{_G}/Steuern vom Einkommen und vom Ertrag", "guv"),
    # Der Verlustvortrag steht im GuV-Nachweis, ist aber Eigenkapital.
    "3510": (f"{_EK}/IV Gewinnvortrag Verlustvortrag", "bilanz_passiv"),
}

#: Positionen, die nur eine Zwischensumme sind und keine eigenen Konten tragen.
_NUR_SUMME = {"3420", "3430", "3800", "2115", "2125"}

_BETRAG = r"-?[\d.]+,\d{2}"
_POS = re.compile(rf"^(?P<pos>\d{{4}})\s+\*\*\*\s+(?P<bez>.*?)\s*(?P<betrag>{_BETRAG})?$")
_KONTO = re.compile(rf"^(?P<konto>\d{{5,6}})\s+(?P<bez>.*?)\s*(?P<betrag>{_BETRAG})?$")
_SUMME = re.compile(rf"^\*\*\*\s+SUMME\s+(?P<seite>AKTIVA|PASSIVA)\s+(?P<betrag>{_BETRAG})$")
_NUR_BETRAG = re.compile(rf"^(?P<rest>.*?)\s*(?P<betrag>{_BETRAG})$")
_STICHTAG = re.compile(r"zum (\d{1,2})\.\s*(\w+)\s*(\d{4})")

_MONATE = {"Januar": 1, "Februar": 2, "März": 3, "April": 4, "Mai": 5, "Juni": 6,
           "Juli": 7, "August": 8, "September": 9, "Oktober": 10,
           "November": 11, "Dezember": 12}


def _betrag(s: str) -> float:
    return float(s.replace(".", "").replace(",", "."))


@dataclass
class KNZeile:
    konto: str
    bezeichnung: str
    pos: str
    sektion: str                 # "AKTIVA" | "PASSIVA" | "GUV"
    betrag: float

    @property
    def hgb_pfad(self) -> Optional[str]:
        eintrag = POSITIONEN.get(self.pos)
        return eintrag[0] if eintrag else None

    @property
    def kontotyp(self) -> Optional[str]:
        eintrag = POSITIONEN.get(self.pos)
        return eintrag[1] if eintrag else None

    def vorzeichenrichtig(self) -> float:
        """Databook-Konvention: Soll positiv, Haben negativ.

        Der Abschluss druckt die Passivseite positiv (mit dem Verlust als
        Minus) und die Aufwendungen negativ. Beide Seiten werden gedreht;
        die Aktivseite bleibt, wie sie steht."""
        return -self.betrag if self.sektion in ("PASSIVA", "GUV") else self.betrag


@dataclass
class BrehnaAbschluss:
    periode: str
    stichtag: str
    entity: str
    zeilen: list[KNZeile] = field(default_factory=list)
    #: Positionssummen laut Abschluss, zur Selbstkontrolle des Parsers.
    positionssummen: dict[str, float] = field(default_factory=dict)
    bilanzsumme: dict[str, float] = field(default_factory=dict)
    quelle_datei: str = ""
    ist_zwischenabschluss: bool = False
    warnungen: list[str] = field(default_factory=list)

    def summe(self, sektion: str) -> float:
        return sum(z.betrag for z in self.zeilen if z.sektion == sektion)

    def probe(self) -> list[tuple[str, bool, str]]:
        """Parser-Selbstkontrolle (QA C5): jede gelesene Position gegen ihre
        gedruckte Summe, und die Seitensummen gegen die gedruckte Bilanzsumme."""
        tests: list[tuple[str, bool, str]] = []
        for pos, gedruckt in sorted(self.positionssummen.items()):
            gelesen = sum(z.betrag for z in self.zeilen if z.pos == pos)
            if pos in _NUR_SUMME and abs(gelesen) < 0.005:
                continue          # Position ohne Kontenebene
            ok = abs(gelesen - gedruckt) <= 0.01
            tests.append((f"Position {pos}", ok,
                          f"gelesen {gelesen:,.2f} gegen gedruckt {gedruckt:,.2f}"))
        for seite in ("AKTIVA", "PASSIVA"):
            gedruckt = self.bilanzsumme.get(seite)
            if gedruckt is None:
                continue
            # Die Passivseite trägt den Bilanzverlust als Position OHNE eigenes
            # Konto: er setzt sich aus dem Vortrag und dem Periodenergebnis
            # zusammen, und beide sitzen im GuV-Nachweis. Die Probe muss sie
            # deshalb mitzählen, sonst ginge sie um das Ergebnis daneben.
            gelesen = self.summe(seite)
            if seite == "PASSIVA":
                gelesen += self.summe("GUV")
            ok = abs(gelesen - gedruckt) <= 0.01
            tests.append((f"Summe {seite}", ok,
                          f"gelesen {gelesen:,.2f} gegen gedruckt {gedruckt:,.2f}"
                          + (" (inkl. Periodenergebnis aus dem GuV-Nachweis)"
                             if seite == "PASSIVA" else "")))
        aktiva, passiva = self.bilanzsumme.get("AKTIVA"), self.bilanzsumme.get("PASSIVA")
        if aktiva is not None and passiva is not None:
            tests.append(("Aktiva gleich Passiva", abs(aktiva - passiva) <= 0.01,
                          f"{aktiva:,.2f} gegen {passiva:,.2f}"))
        return tests


def lies_brehna_ja(pfad: str) -> BrehnaAbschluss:
    import pdfplumber

    zeilen: list[KNZeile] = []
    positionssummen: dict[str, float] = {}
    bilanzsumme: dict[str, float] = {}
    warnungen: list[str] = []
    entity, stichtag, zwischen = "Unbekannt", "", False

    with pdfplumber.open(pfad) as pdf:
        for seite in pdf.pages:
            text = seite.extract_text() or ""
            if "Zwischenabschluss" in text:
                zwischen = True
            if entity == "Unbekannt":
                m = re.search(r"^(Brehna .*?)\s*$", text, re.M)
                if m:
                    entity = m.group(1).strip()
            if not stichtag:
                m = _STICHTAG.search(text)
                if m:
                    stichtag = f"{int(m.group(1)):02d}.{_MONATE.get(m.group(2), 12):02d}.{m.group(3)}"
            if "Kontennachweis" not in text:
                continue
            sektion = ("AKTIVA" if "Kontennachweis Aktivseite" in text else
                       "PASSIVA" if "Kontennachweis Passivseite" in text else "GUV")
            _lies_seite(text, sektion, zeilen, positionssummen, bilanzsumme, warnungen)

    jahr = stichtag[-4:] if stichtag else "?"
    periode = f"YTD {stichtag[3:5]}/{jahr}" if zwischen else f"FY{jahr}"
    return BrehnaAbschluss(
        periode=periode, stichtag=stichtag, entity=entity, zeilen=zeilen,
        positionssummen=positionssummen, bilanzsumme=bilanzsumme,
        quelle_datei=pfad, ist_zwischenabschluss=zwischen, warnungen=warnungen)


def _lies_seite(text: str, sektion: str, zeilen: list, positionssummen: dict,
                bilanzsumme: dict, warnungen: list) -> None:
    """Eine Kontennachweis-Seite.

    Bezeichnungen brechen um und werden dabei getrennt; der Betrag steht dann
    auf der Folgezeile. Deshalb wird eine Zeile ohne Betrag offen gehalten,
    bis die nächste Zeile den Betrag nachliefert."""
    offen: Optional[dict] = None
    pos_aktuell = ""

    for roh in text.split("\n"):
        zeile = roh.strip()
        if not zeile or zeile.startswith(("JAHRESABSCHLUSS", "Brehna", "Pos Konto",
                                          "EUR", "- ", "55543")):
            continue

        m = _SUMME.match(zeile)
        if m:
            bilanzsumme[m.group("seite")] = _betrag(m.group("betrag"))
            offen = None
            continue

        if offen is not None:
            # Fortsetzungszeile: Resttext plus Betrag.
            mb = _NUR_BETRAG.match(zeile)
            if mb:
                offen["bez"] = (offen["bez"] + " " + mb.group("rest")).strip()
                _uebernimm(offen, _betrag(mb.group("betrag")), sektion, zeilen,
                           positionssummen)
                offen = None
                continue
            offen["bez"] = (offen["bez"] + " " + zeile).strip()
            continue

        mp = _POS.match(zeile)
        if mp:
            pos_aktuell = mp.group("pos")
            eintrag = {"art": "pos", "pos": pos_aktuell, "bez": mp.group("bez")}
            if mp.group("betrag"):
                _uebernimm(eintrag, _betrag(mp.group("betrag")), sektion, zeilen,
                           positionssummen)
            else:
                offen = eintrag
            continue

        mk = _KONTO.match(zeile)
        if mk:
            eintrag = {"art": "konto", "pos": pos_aktuell,
                       "konto": mk.group("konto"), "bez": mk.group("bez")}
            if mk.group("betrag"):
                _uebernimm(eintrag, _betrag(mk.group("betrag")), sektion, zeilen,
                           positionssummen)
            else:
                offen = eintrag
            continue


def _uebernimm(eintrag: dict, betrag: float, sektion: str, zeilen: list,
               positionssummen: dict) -> None:
    if eintrag["art"] == "pos":
        positionssummen[eintrag["pos"]] = betrag
        return
    zeilen.append(KNZeile(konto=eintrag["konto"],
                          bezeichnung=_entsilbe(eintrag["bez"]),
                          pos=eintrag["pos"], sektion=sektion, betrag=betrag))


def _entsilbe(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("- ", "")).strip()


def lies_brehna_jahre(pfade: list[str]):
    """Vier Abschlüsse zu einem Ledger. Jeder Abschluss liefert genau eine
    Periode; Konten werden über die Kontonummer zusammengeführt."""
    from ..core.model import Account, NormalizedLedger, PeriodBalance

    abschluesse = sorted((lies_brehna_ja(p) for p in pfade),
                         key=lambda a: a.stichtag[-4:] + a.stichtag[3:5])
    perioden = [a.periode for a in abschluesse]
    stamm: dict[str, tuple[str, Optional[str], Optional[str]]] = {}
    werte: dict[str, dict[str, float]] = {}
    warnungen: list[str] = []

    for a in abschluesse:
        for z in a.zeilen:
            if z.hgb_pfad is None:
                warnungen.append(
                    f"[{a.periode}] Konto {z.konto} '{z.bezeichnung[:34]}' steht "
                    f"unter Position {z.pos}, für die kein HGB-Pfad hinterlegt "
                    "ist — fällt auf die übrige Kaskade zurück.")
            stamm[z.konto] = (z.bezeichnung, z.hgb_pfad, z.kontotyp)
            werte.setdefault(z.konto, {})[a.periode] = (
                werte.get(z.konto, {}).get(a.periode, 0.0) + z.vorzeichenrichtig())

    accounts = [
        Account(konto=k, bezeichnung=b, fs_pfad=pf, kontotyp=t,
                entity=abschluesse[0].entity,
                salden=tuple(PeriodBalance(p, werte[k].get(p, 0.0)) for p in perioden))
        for k, (b, pf, t) in sorted(stamm.items())
    ]
    ledger = NormalizedLedger(
        accounts=accounts, perioden=perioden, entity=abschluesse[0].entity,
        quelle_datei=" + ".join(pfade), hat_kontennachweis=True,
        fingerprint=fingerprint(pfade[0]), warnungen=warnungen)
    return ledger, abschluesse
