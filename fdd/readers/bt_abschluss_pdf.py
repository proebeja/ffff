"""Abschlusszahlen der BT Imaging Pty Ltd als Abstimmziel.

Zwei Dateien, zwei Rechenwerke, zwei Zahlenformate:

* ``BT_Audited_Accounts_3Years.pdf`` — Bilanz und GuV zum 31. März 2023, 2024
  und 2025, Beträge als ``1,399,443`` und Negativa in Klammern.
* ``Financial_Statements_BT_Imaging_DraftFY26.pdf`` — Entwurf für das
  Geschäftsjahr 2026, quartalsweise plus YTD, Beträge als ``$248,719`` und
  ``-$3,024,664``.

Gelesen wird, was zum Abgleich gebraucht wird: die Bilanzsumme, das
Nettovermögen und das Periodenergebnis je Stichtag. Getippt wird nichts —
eine abgeschriebene Zahl ist eine Zahl ohne Beleg, und beim nächsten Entwurf
stimmt sie nicht mehr.

**Wichtig für die Überleitung:** beide Abschlüsse sind *konsolidiert* ("and
its controlled entities"), die Saldenliste ist eine Division. Die Differenz
ist deshalb erwartbar und wird ausgewiesen, nicht wegdefiniert.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from typing import Optional

import pdfplumber

#: Zahlen der beiden Dateien: ``1,399,443``, ``(8,780,521)``, ``$248,719``,
#: ``-$3,024,664``, ``-``.
_ZAHL = re.compile(r"\(?-?\$?-?[\d,]+\)?")

_MONATE = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun",
     "jul", "aug", "sep", "oct", "nov", "dec"], start=1)}


def _zahl(text: str) -> Optional[float]:
    t = text.strip()
    if t in ("-", "", "$-"):
        return 0.0
    negativ = t.startswith("(") and t.endswith(")")
    t = t.strip("()").replace("$", "").replace(",", "")
    if t.startswith("-"):
        negativ, t = True, t[1:]
    if not t.replace(".", "").isdigit():
        return None
    return -float(t) if negativ else float(t)


def _stichtag(text: str) -> Optional[date]:
    """``30-Jun -25``, ``31-Mar-26``, ``2025`` -> Datum.

    Die Kopfzeile des Entwurfs trägt Leerzeichen mitten in den Wörtern; das
    PDF ist so gesetzt. Sie werden vor dem Lesen entfernt.
    """
    t = text.replace(" ", "")
    m = re.fullmatch(r"(\d{1,2})-([A-Za-z]{3})-?(\d{2,4})", t)
    if m:
        jahr = int(m.group(3))
        jahr += 2000 if jahr < 100 else 0
        monat = _MONATE.get(m.group(2).lower())
        if monat:
            naechster = date(jahr + (monat == 12), monat % 12 + 1, 1)
            return date.fromordinal(naechster.toordinal() - 1)
    if re.fullmatch(r"(20\d{2})", t):
        # Die geprüften Abschlüsse nennen nur das Jahr; der Stichtag steht in
        # der Überschrift ("FOR THE PERIOD ENDED 31 MARCH").
        return date(int(t), 3, 31)
    return None


@dataclass
class Abschlusszahlen:
    """Positionen je Stichtag, aus beiden Dateien zusammengeführt."""

    quelle: dict[date, str] = field(default_factory=dict)
    bilanz: dict[date, dict[str, float]] = field(default_factory=dict)
    guv: dict[date, dict[str, float]] = field(default_factory=dict)

    @staticmethod
    def _schluessel(text: str) -> str:
        """Vergleichsform: das PDF setzt Leerzeichen mitten in die Wörter
        ("Cash and Cash Equ ivalents"), also wird ohne sie verglichen."""
        return "".join(str(text).split()).lower()

    def wert(self, rechenwerk: str, stichtag: date, position: str
             ) -> Optional[float]:
        werke = self.bilanz if rechenwerk == "bilanz" else self.guv
        gesucht = self._schluessel(position)
        for name, wert in werke.get(stichtag, {}).items():
            if self._schluessel(name) == gesucht:
                return wert
        return None

    @property
    def stichtage(self) -> list[date]:
        return sorted(set(self.bilanz) | set(self.guv))


def _kopfstichtage(zeile: str) -> list[date]:
    # ``30-Jun -25`` steht so im PDF. Ohne diese Naht zerfällt der Stichtag
    # beim Trennen in zwei Wörter, und die Kopfzeile liefert eine Spalte zu
    # wenig — die Werte rutschen dann um eine Spalte.
    zeile = re.sub(r"\s*-\s*", "-", zeile)
    teile = re.split(r"\s{1,}", zeile.strip())
    tage: list[date] = []
    for t in teile:
        d = _stichtag(t)
        if d:
            tage.append(d)
    return tage


def _lies_block(text: str, stichtage: list[date], ziel: dict, quelle: str,
                abschlusszahlen: Abschlusszahlen) -> None:
    """Jede Zeile mit so vielen Zahlen wie Stichtagen ist eine Position."""
    for zeile in text.split("\n"):
        zahlen = _ZAHL.findall(zeile)
        werte = [_zahl(z) for z in zahlen]
        werte = [w for w in werte if w is not None]
        if len(werte) < len(stichtage):
            continue
        # Die Beschriftung ist alles vor der ersten Zahl. Das PDF setzt
        # Leerzeichen mitten in Wörter ("Cash and Cash Equ ivalents").
        pos = _ZAHL.search(zeile)
        # Die Beschriftung bleibt, wie sie im PDF steht. Sie zu "reparieren"
        # hiesse raten, wo ein Wortende ist: aus "Profit (Loss) for the year"
        # würde "fortheyear". Verglichen wird stattdessen ohne Leerzeichen.
        label = re.sub(r"\s+", " ", zeile[:pos.start()]).strip()
        if not label or label.replace(" ", "").isdigit():
            continue
        for tag, wert in zip(stichtage, werte[-len(stichtage):]):
            ziel.setdefault(tag, {})[label] = wert
            abschlusszahlen.quelle[tag] = quelle


def lies_abschluesse(geprueft: str, entwurf: Optional[str] = None
                     ) -> Abschlusszahlen:
    """Beide Dateien -> ein Satz Abschlusszahlen je Stichtag."""
    a = Abschlusszahlen()

    with pdfplumber.open(geprueft) as pdf:
        for seite in pdf.pages:
            text = seite.extract_text() or ""
            if "STATEMENT OF FINANCIAL POSITION" in text.upper():
                tage = _kopfstichtage(
                    next(z for z in text.split("\n")
                         if re.search(r"\b20\d{2}\s+20\d{2}", z)))
                _lies_block(text, tage, a.bilanz, "audited accounts", a)
            elif "PROFIT OR LOSS" in text.upper():
                zeile = next((z for z in text.split("\n")
                              if "March 31" in z), "")
                tage = [date(int(j), 3, 31)
                        for j in re.findall(r"March 31, (\d{4})", zeile)]
                _lies_block(text, tage, a.guv, "audited accounts", a)

    if entwurf:
        with pdfplumber.open(entwurf) as pdf:
            for seite in pdf.pages:
                text = seite.extract_text() or ""
                kopf = next((z for z in text.split("\n")
                             if re.search(r"\d{1,2}-[A-Za-z]{3}\s?-\s?\d{2}", z)),
                            None)
                if not kopf:
                    continue
                tage = _kopfstichtage(kopf)
                if not tage:
                    continue
                ziel = (a.guv if "COMPREHENSIVE INCOME" in text.upper()
                        else a.bilanz)
                # Die GuV des Entwurfs führt vier Quartale und eine
                # YTD-Spalte. Für den Abgleich zählt das Jahr, also die letzte.
                if ziel is a.guv:
                    _lies_block(text, tage + ["YTD"], ziel, "draft FY26", a)
                    for pos, wert in list(ziel.get("YTD", {}).items()):
                        ziel.setdefault(tage[-1], {})[pos] = wert
                    ziel.pop("YTD", None)
                    a.quelle[tage[-1]] = "draft FY26"
                else:
                    _lies_block(text, tage, ziel, "draft FY26", a)
    return a
