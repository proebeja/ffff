"""Projekt Luma — jährliche Saldenliste aus einem MYOB-Kontenplan.

Diese Quelle unterscheidet sich in vier Punkten von allem bisherigen, und
jeder davon entscheidet über das Databook:

* **Der Kontenplan ist kein SKR.** Kontonummern wie ``1-12300`` und
  Bezeichnungen wie ``Trade Debtors - USD`` treffen weder die
  SKR03-Bereichstabelle noch die deutschen Stichworte der Typ-1-Regeln. Der
  Export bringt aber seine eigene Gliederung mit (``ClassDescription`` und
  ``AccountGroupDesc``), und die wird hier auf HGB-Pfade übersetzt — genau
  wie beim SAP-BW-Export, dessen FS-Hierarchie denselben Dienst tut. Damit
  greift Stufe 1 der Kaskade, und Klasse und NA-Zeile leitet die
  Reklassifizierung wie gewohnt ab.

* **Die Wertespalte ist nicht die erste Zahlenspalte.** Der Export führt in
  Spalte L die Bewegung der Periode und in Spalte O den Schlussbestand. Die
  Kopfzeile nennt Spalte O irreführend ``Year``. Wer L nimmt, baut ein
  Databook aus Veränderungen: die Aktivseite läge in FY2023 bei 260 TAUD
  statt bei 10,4 Mio. Die Probe steht in :meth:`MyobDiagnose.spaltenprobe`.

* **Das Geschäftsjahr endet am 31. März.** FY2023 ist zudem ein Rumpfjahr
  von neun Monaten (01.07.2022 bis 31.03.2023). Periodenlänge und Stichtag
  werden deshalb aus den Spalten ``PeriodFrom``/``PeriodTo`` gelesen und
  nicht aus dem Tabellennamen geraten.

* **Es gibt keine GuV.** Der Export führt nur die Klassen 10 bis 50, also
  Bilanz und Eigenkapital. Das Periodenergebnis steht als ein Betrag auf dem
  Konto ``3-90000 Current Earnings``. Ein Lead PL lässt sich daraus nicht
  füllen, und das ist eine Aussage über die Datenlage, kein Mangel des Laufs.
"""

from __future__ import annotations

import collections
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Optional

import openpyxl

from ..core.model import Account, NormalizedLedger, PeriodBalance
from .base import Reader, fingerprint

#: Spalten des Exports (1-basiert).
SP_KLASSE = 5           # ClassDescription
SP_GRUPPE = 7           # AccountGroupDesc
SP_KONTO = 8            # AccountNo
SP_KOSTENSTELLE = 9
SP_BEZEICHNUNG = 11     # GLDescription
SP_BEWEGUNG = 12        # PeriodMovementTYD
SP_SALDO = 15           # laut Kopfzeile "Year", tatsächlich der Schlussbestand
SP_VON = 16
SP_BIS = 17

_A = "/Aktiva/A Anlagevermoegen"
_UV = "/Aktiva/B Umlaufvermoegen"
_FORD = f"{_UV}/II Forderungen und sonstige Vermoegensgegenstaende"
_P = "/Passiva"

#: Kontogruppe des Exports -> HGB-Pfad. Die Zuordnung ist eine fachliche
#: Entscheidung je Gruppe und steht deshalb an einer Stelle, nicht verstreut
#: im Code.
#:
#: Zwei Grundsätze:
#:
#: * Der **Inhalt** entscheidet, nicht die Rubrik des Exports. ``Trade
#:   Facility``, ``R&D finance``, ``Finance Leases`` und die Firmenkreditkarte
#:   stehen im Export unter den kurzfristigen Verbindlichkeiten; es sind
#:   Finanzierungen und damit Net Debt. Sie werden direkt auf den
#:   Kreditinstitute-Pfad gelegt, statt über die gemischte Position
#:   ``Sonstige Verbindlichkeiten`` zu laufen — dort entschiede eine
#:   Typ-2-Regel über deutsche Stichworte, die auf englischen Bezeichnungen
#:   ohnehin nicht greifen.
#: * Wo der Inhalt **nicht** eindeutig ist, wird bewusst auf eine der drei
#:   gemischten Positionen gelegt. Dann greift die Typ-2-Logik, und was sie
#:   nicht auflöst, landet sichtbar in der Review-Queue — statt unsichtbar in
#:   einer Entscheidung, die hier niemand belegen kann.
GRUPPEN: dict[str, str] = {
    # --- Zahlungsmittel -----------------------------------------------------
    "efine": f"{_UV}/IV Kassenbestand und Guthaben bei Kreditinstituten",
    "petty cash": f"{_UV}/IV Kassenbestand und Guthaben bei Kreditinstituten",

    # --- Forderungen --------------------------------------------------------
    "trade debtors": f"{_FORD}/Forderungen aus Lieferungen und Leistungen",
    "provision for bad & doubtful debts":
        f"{_FORD}/Forderungen aus Lieferungen und Leistungen",
    "accrued income": f"{_FORD}/Forderungen aus Lieferungen und Leistungen",
    "sundry debtors": f"{_FORD}/Sonstige Vermoegensgegenstaende",
    "employee advance": f"{_FORD}/Sonstige Vermoegensgegenstaende",
    "income tax asset": f"{_FORD}/Sonstige Vermoegensgegenstaende",
    "trade-in clearing account": f"{_FORD}/Sonstige Vermoegensgegenstaende",
    "bt imaging product": f"{_FORD}/Sonstige Vermoegensgegenstaende",

    # Konzernforderungen. Zins tragende Ausleihung an eine Schwestergesellschaft,
    # in der Hausconvention Net Debt.
    "inter company loan - china":
        f"{_FORD}/Forderungen gegen verbundene Unternehmen",
    "inter company loan - aurora":
        f"{_FORD}/Forderungen gegen verbundene Unternehmen",
    "accrued interest - aurora":
        f"{_FORD}/Forderungen gegen verbundene Unternehmen",

    # --- Abgrenzung ---------------------------------------------------------
    "prepayments": "/Aktiva/C Rechnungsabgrenzungsposten",
    "rent in advance": "/Aktiva/C Rechnungsabgrenzungsposten",
    "deposits paid": "/Aktiva/C Rechnungsabgrenzungsposten",
    "loan borrowing costs": "/Aktiva/C Rechnungsabgrenzungsposten",

    # Ein Aktivkonto mit Habensaldo, das abgegrenzte Umsatzerlöse trägt. Der
    # Inhalt ist eine erhaltene Anzahlung, deshalb steht es auf der Passivseite.
    "unearned revenue - revenue recognition":
        f"{_P}/C Verbindlichkeiten/Erhaltene Anzahlungen auf Bestellungen",

    # --- Vorräte ------------------------------------------------------------
    "total inventory": f"{_UV}/I Vorraete/Roh- Hilfs- und Betriebsstoffe",
    "wip": f"{_UV}/I Vorraete/Unfertige Erzeugnisse und Leistungen",
    "finished goods": f"{_UV}/I Vorraete/Fertige Erzeugnisse und Waren",
    "finished goods-labour": f"{_UV}/I Vorraete/Fertige Erzeugnisse und Waren",

    # --- Anlagevermögen -----------------------------------------------------
    "furniture & fittings": f"{_A}/II Sachanlagen",
    "office equipment": f"{_A}/II Sachanlagen",
    "computer equipment": f"{_A}/II Sachanlagen",
    "tools & lab equipment": f"{_A}/II Sachanlagen",
    "warehouse equipment": f"{_A}/II Sachanlagen",
    "prototype": f"{_A}/II Sachanlagen",
    "leasehold property improvements": f"{_A}/II Sachanlagen",
    "right of use assets": f"{_A}/II Sachanlagen",
    "property plant & equipment": f"{_A}/II Sachanlagen",
    "us office fixed assets": f"{_A}/II Sachanlagen",
    # Die Demo- und Entwicklungswerkzeuge führt der Export je Werkzeug mit
    # eigener Gruppe für die kumulierte Abschreibung. Alle gehören zum
    # Sachanlagevermögen; die Trennung Anschaffung/Abschreibung ist eine
    # Frage des Aufrisses, nicht der Position.
    "r&d/demo tools - waterloo": f"{_A}/II Sachanlagen",
    "r&d r3-169 & 181 - acc dep": f"{_A}/II Sachanlagen",
    "r&d b3-106 acc dep": f"{_A}/II Sachanlagen",
    "r&d m1-ct-102 acc dep": f"{_A}/II Sachanlagen",
    "intellectual property": f"{_A}/I Immaterielle Vermoegensgegenstaende",
    "goodwill": f"{_A}/I Immaterielle Vermoegensgegenstaende",
    "other intangibles": f"{_A}/I Immaterielle Vermoegensgegenstaende",
    "intangibles": f"{_A}/I Immaterielle Vermoegensgegenstaende",

    # --- Latente Steuern ----------------------------------------------------
    "deferred tax asset": "/Aktiva/D Aktive latente Steuern",

    # --- Verbindlichkeiten aus L+L -----------------------------------------
    "trade creditors":
        f"{_P}/C Verbindlichkeiten/Verbindlichkeiten aus Lieferungen und Leistungen",
    "accrued trade creditors":
        f"{_P}/C Verbindlichkeiten/Verbindlichkeiten aus Lieferungen und Leistungen",

    # --- Finanzierung (Net Debt) -------------------------------------------
    "credit cards":
        f"{_P}/C Verbindlichkeiten/Verbindlichkeiten gegenueber Kreditinstituten",
    "trade facility":
        f"{_P}/C Verbindlichkeiten/Verbindlichkeiten gegenueber Kreditinstituten",
    "r&d finance":
        f"{_P}/C Verbindlichkeiten/Verbindlichkeiten gegenueber Kreditinstituten",
    "finance leases":
        f"{_P}/C Verbindlichkeiten/Verbindlichkeiten gegenueber Kreditinstituten",
    # Partners for Growth ist der Wagnisfremdkapitalgeber des Unternehmens —
    # die Warrants im Eigenkapital tragen denselben Namen. Ein Darlehen,
    # keine Lieferantenverbindlichkeit.
    "loan pfg":
        f"{_P}/C Verbindlichkeiten/Verbindlichkeiten gegenueber Kreditinstituten",
    "financial liability (convertible note)": f"{_P}/C Verbindlichkeiten/Anleihen",

    # --- Erhaltene Anzahlungen ---------------------------------------------
    "sales in advance":
        f"{_P}/C Verbindlichkeiten/Erhaltene Anzahlungen auf Bestellungen",

    # --- Rückstellungen -----------------------------------------------------
    "accrued expenses": f"{_P}/B Rueckstellungen/Sonstige Rueckstellungen",
    "provision for annual leave": f"{_P}/B Rueckstellungen/Sonstige Rueckstellungen",
    "provision for lsl": f"{_P}/B Rueckstellungen/Sonstige Rueckstellungen",

    # --- Sonstige Verbindlichkeiten ----------------------------------------
    "ato liabilities": f"{_P}/C Verbindlichkeiten/Sonstige Verbindlichkeiten",
    "payroll liabilities": f"{_P}/C Verbindlichkeiten/Sonstige Verbindlichkeiten",
    "sundry creditors": f"{_P}/C Verbindlichkeiten/Sonstige Verbindlichkeiten",
    "interest withholding payable":
        f"{_P}/C Verbindlichkeiten/Sonstige Verbindlichkeiten",

    # --- Eigenkapital -------------------------------------------------------
    "share capital": f"{_P}/A Eigenkapital",
    "series a pref shares- redemption payments": f"{_P}/A Eigenkapital",
    "warrant - partners for growth iii lp": f"{_P}/A Eigenkapital",
    "warrant - silicon valley bank": f"{_P}/A Eigenkapital",
    "ast equity interest": f"{_P}/A Eigenkapital",
    "fundraising equity costs": f"{_P}/A Eigenkapital",
    "retained earnings": f"{_P}/A Eigenkapital",
    "current earnings": f"{_P}/A Eigenkapital",
}

#: Kontonummern, die bewusst KEINEN Pfad bekommen und in die Review-Queue
#: laufen. Ein Sammel-/Verrechnungskonto ohne erkennbaren Inhalt gehört dorthin
#: und nicht in eine Bilanzposition, die jemand später für belegt hält.
OHNE_PFAD = {
    "9-99999": "Suspense Account — Verrechnungskonto ohne erkennbaren Inhalt.",
}

#: Einzelne Konten, deren Gruppe zu grob ist.
#:
#: Für die Positionszeile der Vorlage macht das nichts — sie führt eine Zeile
#: ``Vorräte``. Für den HGB-Pfad im Mastersheet schon, und ein Aufriss der
#: Vorräte nach § 266 lebt genau von dieser Trennung.
KONTEN: dict[str, str] = {
    # Payroll Liabilities, inhaltlich aber Rückstellungen.
    "2-70350": f"{_P}/B Rueckstellungen/Sonstige Rueckstellungen",
    "2-53000": f"{_P}/B Rueckstellungen/Steuerrueckstellungen",
    "2-54000": f"{_P}/B Rueckstellungen/Steuerrueckstellungen",
    # 'Total Inventory' fasst Roh-, Hilfs- und Betriebsstoffe, unfertige und
    # fertige Erzeugnisse in einer Gruppe zusammen.
    "1-51000": f"{_UV}/I Vorraete/Unfertige Erzeugnisse und Leistungen",
    "1-51500": f"{_UV}/I Vorraete/Unfertige Erzeugnisse und Leistungen",
    "1-52000": f"{_UV}/I Vorraete/Unfertige Erzeugnisse und Leistungen",
    "1-60000": f"{_UV}/I Vorraete/Fertige Erzeugnisse und Waren",
    "1-40300": f"{_UV}/I Vorraete/Fertige Erzeugnisse und Waren",
}


def _norm(text) -> str:
    """Gruppenbezeichnung vergleichbar machen. Der Export schneidet lange
    Namen ab (``Provision for Bad & Doubtful`` statt ``... Debts``), deshalb
    wird zusätzlich auf Präfix geprüft."""
    return re.sub(r"\s+", " ", str(text or "").strip().lower())


def finde_pfad(gruppe: str, konto: str) -> Optional[str]:
    """HGB-Pfad einer Kontogruppe. Kontospezifisches schlägt die Gruppe."""
    if konto in OHNE_PFAD:
        return None
    if konto in KONTEN:
        return KONTEN[konto]
    g = _norm(gruppe)
    if g in GRUPPEN:
        return GRUPPEN[g]
    # Abgeschnittene Gruppennamen: der gespeicherte Schlüssel beginnt mit dem
    # gelesenen Namen oder umgekehrt.
    for schluessel, pfad in GRUPPEN.items():
        if g and (schluessel.startswith(g) or g.startswith(schluessel)):
            return pfad
    return None


def _stichtag(bis) -> date:
    """``PeriodTo`` nennt den ERSTEN Tag des letzten Periodenmonats.

    ``2023-03-01`` ist die Periode bis einschließlich März 2023, also der
    Stichtag 31.03.2023. Wer den Wert direkt übernimmt, legt den Abschluss
    einen Monat zu früh und verschiebt die ganze Zeitachse der Vorlage.
    """
    d = bis if isinstance(bis, datetime) else datetime.fromisoformat(str(bis))
    naechster = date(d.year + (d.month == 12), d.month % 12 + 1, 1)
    return date.fromordinal(naechster.toordinal() - 1)


@dataclass
class MyobDiagnose:
    """Was der Reader über die Quelle sagen kann."""

    perioden: list[str] = field(default_factory=list)
    stichtage: dict[str, date] = field(default_factory=dict)
    monate: dict[str, int] = field(default_factory=dict)
    kontozeilen: dict[str, int] = field(default_factory=dict)
    #: Periode -> (Summe Schlussbestände, Summe Bewegungen)
    spaltensummen: dict[str, tuple[float, float]] = field(default_factory=dict)
    #: Konten, die in mehreren Zeilen (Kostenstellen) geführt werden.
    mehrfach: dict[str, int] = field(default_factory=dict)
    ohne_pfad: list[tuple[str, str, str]] = field(default_factory=list)
    gruppen_ohne_zuordnung: set = field(default_factory=set)

    def spaltenprobe(self) -> list[tuple[str, bool, str]]:
        """Belegt, dass Spalte O der Schlussbestand ist und nicht die Bewegung.

        Beide Spalten summieren sich je Periode auf null — eine Saldenliste
        tut das, und eine Bewegungsspalte auch. Unterscheiden lassen sie sich
        an der Größenordnung und daran, dass Schlussbestand der Vorperiode
        plus Bewegung den Schlussbestand ergibt.
        """
        proben: list[tuple[str, bool, str]] = []
        for p in self.perioden:
            saldo, bewegung = self.spaltensummen[p]
            proben.append((f"{p}: Bilanzidentität", abs(saldo) <= 1.0,
                           f"Summe aller Schlussbestände {saldo:,.2f}"))
            proben.append((f"{p}: Bewegungsspalte summiert ebenfalls auf null",
                           abs(bewegung) <= 1.0,
                           f"Summe aller Bewegungen {bewegung:,.2f}"))
        return proben


class MyobSusaReader(Reader):
    """Liest alle Jahrestabs einer MYOB-Saldenliste in ein Ledger."""

    name = "myob_susa"

    @classmethod
    def kann_lesen(cls, pfad: str) -> bool:
        if not pfad.lower().endswith((".xlsx", ".xlsm")):
            return False
        try:
            wb = openpyxl.load_workbook(pfad, read_only=True, data_only=True)
        except Exception:
            return False
        try:
            ws = wb.worksheets[0]
            kopf = [str(ws.cell(1, c).value or "").strip()
                    for c in range(1, 18)]
            return "AccountNo" in kopf and "PeriodMovementTYD" in kopf
        finally:
            wb.close()

    def lesen(self, pfad: str) -> NormalizedLedger:
        ledger, _ = lies_myob(pfad)
        return ledger


def lies_myob(pfad: str, entity: str = "Projekt Luma"
              ) -> tuple[NormalizedLedger, MyobDiagnose]:
    """Alle Jahrestabs -> ein Ledger plus Diagnose."""
    wb = openpyxl.load_workbook(pfad, data_only=True)
    diag = MyobDiagnose()
    warnungen: list[str] = []

    salden: dict[str, dict[str, float]] = collections.defaultdict(dict)
    stamm: dict[str, tuple[str, str, str]] = {}     # konto -> (bez, gruppe, klasse)
    zeilen_je_konto: collections.Counter = collections.Counter()

    for ws in wb.worksheets:
        periode = str(ws.title).replace(" annual", "").strip()
        diag.perioden.append(periode)
        summe_saldo = summe_bewegung = 0.0
        zeilen = 0
        je_konto: dict[str, float] = collections.defaultdict(float)

        for r in range(2, ws.max_row + 1):
            konto = ws.cell(r, SP_KONTO).value
            if not konto:
                continue
            konto = str(konto).strip()
            saldo = ws.cell(r, SP_SALDO).value or 0.0
            bewegung = ws.cell(r, SP_BEWEGUNG).value or 0.0
            # Ein Konto steht je Kostenstelle einmal in der Liste. Für das
            # Mastersheet ist die Kostenstelle keine eigene Zeile — sie wäre
            # ein zweiter Schlüssel neben der Kontonummer und liefe der
            # Single-Source-Regel zuwider.
            je_konto[konto] += float(saldo)
            summe_saldo += float(saldo)
            summe_bewegung += float(bewegung)
            zeilen += 1
            zeilen_je_konto[konto] += 1
            stamm.setdefault(konto, (
                str(ws.cell(r, SP_BEZEICHNUNG).value or "").strip(),
                str(ws.cell(r, SP_GRUPPE).value or "").strip(),
                str(ws.cell(r, SP_KLASSE).value or "").strip()))

            if diag.stichtage.get(periode) is None:
                bis, von = ws.cell(r, SP_BIS).value, ws.cell(r, SP_VON).value
                if bis:
                    diag.stichtage[periode] = _stichtag(bis)
                    if von:
                        v = von if isinstance(von, datetime) else \
                            datetime.fromisoformat(str(von))
                        diag.monate[periode] = (
                            (diag.stichtage[periode].year - v.year) * 12
                            + diag.stichtage[periode].month - v.month + 1)

        for konto, wert in je_konto.items():
            salden[konto][periode] = wert
        diag.kontozeilen[periode] = zeilen
        diag.spaltensummen[periode] = (summe_saldo, summe_bewegung)

    diag.mehrfach = {k: v // len(diag.perioden)
                     for k, v in zeilen_je_konto.items()
                     if v > len(diag.perioden)}

    accounts: list[Account] = []
    for konto in sorted(salden):
        bez, gruppe, klasse = stamm[konto]
        fs_pfad = finde_pfad(gruppe, konto)
        if fs_pfad is None:
            diag.ohne_pfad.append((konto, bez, OHNE_PFAD.get(
                konto, f"Kontogruppe '{gruppe}' ohne Zuordnung.")))
            if konto not in OHNE_PFAD:
                diag.gruppen_ohne_zuordnung.add(gruppe)
        kontotyp = ("bilanz_aktiv" if klasse.lower().endswith("assets")
                    else "bilanz_passiv")
        accounts.append(Account(
            konto=konto, bezeichnung=bez,
            salden=tuple(PeriodBalance(p, round(salden[konto].get(p, 0.0), 2))
                         for p in diag.perioden),
            entity=entity, fs_pfad=fs_pfad, kontotyp=kontotyp))

    if diag.gruppen_ohne_zuordnung:
        warnungen.append("Kontogruppen ohne Zuordnung: "
                         + ", ".join(sorted(diag.gruppen_ohne_zuordnung)))
    for periode, (saldo, _) in diag.spaltensummen.items():
        if abs(saldo) > 1.0:
            warnungen.append(f"{periode}: Bilanzidentität verfehlt um {saldo:,.2f}")

    ledger = NormalizedLedger(
        accounts=accounts, perioden=list(diag.perioden), entity=entity,
        quelle_datei=pfad, hat_kontennachweis=True,
        fingerprint=fingerprint(pfad), warnungen=warnungen)
    return ledger, diag
