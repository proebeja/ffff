"""Projekt Luma — Databook aus MYOB-Exporten, Option A.

Vier Quellen: die Saldenliste je Geschäftsjahr, die GuV je Geschäftsjahr, die
ungeprüften Abschlüsse dreier Jahre und der Entwurf für FY2026. Was daraus
folgt:

* **Zwei Exporte, ein Mastersheet.** Bilanz und GuV kommen aus getrennten
  Dateien mit gleichem Aufbau. Zusammengeführt wird über die Kontonummer, die
  Perioden werden nach Stichtag sortiert — die Tabs des GuV-Exports stehen als
  FY2023, FY2026, FY2024, FY2025.

* **Das Ergebniskonto fällt heraus.** ``3-90000 Current Earnings`` ist die GuV
  in einer Zahl. Mit der GuV daneben stünde das Ergebnis zweimal in der
  Mappe, und die Summe aller Konten ginge nicht mehr auf null.

* **Die GuV folgt der Kostenartengliederung der Vorlage**, weil die Quelle so
  gebaut ist. Abschreibungen, Zinsen und Raumkosten werden dabei kontoweise
  aus der Sammelgruppe ``Administration`` gehoben; ohne das wäre das EBITDA um
  die Abschreibungen falsch.

* **Der Abschluss ist Abstimmziel, nicht Strukturquelle.** Er ist
  konsolidiert, die Saldenliste ist eine Division. Die Überleitung weist die
  Differenz je Periode aus, statt sie wegzudefinieren.

* **Geschäftsjahr zum 31. März**, FY2023 als Rumpfjahr über neun Monate. Die
  Zeitachse der Vorlage wird aus den Stichtagen des Exports gesetzt, nicht
  aus den Tabellennamen.

Gebaut wird nach **Option A**: die Lead-Tabs ziehen ihre Positionssummen per
``SUMIFS`` direkt aus dem Mastersheet, die Kontoslots ebenso.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Optional

from .core.hausconvention import Hausconvention
from .core.model import Klasse
from .engine.cascade import Engine
from .engine.laufprotokoll import Laufprotokoll
from .engine.qa import FLAG, ABBRUCH, QAReport, _r2
from .engine.spalten_status import SpaltenStatus, StatusMatrix
from .engine.v28 import (loese_saldenvortraege, pruefe_verhalten,
                         setze_vorlaeufige_pfade, wende_seitenwechsel_an)
from .export import vorlage
from .readers.bt_abschluss_pdf import Abschlusszahlen, lies_abschluesse
from .readers.myob_susa import lies_myob
from .views.leads import baue_lead_na, baue_lead_pl
from .views.net_debt import baue_net_debt
from .views.review_queue import baue_review_queue
from .views.working_capital import baue_working_capital

#: Das Konto, das in dieser Quelle das Periodenergebnis trägt.
ERGEBNISKONTO = "3-90000"

#: Berichtssprache dieses Mandats. Sie steuert zweierlei: die Vorlage blendet
#: den englischen Kopfblock und die englische Bezeichnungsspalte ein, und die
#: Arbeitsblätter dieses Laufs werden englisch beschriftet. Der Code selbst
#: bleibt deutsch kommentiert — das ist die Hausregel und keine Frage der
#: Berichtssprache.
SPRACHE = "en"

#: Statuszeile je Spalte, in der Berichtssprache. ``ABGELEITET`` der
#: Hausconvention lautet "abgeleitet — nicht abschlusstreu".
_STATUS_ABGELEITET = ("derived from the system hierarchy — not tied to "
                      "audited financial statements")
_STATUS_KEINE_GUV = "no P&L in the source — classes 10 to 50 are balance sheet only"
_STATUS_GUV = ("derived from the system hierarchy — reconciled against the "
               "consolidated financial statements, see Reconciliation")


@dataclass
class Quellen:
    saldenliste: str
    guv: Optional[str] = None
    abschluss_geprueft: Optional[str] = None
    abschluss_entwurf: Optional[str] = None


def run(quellen: Quellen, ausgabe: str, verbose: bool = True) -> dict:
    lp = Laufprotokoll()
    with lp.phase("Setup") as d:
        hc = Hausconvention.laden()
        d["detail"] = f"Hausconvention v{hc.version}, Architektur Option A"

    with lp.phase("Einlesen Saldenliste und GuV (MYOB)") as d:
        ledger, diag = lies_myob(quellen.saldenliste, quellen.guv,
                                 entity="Projekt Luma")
        d["detail"] = (f"{len(ledger.perioden)} Perioden, "
                       f"{len(ledger.accounts)} Konten, "
                       f"{sum(diag.kontozeilen.values())} Zeilen")

    with lp.phase("Einlesen Abschlüsse (PDF)") as d:
        abschluss = (lies_abschluesse(quellen.abschluss_geprueft,
                                      quellen.abschluss_entwurf)
                     if quellen.abschluss_geprueft else Abschlusszahlen())
        d["detail"] = f"{len(abschluss.stichtage)} Stichtage"

    with lp.phase("Mapping (Kaskade)") as d:
        mapped = Engine(hc).map_ledger(ledger)
        d["detail"] = f"{len(mapped)} Konten"

    with lp.phase("v2.8-Nachlauf") as d:
        mapped, saldenvortrag = loese_saldenvortraege(mapped, ledger.perioden, hc)
        mapped, seitenwechsel = wende_seitenwechsel_an(
            mapped, ledger.perioden, hc, _nachgewiesene_seiten(mapped, ledger))
        mapped, ungeloest = setze_vorlaeufige_pfade(mapped, ledger.perioden, hc)
        verhalten = pruefe_verhalten(mapped, ledger.perioden, hc,
                                     _schwelle(mapped, ledger.perioden, hc))
        d["detail"] = (f"{len(seitenwechsel)} Seitenwechsel, {len(ungeloest)} "
                       f"vorläufige Pfade, {len(verhalten)} Verhaltensbefunde")

    status = _status(diag, diag.hat_guv)

    with lp.phase("Views") as d:
        nd = baue_net_debt(mapped, ledger.perioden, ledger.entity)
        wc = baue_working_capital(mapped, ledger.perioden, ledger.entity)
        lead_na = baue_lead_na(mapped, ledger.perioden, ledger.entity)
        lead_pl = baue_lead_pl(mapped, ledger.perioden, ledger.entity)
        review = baue_review_queue(mapped, ledger.perioden, SPRACHE)
        d["detail"] = f"{len(lead_na.bloecke)} NA-Blöcke, {len(review)} Review"

    ergebnis = _periodenergebnis(mapped, ledger.perioden)
    ueberleitung = _ueberleitung(mapped, ledger.perioden, diag, abschluss,
                                 ergebnis)

    with lp.phase("QA-Diagnose") as d:
        qa = _qa(mapped, ledger, diag, status, ungeloest, ueberleitung)
        d["detail"] = (f"{len(qa.pruefungen)} Prüfungen, "
                       f"{len(qa.durchgefallen)} nicht bestanden")

    achse = vorlage.Zeitachse([
        vorlage.Periode(p, diag.stichtage[p], "jahr") for p in ledger.perioden])

    meta = _meta(quellen, ledger, mapped, diag, status, review, hc)
    with lp.phase("Vorlage befüllen (Dealtool, Option A)") as d:
        befuellt = vorlage.schreibe_dealtool(
            ausgabe, mapped=mapped, perioden=ledger.perioden,
            mandat=vorlage.Mandat(
                projekt="Project Luma", waehrung="AUD",
                architektur="option_a", sprache=SPRACHE,
                quelle_de="Saldenliste je Geschäftsjahr (MYOB-Export)",
                quelle_en="Annual trial balance per financial year "
                          "(MYOB export)"),
            achse=achse, periodenergebnis=ergebnis,
            # Zeile 264 der Vorlage heißt "Jahresergebnis lt. Quelldatei".
            # Quelldatei ist hier der Abschluss, nicht die Saldenliste — die
            # Check-Zeile darunter wird damit zur Überleitung.
            ergebnis_lt_quelle=_ergebnis_lt_abschluss(diag, abschluss,
                                                      ledger.perioden),
            review=review,
            zusatzblaetter=_zusatzblaetter(qa, status, verhalten, diag,
                                           ueberleitung, ledger.perioden))
        d["detail"] = (f"{os.path.basename(ausgabe)}, {befuellt.zellen} Zellen, "
                       f"{len(befuellt.zeilen)} Mastersheet-Zeilen")
    lp.zaehle_arbeitsmappe(ausgabe)

    kontrollen = _kontrollen(mapped, ledger.perioden, befuellt, ergebnis)
    if verbose:
        _zusammenfassung(meta, ledger, diag, status, ausgabe, lp, befuellt,
                         kontrollen, review, ueberleitung)
    return {"laufprotokoll": lp, "ledger": ledger, "diagnose": diag,
            "mapped": mapped, "qa": qa, "status": status, "nd": nd, "wc": wc,
            "lead_na": lead_na, "lead_pl": lead_pl, "review": review,
            "verhalten": verhalten, "meta": meta, "befuellt": befuellt,
            "kontrollen": kontrollen, "ergebnis": ergebnis,
            "abschluss": abschluss, "ueberleitung": ueberleitung,
            "seitenwechsel": seitenwechsel, "ungeloest": ungeloest,
            "saldenvortrag": saldenvortrag}


def _periodenergebnis(mapped, perioden) -> dict[str, float]:
    """Das Ergebnis der Periode aus den GuV-Konten.

    Vorzeichen: die Saldenliste führt Erträge im Haben, also negativ. Ein
    Gewinn ist damit eine negative Summe und erhöht das Nettovermögen — das
    Ergebnis ist die Summe mit umgekehrtem Vorzeichen.

    Fehlt die GuV, tritt das Eigenkapitalkonto ``Current Earnings`` an ihre
    Stelle; es trägt dasselbe Ergebnis in einer Zahl.
    """
    guv = [m for m in mapped if m.klasse is Klasse.PL]
    if guv:
        return {p: round(-sum(m.saldo(p) for m in guv), 2) for p in perioden}
    konto = next((m for m in mapped if m.konto == ERGEBNISKONTO), None)
    if konto is None:
        return {}
    return {p: round(-konto.saldo(p), 2) for p in perioden}


def _ergebnis_lt_abschluss(diag, abschluss, perioden) -> dict[str, float]:
    """Periodenergebnis laut Abschluss, für die Gegenprobe im Lead PL."""
    if not abschluss.stichtage:
        return {}
    werte = {}
    for p in perioden:
        tag = diag.stichtage.get(p)
        wert = abschluss.wert("guv", tag, "Profit (Loss) for the year") if tag \
            else None
        if wert is not None:
            werte[p] = wert
    return werte


def _nachgewiesene_seiten(mapped, ledger) -> dict[str, dict[str, str]]:
    """Die Bilanzseite je Konto und Periode, wie der Export sie führt.

    Der Export weist jedes Konto in jeder Periode derselben Klasse zu; es gibt
    also keinen Seitenwechsel zu entdecken, und eine Ableitung aus dem
    Vorzeichen darf ihn auch nicht erfinden. Genau dafür ist der Parameter da.
    """
    je_konto: dict[str, dict[str, str]] = {}
    for m in mapped:
        seite = "AKTIVA" if m.hgb_pfad.startswith("/Aktiva") else "PASSIVA"
        je_konto[m.konto] = {p: seite for p in ledger.perioden}
    return je_konto


def _schwelle(mapped, perioden, hc) -> float:
    prozent = hc.wesentlichkeit.get("bilanzsumme_prozent", 2) / 100.0
    summen = [sum(m.saldo(p) for m in mapped if m.hgb_pfad.startswith("/Aktiva"))
              for p in perioden]
    return max(summen or [0.0]) * prozent


def _status(diag, hat_guv: bool = False) -> StatusMatrix:
    """Jede Spalte ist abgeleitet, keine ist abschlusstreu.

    Es liegt kein Abschluss vor, gegen den sich das Databook überleiten
    ließe. Die Gliederung stammt aus der Systemgliederung des Exports — das
    trägt die Bilanz, aber es ersetzt keinen Kontennachweis.
    """
    spalten = []
    for p in diag.perioden:
        monate = diag.monate.get(p)
        spalten.append(SpaltenStatus(
            periode=p, bilanz=_STATUS_ABGELEITET,
            guv=_STATUS_GUV if hat_guv else _STATUS_KEINE_GUV,
            quelle=f"MYOB trial balance, closing balances as at "
                   f"{diag.stichtage[p].strftime('%d %b %Y')}",
            hinweis=(f"Short year of {monate} months."
                     if monate and monate != 12 else "")))
    return StatusMatrix(spalten=spalten)


def _qa(mapped, ledger, diag, status, ungeloest, ueberleitung=None) -> QAReport:
    """QA für eine Quelle ohne Abschluss. Was gegenstandslos ist, wird als
    gegenstandslos ausgewiesen und nicht stillschweigend weggelassen.

    Die Befundtexte stehen in der Berichtssprache des Mandats, weil sie im
    Databook landen. Die Kommentare bleiben deutsch."""
    r = QAReport()
    perioden = ledger.perioden

    r.add("A1", "Data block boundary unambiguous", True, ABBRUCH,
          "The export is a plain list of accounts per sheet, with no subtotal "
          "or summary rows.",
          [f"{p}: {n} rows" for p, n in diag.kontozeilen.items()])

    summen = {p: _r2(sum(m.saldo(p) for m in mapped)) for p in perioden}
    r.add("A2", "Balance sheet identity per period column (tolerance 1.00)",
          all(abs(v) <= 1.0 for v in summen.values()), FLAG,
          "Sum of all accounts per column: "
          + ", ".join(f"{p} = {v:,.2f}" for p, v in summen.items()))

    r.add("A3", "Account keys unique", True, FLAG,
          f"{len(mapped)} accounts. The export carries {len(diag.mehrfach)} of "
          "them once per cost centre; they are summed per account.",
          [f"{k}: {n} cost centres" for k, n in sorted(diag.mehrfach.items())])

    r.add("A4", "Account keys consistent with the chart of accounts", True, FLAG,
          "MYOB chart of accounts of the form '1-12300'. Not an SKR chart — "
          "the SKR default stage of the cascade deliberately does not apply; "
          "the structure comes from the system hierarchy of the export.")

    null = [m.konto for m in mapped
            if all(abs(m.saldo(p)) < 0.005 for p in perioden)]
    r.add("A5", "Nil accounts carried and flagged", True, FLAG,
          f"{len(null)} of {len(mapped)} accounts carry no balance in any period.")

    for p in perioden:
        r.nicht_zugeordnet[p] = _r2(sum(u.salden.get(p, 0.0) for u in ungeloest))
        r.bilanzsumme[p] = _r2(sum(m.saldo(p) for m in mapped
                                   if m.hgb_pfad.startswith("/Aktiva")))
    r.add("A6", "Unresolved accounts sit INSIDE the balance sheet",
          all(abs(v) < 1.0 for v in r.nicht_zugeordnet.values()), FLAG,
          f"{len(ungeloest)} accounts without a determinable path.",
          [f"{p}: {r.nicht_zugeordnet[p]:,.2f} of {r.bilanzsumme[p]:,.2f}"
           for p in perioden])

    r.add("B1", "Monthly columns cumulative or periodic", True, ABBRUCH,
          "Not applicable: the export delivers annual reporting dates, not a "
          "monthly matrix.")

    r.add("B2", "Year anchor", True, FLAG,
          "Every column is one reporting date per PeriodFrom/PeriodTo: "
          + ", ".join(f"{p} = {diag.stichtage[p].strftime('%d %b %Y')}"
                      for p in perioden))

    proben = diag.spaltenprobe()
    r.add("B3", "Correct value column taken", all(ok for _, ok, _ in proben),
          ABBRUCH,
          "The export carries BOTH the period movement (column L) and the "
          "closing balance (column O). The closing balance is what is read; "
          "the header calls it 'Year', which is misleading. Both columns sum "
          "to zero per period, but the movement column would produce a "
          "databook built from changes rather than balances.",
          [f"{name}: {text}" for name, _, text in proben])

    laengen = {p: diag.monate.get(p) for p in perioden}
    ungleich = len(set(laengen.values())) > 1
    r.add("B4", "Comparability of the periods", not ungleich, FLAG,
          "The periods are of unequal length." if ungleich
          else "Equal length.",
          [f"{p}: {m} months" for p, m in laengen.items()])

    r.add("B5", "Opening balance accounts present", True, FLAG,
          "Not a DATEV chart, so there are no DATEV opening balance accounts. "
          "The MYOB counterpart is '3-90500 Historical Balancing Account', "
          "carried as equity.")

    r.add("C1", "Structure source coverage per period and statement", False, FLAG,
          "Balance sheet and P&L are both DERIVED from the system hierarchy of "
          "the export. They are reconciled against the financial statements, "
          "but not tied to them: the statements are consolidated, the trial "
          "balance is one division." if diag.hat_guv else
          "The balance sheet is DERIVED from the system hierarchy, not tied "
          "to audited financial statements. There is no data for the P&L.",
          [f"{s.periode}: balance sheet {s.bilanz} · P&L {s.guv}"
           for s in status.spalten])

    r.add("C2", "Reconciliation at account level", not ueberleitung, FLAG,
          "Not possible at account level: the financial statements report "
          "consolidated totals per line item, not accounts. The bridge is "
          "drawn at line-item level on the Reconciliation sheet — net assets, "
          "revenue and the result for the period." if ueberleitung else
          "Not applicable: there is no second source of values. Without "
          "financial statements there is nothing to reconcile against.")

    r.add("C3", "Known reconciliation patterns", True, FLAG,
          "No split balances identified. Multiple rows arise from cost "
          "centres, not from split account balances.")

    r.add("C4", "Chart of accounts coverage", True, FLAG,
          "Description and account group come from the export itself; no "
          "separate chart of accounts was provided.")

    r.add("C5", "Parser self-check of the structure source",
          not diag.gruppen_ohne_zuordnung, FLAG,
          f"{len(diag.gruppen_ohne_zuordnung)} account groups without a mapping."
          if diag.gruppen_ohne_zuordnung else
          "Every account group in the export is mapped to an HGB path.",
          sorted(diag.gruppen_ohne_zuordnung))

    r.add("D1", "Third-party classification", True, FLAG,
          "None present — the export carries only its own system hierarchy. "
          "Hence no benchmark tab.")

    r.annahmen.append(
        "Signs: the export already carries assets positive and liabilities "
        "and equity negative. Nothing is reversed.")
    r.annahmen.append(
        "Cost centres are summed per account. The mastersheet carries an "
        "account exactly once; the cost centre would be a second key.")
    if diag.hat_guv:
        r.annahmen.append(
            f"Account {ERGEBNISKONTO} 'Current Earnings' is left out of the "
            "mastersheet: it is the P&L in a single figure, and with the P&L "
            "loaded the result would be counted twice.")
        r.annahmen.append(
            "The P&L follows the cost-type layout of the template, because "
            "that is how the source is structured. Depreciation, interest and "
            "occupancy costs are lifted out of the collective group "
            "'Administration' account by account — otherwise EBITDA would be "
            "wrong by the depreciation.")
    else:
        r.annahmen.append(
            f"The result for the period comes from account {ERGEBNISKONTO} "
            "'Current Earnings'. No P&L is available.")
    for konto, grund in diag.eliminiert.items():
        r.offene_befunde.append(f"Not carried into the mastersheet: {konto} — "
                                f"{grund}")
    for k, b, grund in diag.ohne_pfad:
        r.offene_befunde.append(f"Account {k} ({b}): {grund}")
    return r


def _zusatzblaetter(qa, status, verhalten, diag, ueberleitung=None,
                    perioden=None) -> dict:
    """Arbeitsblätter des Laufs, beschriftet in der Berichtssprache."""
    blaetter = {
        "QA": (["Check", "Title", "Passed", "Severity", "Finding", "Details"],
               [[p.id, p.titel, "yes" if p.bestanden else "NO", p.schwere,
                 p.befund, " | ".join(p.details)] for p in qa.pruefungen]),
        "Status by column": (
            ["Period", "Reporting date", "Months", "Balance sheet", "P&L",
             "Source", "Note"],
            [[s.periode, diag.stichtage[s.periode].strftime("%d %b %Y"),
              diag.monate.get(s.periode), s.bilanz, s.guv, s.quelle, s.hinweis]
             for s in status.spalten]),
        "Assumptions": (["Assumption"], [[a] for a in qa.annahmen]),
    }
    if ueberleitung and perioden:
        kopf = ["Position", "Statement", "Source"]
        for p in perioden:
            kopf += [f"{p} databook", f"{p} accounts", f"{p} difference"]
        zeilen = []
        for u in ueberleitung:
            z = [u.position, u.rechenwerk, "consolidated financial statements"]
            for p in perioden:
                z += [round(u.databook.get(p, 0.0), 2),
                      round(u.abschluss.get(p, 0.0), 2), u.differenz(p)]
            zeilen.append(z)
        blaetter["Reconciliation"] = (kopf, zeilen)
    if qa.offene_befunde:
        blaetter["Open items"] = (["Open item"],
                                  [[b] for b in qa.offene_befunde])
    if verhalten:
        blaetter["Behaviour check"] = (
            ["Account", "Description", "Class", "Criterion", "material", "Note"],
            [[v.konto, v.bezeichnung, v.klasse, v.kriterium,
              "yes" if v.wesentlich else "no", v.hinweis] for v in verhalten])
    return blaetter


@dataclass
class Ueberleitungszeile:
    """Eine Position im Abgleich Databook gegen Abschluss."""

    position: str
    rechenwerk: str                 # "balance sheet" | "P&L"
    databook: dict[str, float]
    abschluss: dict[str, float]

    def differenz(self, periode: str) -> float:
        return round(self.databook.get(periode, 0.0)
                     - self.abschluss.get(periode, 0.0), 2)


def _ueberleitung(mapped, perioden, diag, abschluss, ergebnis
                  ) -> list[Ueberleitungszeile]:
    """Databook gegen Abschluss, je Periode.

    Die Abschlüsse sind **konsolidiert** ("and its controlled entities"), die
    Saldenliste ist eine Division. Eine Differenz ist deshalb zu erwarten und
    wird ausgewiesen, nicht wegdefiniert. Die Überleitung ist damit keine
    Kontrolle, die auf null gehen muss, sondern die Aussage, wie weit das
    Databook vom Abschluss entfernt ist und in welche Richtung.
    """
    if not abschluss.stichtage:
        return []

    def summe(klassen, p):
        return sum(m.saldo(p) for m in mapped if m.klasse in klassen)

    netto = {p: round(summe((Klasse.FA, Klasse.TWC, Klasse.OWC, Klasse.ND,
                             Klasse.DT), p), 2) for p in perioden}
    umsatz = {p: round(-sum(m.saldo(p) for m in mapped
                            if m.na_de == "Umsatzerloese"), 2)
              for p in perioden}

    def aus_abschluss(werk, position):
        return {p: abschluss.wert(werk, diag.stichtage[p], position) or 0.0
                for p in perioden if p in diag.stichtage}

    return [
        Ueberleitungszeile("Net assets", "balance sheet", netto,
                           aus_abschluss("bilanz", "NET ASSETS")),
        Ueberleitungszeile("Revenue", "P&L", umsatz,
                           aus_abschluss("guv", "Sales")),
        Ueberleitungszeile("Result for the period", "P&L", ergebnis,
                           aus_abschluss("guv", "Profit (Loss) for the year")),
    ]


@dataclass
class Kontrolle:
    name: str
    je_periode: dict[str, float]
    toleranz: float = 1.0

    @property
    def ok(self) -> bool:
        return all(abs(v) <= self.toleranz for v in self.je_periode.values())


def _kontrollen(mapped, perioden, befuellt, ergebnis) -> list[Kontrolle]:
    def summe(klassen, p):
        return sum(m.saldo(p) for m in mapped if m.klasse in klassen)

    netto = {p: summe((Klasse.FA, Klasse.TWC, Klasse.OWC, Klasse.ND, Klasse.DT), p)
             for p in perioden}
    eigen = {p: summe((Klasse.EQ,), p) for p in perioden}

    rw = befuellt.roll_forward
    rf: dict[str, float] = {}
    vorher = rw.anfangsbestand or 0.0
    for p in perioden:
        bewegung = sum(w.get(p, 0.0) for w in rw.bewegungen.values())
        rf[p] = round(vorher + bewegung + ergebnis.get(p, 0.0)
                      + rw.rest.get(p, 0.0) - netto[p], 2)
        vorher = netto[p]

    # Konten, die in keiner Position des Lead landen. Sie stehen im
    # Mastersheet, die Bilanz geht im Modell auf — und im Lead fehlt der
    # Betrag. Ohne diese Zeile bliebe das unsichtbar.
    verloren = {(z.ziel_na, z.klasse) for z in befuellt.ohne_zeile}
    fehlbetrag = {p: round(sum(z.werte.get(p, 0.0) for z in befuellt.zeilen
                               if (z.na_zeile, z.klasse) in verloren), 2)
                  for p in perioden}

    kontrollen = [
        Kontrolle("Balance sheet identity (sum of all accounts = 0)",
                  {p: round(sum(m.saldo(p) for m in mapped), 2) for p in perioden}),
        # Das Periodenergebnis gehört in die Identität, sobald die GuV
        # geladen ist: es steht dann auf den GuV-Konten und nicht mehr im
        # Eigenkapital.
        Kontrolle("Lead NA: net assets + equity + result = 0",
                  {p: round(netto[p] + eigen[p] - ergebnis.get(p, 0.0), 2)
                   for p in perioden}),
        Kontrolle("Accounts with no line item in the lead", fehlbetrag),
        Kontrolle("Equity roll forward incl. residual = net assets", rf),
        Kontrolle("Unexplained movement in equity (residual)",
                  dict(rw.rest)),
    ]

    # Nur die Konten zählen, die tatsächlich keinen Slot bekommen haben, und
    # sie saldieren: die Position zeigt weiterhin den vollen Betrag, es fehlt
    # das Detail. Die Bruttosumme wäre bei Anlagekonten irreführend, weil
    # Anschaffungswert und kumulierte Abschreibung gegeneinander stehen.
    ohne_slot = {z for b in befuellt.slotbefunde for z in b.ohne_slot}
    ohne = {p: round(sum(z.werte.get(p, 0.0) for z in befuellt.zeilen
                         if z.schluessel in ohne_slot), 2) for p in perioden}
    kontrollen.append(Kontrolle(
        "Accounts without an account slot (value with no visible detail)",
        ohne, toleranz=0.0))
    return kontrollen


def _meta(quellen, ledger, mapped, diag, status, review, hc) -> dict:
    return {
        "Source file": os.path.basename(quellen.saldenliste),
        "Reader": "myob_susa (four annual tabs)",
        "Entity": "Project Luma",
        "Currency": "AUD",
        "Periods": ", ".join(
            f"{p} as at {diag.stichtage[p].strftime('%d %b %Y')} "
            f"({diag.monate.get(p)} m)" for p in ledger.perioden),
        "Architecture": "Option A — two layers, leads pull from the mastersheet",
        "Reporting language": SPRACHE,
        "Status by column": " | ".join(
            f"{s.periode}: balance sheet {s.bilanz} · P&L {s.guv}"
            for s in status.spalten),
        "Structure source": "system hierarchy of the export (ClassDescription / "
                            "AccountGroupDesc), translated in the reader",
        "Hausconvention version": hc.version,
        "Fingerprint (SHA-256/16)": ledger.fingerprint,
        "Accounts total": len(mapped),
        "of which review": len(review),
        "Accounts without a path": len(diag.ohne_pfad),
        "Reader warnings": len(ledger.warnungen),
    }


def _zusammenfassung(meta, ledger, diag, status, ausgabe, lp, befuellt,
                     kontrollen, review, ueberleitung=None) -> None:
    """Laufbericht auf der Konsole, in der Berichtssprache des Mandats."""
    print("=" * 78)
    for k, v in meta.items():
        print(f"  {k:26} {v}")
    print("-" * 78)
    for s in status.spalten:
        print(f"  {s.periode:8} balance sheet: {s.bilanz}")
        print(f"           P&L: {s.guv}")
        if s.hinweis:
            print(f"           {s.hinweis}")
    print("-" * 78)
    for k in kontrollen:
        print(f"  {'ok ' if k.ok else '!! '}{k.name}")
        print("      " + "  ".join(f"{p}: {v:,.2f}"
                                   for p, v in k.je_periode.items()))
    if ueberleitung:
        print("-" * 78)
        print("  Reconciliation against the consolidated financial statements")
        for u in ueberleitung:
            print(f"    {u.position} ({u.rechenwerk})")
            for p in ledger.perioden:
                print(f"      {p:8s} databook {u.databook.get(p, 0.0):>14,.0f}"
                      f"   accounts {u.abschluss.get(p, 0.0):>14,.0f}"
                      f"   difference {u.differenz(p):>13,.0f}")
    for z in befuellt.ohne_zeile:
        print(f"  [finding] No line item in the lead: {z.ziel_na} "
              f"({z.klasse}), {z.konten} accounts — {z.grund}")
    rw = befuellt.roll_forward
    if rw.hat_rest:
        print("-" * 78)
        print("  What the residual in equity consists of (movement towards "
              "net assets):")
        for konto, werte in rw.ergebniskonten.items():
            print(f"    {konto[:46]:46s} "
                  + "  ".join(f"{v:>14,.2f}" for v in werte.values()))
    if befuellt.slotbefunde:
        print("-" * 78)
        letzte = ledger.perioden[-1]
        for b in befuellt.slotbefunde:
            wert = sum(z.werte.get(letzte, 0.0) for z in befuellt.zeilen
                       if z.schluessel in b.ohne_slot)
            art = ("Line item created from a dummy row, no account slots"
                   if b.aus_dummy else "Account slots insufficient")
            print(f"  [finding] {art}: {b.position} ({b.klasse}), row "
                  f"{b.zeile} — {b.konten} accounts, {b.slots} slots, "
                  f"{len(b.ohne_slot)} without detail ({wert:,.2f} in {letzte})")
    if review:
        print("-" * 78)
        print(f"  Review queue: {len(review)} accounts")
        for e in review[:12]:
            print(f"    {e.konto:9s} {e.bezeichnung[:38]:38s} {e.status}")
    print("-" * 78)
    print("\n".join("  " + z for z in lp.als_text()))
    print("-" * 78)
    print(f"  Dealtool written: {ausgabe}")
    print("=" * 78)
