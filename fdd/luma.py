"""Projekt Luma — Databook aus einer jährlichen MYOB-Saldenliste, Option A.

Die Datenlage ist die dünnste bisher, und das bestimmt, was das Databook
leisten kann:

* **Nur eine Quelle.** Es gibt keinen Jahresabschluss, keinen Kontennachweis
  und keinen Prüfbericht. Damit gibt es auch nichts abzustimmen: die
  Saldenliste ist Werte- und Strukturquelle zugleich, und die Bilanz ist
  **nicht abschlusstreu**. Sie ist aus der Systemgliederung des Exports
  **abgeleitet**, und die Zuordnung jeder Kontogruppe steht im Reader.

* **Keine GuV.** Der Export führt nur die Klassen 10 bis 50. Das
  Periodenergebnis steht als ein Betrag auf ``3-90000 Current Earnings``.
  Der Lead PL bleibt deshalb leer — und zwar sichtbar: die Check-Zeile des
  Lead PL weist das Ergebnis laut Quelle gegen eine leere GuV aus.

* **Geschäftsjahr zum 31. März**, FY2023 als Rumpfjahr über neun Monate. Die
  Zeitachse der Vorlage wird aus den Stichtagen des Exports gesetzt, nicht
  aus den Tabellennamen.

Gebaut wird nach **Option A**: die Lead-Tabs ziehen ihre Positionssummen per
``SUMIFS`` direkt aus dem Mastersheet, die Kontoslots ebenso. Für ein Mandat
ohne Normalisierungen und ohne Aufrisse ist das die kurze Kette — jede Zelle
ist in einem Klick nachvollziehbar.
"""

from __future__ import annotations

import os
from dataclasses import dataclass

from .core.hausconvention import Hausconvention
from .core.model import Klasse
from .engine.cascade import Engine
from .engine.laufprotokoll import Laufprotokoll
from .engine.qa import FLAG, ABBRUCH, QAReport, _r2
from .engine.spalten_status import ABGELEITET, SpaltenStatus, StatusMatrix
from .engine.v28 import (loese_saldenvortraege, pruefe_verhalten,
                         setze_vorlaeufige_pfade, wende_seitenwechsel_an)
from .export import vorlage
from .readers.myob_susa import lies_myob
from .views.leads import baue_lead_na, baue_lead_pl
from .views.net_debt import baue_net_debt
from .views.review_queue import baue_review_queue
from .views.working_capital import baue_working_capital

#: Das Konto, das in dieser Quelle das Periodenergebnis trägt.
ERGEBNISKONTO = "3-90000"


@dataclass
class Quellen:
    saldenliste: str


def run(quellen: Quellen, ausgabe: str, verbose: bool = True) -> dict:
    lp = Laufprotokoll()
    with lp.phase("Setup") as d:
        hc = Hausconvention.laden()
        d["detail"] = f"Hausconvention v{hc.version}, Architektur Option A"

    with lp.phase("Einlesen Saldenliste (MYOB)") as d:
        ledger, diag = lies_myob(quellen.saldenliste, entity="Projekt Luma")
        d["detail"] = (f"{len(ledger.perioden)} Perioden, "
                       f"{len(ledger.accounts)} Konten, "
                       f"{sum(diag.kontozeilen.values())} Zeilen")

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

    status = _status(diag)

    with lp.phase("Views") as d:
        nd = baue_net_debt(mapped, ledger.perioden, ledger.entity)
        wc = baue_working_capital(mapped, ledger.perioden, ledger.entity)
        lead_na = baue_lead_na(mapped, ledger.perioden, ledger.entity)
        lead_pl = baue_lead_pl(mapped, ledger.perioden, ledger.entity)
        review = baue_review_queue(mapped, ledger.perioden)
        d["detail"] = f"{len(lead_na.bloecke)} NA-Blöcke, {len(review)} Review"

    with lp.phase("QA-Diagnose") as d:
        qa = _qa(mapped, ledger, diag, status, ungeloest)
        d["detail"] = (f"{len(qa.pruefungen)} Prüfungen, "
                       f"{len(qa.durchgefallen)} nicht bestanden")

    ergebnis = _periodenergebnis(mapped, ledger.perioden)
    achse = vorlage.Zeitachse([
        vorlage.Periode(p, diag.stichtage[p], "jahr") for p in ledger.perioden])

    meta = _meta(quellen, ledger, mapped, diag, status, review, hc)
    with lp.phase("Vorlage befüllen (Dealtool, Option A)") as d:
        befuellt = vorlage.schreibe_dealtool(
            ausgabe, mapped=mapped, perioden=ledger.perioden,
            mandat=vorlage.Mandat(
                projekt="Projekt Luma", waehrung="AUD",
                architektur="option_a",
                quelle_de="Saldenliste je Geschäftsjahr (MYOB-Export)",
                quelle_en="Annual trial balance (MYOB export)"),
            achse=achse, periodenergebnis=ergebnis,
            ergebnis_lt_quelle=ergebnis, review=review,
            zusatzblaetter=_zusatzblaetter(qa, status, verhalten, diag))
        d["detail"] = (f"{os.path.basename(ausgabe)}, {befuellt.zellen} Zellen, "
                       f"{len(befuellt.zeilen)} Mastersheet-Zeilen")
    lp.zaehle_arbeitsmappe(ausgabe)

    kontrollen = _kontrollen(mapped, ledger.perioden, befuellt, ergebnis)
    if verbose:
        _zusammenfassung(meta, ledger, diag, status, ausgabe, lp, befuellt,
                         kontrollen, review)
    return {"laufprotokoll": lp, "ledger": ledger, "diagnose": diag,
            "mapped": mapped, "qa": qa, "status": status, "nd": nd, "wc": wc,
            "lead_na": lead_na, "lead_pl": lead_pl, "review": review,
            "verhalten": verhalten, "meta": meta, "befuellt": befuellt,
            "kontrollen": kontrollen, "ergebnis": ergebnis,
            "seitenwechsel": seitenwechsel, "ungeloest": ungeloest,
            "saldenvortrag": saldenvortrag}


def _periodenergebnis(mapped, perioden) -> dict[str, float]:
    """Das Ergebnis der Periode, wie die Quelle es ausweist.

    ``3-90000 Current Earnings`` trägt den Saldo des laufenden Jahres und wird
    zum Jahreswechsel geleert. Sein Saldo ist damit unmittelbar das
    Periodenergebnis — mit umgekehrtem Vorzeichen, weil ein Gewinn im
    Eigenkapital im Haben steht und das Nettovermögen erhöht.
    """
    konto = next((m for m in mapped if m.konto == ERGEBNISKONTO), None)
    if konto is None:
        return {}
    return {p: round(-konto.saldo(p), 2) for p in perioden}


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


def _status(diag) -> StatusMatrix:
    """Jede Spalte ist abgeleitet, keine ist abschlusstreu.

    Es liegt kein Abschluss vor, gegen den sich das Databook überleiten
    ließe. Die Gliederung stammt aus der Systemgliederung des Exports — das
    trägt die Bilanz, aber es ersetzt keinen Kontennachweis.
    """
    spalten = []
    for p in diag.perioden:
        monate = diag.monate.get(p)
        spalten.append(SpaltenStatus(
            periode=p, bilanz=ABGELEITET,
            guv="keine GuV in der Quelle — Klassen 10 bis 50 nur Bilanz",
            quelle=f"MYOB-Saldenliste, Schlussbestände zum "
                   f"{diag.stichtage[p].strftime('%d.%m.%Y')}",
            hinweis=(f"Rumpfjahr über {monate} Monate."
                     if monate and monate != 12 else "")))
    return StatusMatrix(spalten=spalten)


def _qa(mapped, ledger, diag, status, ungeloest) -> QAReport:
    """QA für eine Quelle ohne Abschluss. Was gegenstandslos ist, wird als
    gegenstandslos ausgewiesen und nicht stillschweigend weggelassen."""
    r = QAReport()
    perioden = ledger.perioden

    r.add("A1", "Blockgrenze des Datenblocks eindeutig", True, ABBRUCH,
          "Der Export ist eine reine Kontenliste je Blatt, ohne Zwischen- "
          "oder Summenzeilen.",
          [f"{p}: {n} Zeilen" for p, n in diag.kontozeilen.items()])

    summen = {p: _r2(sum(m.saldo(p) for m in mapped)) for p in perioden}
    r.add("A2", "Bilanzidentität je Periodenspalte (Toleranz 1 EUR)",
          all(abs(v) <= 1.0 for v in summen.values()), FLAG,
          "Summe aller Konten je Spalte: "
          + ", ".join(f"{p} = {v:,.2f}" for p, v in summen.items()))

    r.add("A3", "Kontoschlüssel eindeutig", True, FLAG,
          f"{len(mapped)} Konten. {len(diag.mehrfach)} davon führt der Export "
          "je Kostenstelle mehrfach; sie werden je Konto summiert.",
          [f"{k}: {n} Kostenstellen" for k, n in sorted(diag.mehrfach.items())])

    r.add("A4", "Kontoschlüssel im Rahmen des Kontenrahmens", True, FLAG,
          "MYOB-Kontenplan der Form '1-12300'. Kein SKR — die "
          "SKR-Default-Stufe der Kaskade greift bewusst nicht, die Struktur "
          "kommt aus der Systemgliederung des Exports.")

    null = [m.konto for m in mapped
            if all(abs(m.saldo(p)) < 0.005 for p in perioden)]
    r.add("A5", "Nullkonten mitgeführt und markiert", True, FLAG,
          f"{len(null)} von {len(mapped)} Konten ohne Saldo in allen Perioden.")

    for p in perioden:
        r.nicht_zugeordnet[p] = _r2(sum(u.salden.get(p, 0.0) for u in ungeloest))
        r.bilanzsumme[p] = _r2(sum(m.saldo(p) for m in mapped
                                   if m.hgb_pfad.startswith("/Aktiva")))
    r.add("A6", "Ungelöste Konten liegen INNERHALB der Bilanz",
          all(abs(v) < 1.0 for v in r.nicht_zugeordnet.values()), FLAG,
          f"{len(ungeloest)} Konten ohne bestimmbaren Pfad.",
          [f"{p}: {r.nicht_zugeordnet[p]:,.2f} von {r.bilanzsumme[p]:,.2f}"
           for p in perioden])

    r.add("B1", "Monatsspalten kumuliert oder periodisch", True, ABBRUCH,
          "Gegenstandslos: der Export liefert Jahresstichtage, keine "
          "Monatsmatrix.")

    r.add("B2", "Jahresanker", True, FLAG,
          "Jede Spalte ist ein Stichtag laut PeriodFrom/PeriodTo: "
          + ", ".join(f"{p} = {diag.stichtage[p].strftime('%d.%m.%Y')}"
                      for p in perioden))

    proben = diag.spaltenprobe()
    r.add("B3", "Richtige Wertespalte übernommen", all(ok for _, ok, _ in proben),
          ABBRUCH,
          "Der Export führt Bewegung (Spalte L) UND Schlussbestand (Spalte O). "
          "Gelesen wird der Schlussbestand; die Kopfzeile nennt ihn "
          "irreführend 'Year'. Beide Spalten summieren je Periode auf null, "
          "die Bewegungsspalte ergäbe aber ein Databook aus Veränderungen.",
          [f"{name}: {text}" for name, _, text in proben])

    laengen = {p: diag.monate.get(p) for p in perioden}
    ungleich = len(set(laengen.values())) > 1
    r.add("B4", "Vergleichbarkeit der Perioden", not ungleich, FLAG,
          "Die Perioden sind unterschiedlich lang." if ungleich
          else "gleich lang.",
          [f"{p}: {m} Monate" for p, m in laengen.items()])

    r.add("B5", "Saldenvortragspräsenz", True, FLAG,
          "Kein DATEV-Kontenrahmen, also keine Saldenvortragskonten. Das "
          "MYOB-Pendant ist '3-90500 Historical Balancing Account' und wird "
          "als Eigenkapital geführt.")

    r.add("C1", "Strukturquellen-Abdeckung je Periode und Rechenwerk", False, FLAG,
          "Die Bilanz ist aus der Systemgliederung ABGELEITET, nicht "
          "abschlusstreu. Für die GuV gibt es keine Daten.",
          [f"{s.periode}: Bilanz {s.bilanz} · GuV {s.guv}" for s in status.spalten])

    r.add("C2", "Reconciliation auf Kontenebene", True, FLAG,
          "Gegenstandslos: es gibt keine zweite Wertequelle. Ohne "
          "Jahresabschluss ist nichts abzustimmen.")

    r.add("C3", "Bekannte Abstimmmuster", True, FLAG,
          "Keine Saldenspaltung feststellbar. Mehrfachzeilen entstehen durch "
          "Kostenstellen, nicht durch geteilte Salden.")

    r.add("C4", "Kontenstamm-Abdeckung", True, FLAG,
          "Bezeichnung und Kontogruppe stammen aus dem Export selbst; ein "
          "separater Kontenplan liegt nicht vor.")

    r.add("C5", "Parser-Selbstkontrolle der Strukturquelle",
          not diag.gruppen_ohne_zuordnung, FLAG,
          f"{len(diag.gruppen_ohne_zuordnung)} Kontogruppen ohne Zuordnung."
          if diag.gruppen_ohne_zuordnung else
          "Jede Kontogruppe des Exports ist einem HGB-Pfad zugeordnet.",
          sorted(diag.gruppen_ohne_zuordnung))

    r.add("D1", "Fremde Klassifizierung", True, FLAG,
          "Keine vorhanden — der Export trägt nur seine eigene "
          "Systemgliederung. Deshalb kein Benchmark-Tab.")

    r.annahmen.append(
        "Vorzeichen: der Export führt bereits Aktiva positiv und Passiva "
        "negativ. Es wird nichts gedreht.")
    r.annahmen.append(
        "Kostenstellen werden je Konto summiert. Das Mastersheet führt ein "
        "Konto genau einmal; die Kostenstelle wäre ein zweiter Schlüssel.")
    r.annahmen.append(
        f"Das Periodenergebnis stammt aus Konto {ERGEBNISKONTO} "
        "'Current Earnings'. Eine GuV liegt nicht vor.")
    for k, b, grund in diag.ohne_pfad:
        r.offene_befunde.append(f"Konto {k} ({b}): {grund}")
    return r


def _zusatzblaetter(qa, status, verhalten, diag) -> dict:
    blaetter = {
        "QA": (["Prüfung", "Titel", "Bestanden", "Schwere", "Befund", "Details"],
               [[p.id, p.titel, "ja" if p.bestanden else "NEIN", p.schwere,
                 p.befund, " | ".join(p.details)] for p in qa.pruefungen]),
        "Status je Spalte": (
            ["Periode", "Stichtag", "Monate", "Bilanz", "GuV", "Quelle", "Hinweis"],
            [[s.periode, diag.stichtage[s.periode].strftime("%d.%m.%Y"),
              diag.monate.get(s.periode), s.bilanz, s.guv, s.quelle, s.hinweis]
             for s in status.spalten]),
    }
    if verhalten:
        blaetter["Verhaltensprüfung"] = (
            ["Konto", "Bezeichnung", "Klasse", "Kriterium", "wesentlich", "Hinweis"],
            [[v.konto, v.bezeichnung, v.klasse, v.kriterium,
              "ja" if v.wesentlich else "nein", v.hinweis] for v in verhalten])
    return blaetter


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
        Kontrolle("Bilanzidentität (Summe aller Konten = 0)",
                  {p: round(sum(m.saldo(p) for m in mapped), 2) for p in perioden}),
        Kontrolle("Lead NA: Nettovermögen + Eigenkapital = 0",
                  {p: round(netto[p] + eigen[p], 2) for p in perioden}),
        Kontrolle("Konten ohne Position im Lead", fehlbetrag),
        Kontrolle("Equity Roll Forward inkl. Rest = Nettovermögen", rf),
        Kontrolle("Nicht erklärte Eigenkapitalbewegung (Rest)",
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
        "Konten ohne Kontoslot (Wert ohne sichtbares Detail)", ohne,
        toleranz=0.0))
    return kontrollen


def _meta(quellen, ledger, mapped, diag, status, review, hc) -> dict:
    return {
        "Eingabedatei": os.path.basename(quellen.saldenliste),
        "Reader": "myob_susa (vier Jahrestabs)",
        "Entity": ledger.entity,
        "Währung": "AUD",
        "Perioden": ", ".join(
            f"{p} zum {diag.stichtage[p].strftime('%d.%m.%Y')} "
            f"({diag.monate.get(p)} Mon.)" for p in ledger.perioden),
        "Architektur": "Option A — zwei Schichten, Lead-Tabs aus dem Mastersheet",
        "Status je Spalte": status.zusammenfassung(),
        "Strukturquelle": "Systemgliederung des Exports (ClassDescription / "
                          "AccountGroupDesc), übersetzt im Reader",
        "Hausconvention-Version": hc.version,
        "Fingerprint (SHA-256/16)": ledger.fingerprint,
        "Konten gesamt": len(mapped),
        "davon Review": len(review),
        "Konten ohne Pfad": len(diag.ohne_pfad),
        "Warnungen Reader": len(ledger.warnungen),
    }


def _zusammenfassung(meta, ledger, diag, status, ausgabe, lp, befuellt,
                     kontrollen, review) -> None:
    print("=" * 78)
    for k, v in meta.items():
        print(f"  {k:32} {v}")
    print("-" * 78)
    for s in status.spalten:
        print(f"  {s.periode:10} {s.bilanz}   |   GuV: {s.guv}")
        if s.hinweis:
            print(f"             {s.hinweis}")
    print("-" * 78)
    for k in kontrollen:
        print(f"  {'ok ' if k.ok else '!! '}{k.name}")
        print("      " + "  ".join(f"{p}: {v:,.2f}"
                                   for p, v in k.je_periode.items()))
    for z in befuellt.ohne_zeile:
        print(f"  [Befund] Keine Position im Lead: {z.ziel_na} ({z.klasse}), "
              f"{z.konten} Konten — {z.grund}")
    rw = befuellt.roll_forward
    if rw.hat_rest:
        print("-" * 78)
        print("  Woraus der Rest im Eigenkapital besteht (Bewegung Richtung "
              "Nettovermögen):")
        for konto, werte in rw.ergebniskonten.items():
            print(f"    {konto[:46]:46s} "
                  + "  ".join(f"{v:>14,.2f}" for v in werte.values()))
    if befuellt.slotbefunde:
        print("-" * 78)
        letzte = ledger.perioden[-1]
        for b in befuellt.slotbefunde:
            wert = sum(z.werte.get(letzte, 0.0) for z in befuellt.zeilen
                       if z.schluessel in b.ohne_slot)
            print(f"  [Befund] {b.art}: {b.position} ({b.klasse}), Zeile "
                  f"{b.zeile} — {b.konten} Konten, {b.slots} Slots, "
                  f"{len(b.ohne_slot)} ohne Detail ({wert:,.2f} in {letzte})")
    if review:
        print("-" * 78)
        print(f"  Review-Queue: {len(review)} Konten")
        for e in review[:12]:
            print(f"    {e.konto:9s} {e.bezeichnung[:38]:38s} {e.status}")
    print("-" * 78)
    print("\n".join("  " + z for z in lp.als_text()))
    print("-" * 78)
    print(f"  Dealtool geschrieben: {ausgabe}")
    print("=" * 78)
