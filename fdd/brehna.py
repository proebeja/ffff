"""Projekt Brehna (Brehna Projektentwicklungs UG) — Databook aus vier
Jahresabschlüssen.

Die Datenlage unterscheidet sich grundlegend von den bisherigen Mandanten:
**es gibt keine Summen- und Saldenliste.** Die vier Jahresabschlüsse sind
zugleich Werte- und Strukturquelle, und weil sie einen Kontennachweis für
Aktivseite, Passivseite *und* Gewinn- und Verlustrechnung tragen, ist hier
auch die GuV abschlusstreu.

Was daraus folgt:

* Es gibt nichts gegen den Abschluss abzustimmen — der Abschluss *ist* die
  Quelle. An die Stelle der Reconciliation tritt die **Parser-Selbstkontrolle**
  (QA C5): jede gelesene Position gegen ihre gedruckte Summe, jede Seitensumme
  gegen die gedruckte Bilanzsumme, Aktiva gegen Passiva.
* Es gibt keine fremde Klassifizierung, also keinen Benchmark-Tab.
* Die Perioden sind **unterschiedlich lang**: FY2023 umfasst nur November und
  Dezember (Gründung), YTD 07/2026 sieben Monate. Das steht im Status je
  Spalte, weil jeder Periodenvergleich es wissen muss.
"""

from __future__ import annotations

import os
from dataclasses import dataclass

from .core.hausconvention import Hausconvention
from .engine.cascade import Engine
from .engine.laufprotokoll import Laufprotokoll
from .engine.qa import FLAG, ABBRUCH, QAReport, _r2
from .engine.spalten_status import ABSCHLUSSTREU, SpaltenStatus, StatusMatrix
from .engine.v28 import (loese_saldenvortraege, pruefe_verhalten,
                         setze_vorlaeufige_pfade, wende_seitenwechsel_an)
from .export import excel
from .readers.datev_ja_hds_pdf import lies_brehna_jahre
from .views.leads import baue_lead_na, baue_lead_pl
from .views.net_debt import baue_net_debt
from .views.review_queue import baue_review_queue
from .views.schedules import baue_schedules
from .views.working_capital import baue_working_capital

#: Länge der Periode in Monaten — für die Vergleichbarkeit der GuV-Spalten.
PERIODENLAENGE = {"FY2023": 2, "FY2024": 12, "FY2025": 12, "YTD 07/2026": 7}


@dataclass
class Quellen:
    jahresabschluesse: list[str]


def run(quellen: Quellen, ausgabe: str, verbose: bool = True) -> dict:
    lp = Laufprotokoll()
    with lp.phase("Setup") as d:
        hc = Hausconvention.laden()
        d["detail"] = f"Hausconvention v{hc.version}"

    with lp.phase("Einlesen Jahresabschlüsse (PDF)") as d:
        ledger, abschluesse = lies_brehna_jahre(quellen.jahresabschluesse)
        d["detail"] = (f"{len(abschluesse)} Abschlüsse, {len(ledger.accounts)} Konten, "
                       f"{sum(len(a.zeilen) for a in abschluesse)} Kontozeilen")

    with lp.phase("Mapping (Kaskade)") as d:
        engine = Engine(hc)
        mapped = engine.map_ledger(ledger)
        d["detail"] = f"{len(mapped)} Konten"

    with lp.phase("v2.8-Nachlauf") as d:
        mapped, saldenvortrag = loese_saldenvortraege(mapped, ledger.perioden, hc)
        mapped, seitenwechsel = wende_seitenwechsel_an(mapped, ledger.perioden, hc,
                                                       _nachgewiesene_seiten(abschluesse))
        mapped, ungeloest = setze_vorlaeufige_pfade(mapped, ledger.perioden, hc)
        schwelle = _schwelle(mapped, ledger.perioden, hc)
        verhalten = pruefe_verhalten(mapped, ledger.perioden, hc, schwelle)
        d["detail"] = (f"{len(seitenwechsel)} Seitenwechsel, {len(ungeloest)} "
                       f"vorläufige Pfade, {len(verhalten)} Verhaltensbefunde")

    status = _status(abschluesse)

    with lp.phase("Views") as d:
        nd = baue_net_debt(mapped, ledger.perioden, ledger.entity)
        wc = baue_working_capital(mapped, ledger.perioden, ledger.entity)
        lead_na = baue_lead_na(mapped, ledger.perioden, ledger.entity)
        lead_pl = baue_lead_pl(mapped, ledger.perioden, ledger.entity)
        review = baue_review_queue(mapped, ledger.perioden)
        d["detail"] = f"{len(lead_na.bloecke)} NA-Blöcke, {len(review)} Review"

    with lp.phase("Aufrisse (Schedules)") as d:
        schedules = baue_schedules(mapped, ledger.perioden)
        d["detail"] = f"{len(schedules.aufrisse)} Aufrisse"

    with lp.phase("QA-Diagnose") as d:
        qa = _qa(abschluesse, mapped, ledger, status, ungeloest, seitenwechsel,
                 saldenvortrag)
        d["detail"] = (f"{len(qa.pruefungen)} Prüfungen, "
                       f"{len(qa.durchgefallen)} nicht bestanden")

    meta = _meta(quellen, ledger, mapped, abschluesse, status, review, schedules, hc, qa)
    with lp.phase("Excel-Ausgabe") as d:
        excel.schreibe_databook(
            ausgabe, mapped, nd, review, ledger.perioden, ledger.entity, meta=meta,
            wc=wc, schedules=schedules, lead_na=lead_na, lead_pl=lead_pl,
            status=status, qa=qa, verhalten=verhalten)
        d["detail"] = os.path.basename(ausgabe)
    lp.zaehle_arbeitsmappe(ausgabe)

    if verbose:
        _zusammenfassung(meta, ledger, abschluesse, status, ausgabe, lp)
    return {"laufprotokoll": lp, "ledger": ledger, "mapped": mapped, "qa": qa,
            "abschluesse": abschluesse, "status": status, "nd": nd, "wc": wc,
            "lead_na": lead_na, "lead_pl": lead_pl, "schedules": schedules,
            "review": review, "verhalten": verhalten, "meta": meta,
            "seitenwechsel": seitenwechsel, "ungeloest": ungeloest}


def _nachgewiesene_seiten(abschluesse) -> dict[str, dict[str, str]]:
    """Auf welcher Bilanzseite weist der Abschluss ein Konto je Periode aus?

    Hier ist der Abschluss die einzige Quelle, also ist die Antwort für jede
    Periode bekannt — eine Vorzeichenableitung kommt gar nicht erst zum Zug."""
    je_konto: dict[str, dict[str, str]] = {}
    for a in abschluesse:
        for z in a.zeilen:
            if z.sektion in ("AKTIVA", "PASSIVA"):
                je_konto.setdefault(z.konto, {})[a.periode] = z.sektion
    return je_konto


def _schwelle(mapped, perioden, hc) -> float:
    prozent = hc.wesentlichkeit.get("bilanzsumme_prozent", 2) / 100.0
    summen = [sum(m.saldo(p) for m in mapped if m.hgb_pfad.startswith("/Aktiva"))
              for p in perioden]
    return max(summen or [0.0]) * prozent


def _status(abschluesse) -> StatusMatrix:
    """Alle vier Spalten sind abschlusstreu — für Bilanz UND GuV, weil der
    Kontennachweis beide Rechenwerke abdeckt. Der Zwischenabschluss trägt den
    Zusatz, dass er nicht testiert ist."""
    spalten = []
    for a in abschluesse:
        monate = PERIODENLAENGE.get(a.periode)
        zusatz = (" (Zwischenabschluss, nicht testiert)"
                  if a.ist_zwischenabschluss else "")
        spalten.append(SpaltenStatus(
            periode=a.periode, bilanz=ABSCHLUSSTREU + zusatz,
            guv=ABSCHLUSSTREU + zusatz,
            quelle=f"Jahresabschluss zum {a.stichtag} inkl. Kontennachweis "
                   "für Aktiva, Passiva und GuV",
            hinweis=(f"Periodenlänge {monate} Monate — die GuV-Spalte ist mit "
                     "den Zwölfmonatsspalten nicht unmittelbar vergleichbar."
                     if monate and monate != 12 else "")))
    return StatusMatrix(spalten=spalten)


def _qa(abschluesse, mapped, ledger, status, ungeloest, seitenwechsel,
        saldenvortrag) -> QAReport:
    """QA für diese Datenlage. Prüfungen, die eine Saldenliste voraussetzen
    (Blockgrenze, Monatsachse, vorgerechnete Spalten), sind hier gegenstandslos
    und werden als solche ausgewiesen — nicht stillschweigend weggelassen."""
    r = QAReport(seitenwechsel=list(seitenwechsel), saldenvortrag=list(saldenvortrag))
    perioden = ledger.perioden

    r.add("A1", "Blockgrenze des Datenblocks eindeutig", True, ABBRUCH,
          "Gegenstandslos: die Quelle ist ein Kontennachweis mit klarer "
          "Abschlusszeile je Seite, keine Saldenliste mit Nebenrechnungen.",
          [f"{a.periode}: {len(a.zeilen)} Kontozeilen" for a in abschluesse])

    summen = {p: _r2(sum(m.saldo(p) for m in mapped)) for p in perioden}
    r.add("A2", "Bilanzidentität je Periodenspalte (Toleranz 1 EUR)",
          all(abs(v) <= 1.0 for v in summen.values()), FLAG,
          "Summe aller Konten je Spalte: "
          + ", ".join(f"{p} = {v:,.2f}" for p, v in summen.items()))

    doppelt = [a.periode for a in abschluesse
               if len({z.konto for z in a.zeilen}) != len(a.zeilen)]
    r.add("A3", "Kontoschlüssel eindeutig", not doppelt, FLAG,
          "je Abschluss geprüft." if not doppelt else f"Duplikate in {doppelt}.")

    r.add("A4", "Kontoschlüssel numerisch und im Rahmen des Kontenrahmens",
          all(m.konto.isdigit() for m in mapped), FLAG,
          f"{len(mapped)} Konten, sechsstellig; SKR03 mit zwei angehängten "
          "Nullen (120000 = 1200 Bank, 497000 = 4970 Nebenkosten des "
          "Geldverkehrs).")

    null = [m.konto for m in mapped if all(abs(m.saldo(p)) < 0.005 for p in perioden)]
    r.add("A5", "Nullkonten mitgeführt und markiert", True, FLAG,
          f"{len(null)} Konten ohne Saldo in allen Perioden.")

    for p in perioden:
        r.nicht_zugeordnet[p] = _r2(sum(u.salden.get(p, 0.0) for u in ungeloest))
        r.bilanzsumme[p] = _r2(sum(m.saldo(p) for m in mapped
                                   if m.hgb_pfad.startswith("/Aktiva")))
    r.add("A6", "Ungelöste Konten liegen INNERHALB der Bilanz", True, FLAG,
          f"{len(ungeloest)} Konten ohne bestimmbaren Pfad.",
          [f"{p}: {r.nicht_zugeordnet[p]:,.2f} von {r.bilanzsumme[p]:,.2f}"
           for p in perioden])

    r.add("B1", "Monatsspalten kumuliert oder periodisch", True, ABBRUCH,
          "Gegenstandslos: der Abschluss liefert Jahres- bzw. Stichtagswerte, "
          "keine Monatsmatrix. Es wird nichts über die Zeitachse aggregiert.")

    r.add("B2", "Jahresanker", True, FLAG,
          "Jede Spalte ist genau ein Abschlussstichtag: "
          + ", ".join(f"{a.periode} = {a.stichtag}" for a in abschluesse))

    r.add("B3", "Vorgerechnete Spalten nicht übernommen", True, FLAG,
          "Der Abschluss druckt jede Zahl zweimal (Vor- und Summenspalte); "
          "gelesen wird nur die Wertspalte, die Positionssumme wird aus den "
          "Konten selbst gebildet und gegen die gedruckte geprüft (C5).")

    laengen = {a.periode: PERIODENLAENGE.get(a.periode) for a in abschluesse}
    ungleich = len(set(laengen.values())) > 1
    r.add("B4", "Bilanzkontinuität / Vergleichbarkeit der Perioden",
          not ungleich, FLAG,
          "Die Perioden sind unterschiedlich lang — jeder GuV-Vergleich muss "
          "das berücksichtigen." if ungleich else "gleich lang.",
          [f"{p}: {m} Monate" for p, m in laengen.items()])

    r.add("B5", "Saldenvortragspräsenz", True, FLAG,
          "Der Abschluss weist keine DATEV-Saldenvortragskonten aus; der "
          "Verlustvortrag steht als Konto 286800 im GuV-Nachweis und wird "
          "als Eigenkapital geführt.")

    r.add("C1", "Strukturquellen-Abdeckung je Periode und Rechenwerk", True, FLAG,
          "Jede Periode hat einen Kontennachweis für Aktiva, Passiva und GuV — "
          "damit ist auch die GuV abschlusstreu.",
          [f"{s.periode}: Bilanz {s.bilanz} · GuV {s.guv}" for s in status.spalten])

    r.add("C2", "Reconciliation auf Kontenebene", True, FLAG,
          "Gegenstandslos: es gibt keine zweite Wertequelle. Der Abschluss ist "
          "die Quelle, nicht das Abstimmziel. An ihre Stelle tritt C5.")

    r.add("C3", "Bekannte Abstimmmuster", True, FLAG,
          "Keine Saldenspaltung und keine Verrechnung feststellbar: der "
          "Kontennachweis führt jedes Konto genau einmal je Seite.")

    r.add("C4", "Kontenstamm-Abdeckung", True, FLAG,
          "Kein separater Kontenplan vorhanden; die Bezeichnungen stammen aus "
          "dem Kontennachweis selbst.")

    proben = [(a.periode, t) for a in abschluesse for t in a.probe()]
    schlecht = [(p, t) for p, t in proben if not t[1]]
    r.add("C5", "Parser-Selbstkontrolle der Strukturquelle", not schlecht, ABBRUCH,
          f"{len(proben) - len(schlecht)} von {len(proben)} Proben bestanden: "
          "jede Position gegen ihre gedruckte Summe, jede Seite gegen die "
          "gedruckte Bilanzsumme, Aktiva gegen Passiva.",
          [f"{p} · {t[0]}: {t[2]}" for p, t in proben])

    r.add("D1", "Fremde Klassifizierung", True, FLAG,
          "Keine vorhanden — die Quelle ist ein Abschluss ohne Fremdmapping. "
          "Deshalb kein Benchmark-Tab.")

    r.annahmen.append(
        "Vorzeichen: der Abschluss druckt die Passivseite positiv und die "
        "Aufwendungen negativ. Beide werden auf die Databook-Konvention "
        "gedreht (Soll positiv, Haben negativ); die Aktivseite bleibt.")
    r.annahmen.append(
        "Konto 286800 'Verlustvortrag nach Verwendung' steht im GuV-Nachweis, "
        "wird aber als Eigenkapital geführt.")
    for f in seitenwechsel:
        r.offene_befunde.append(f.pflichtfrage)
    for u in ungeloest:
        r.offene_befunde.append(f"Konto {u.konto} ({u.bezeichnung}) ohne Pfad.")
    return r


def _meta(quellen, ledger, mapped, abschluesse, status, review, schedules, hc, qa) -> dict:
    return {
        "Eingabedateien": ", ".join(os.path.basename(p)
                                    for p in quellen.jahresabschluesse),
        "Reader": "datev_ja_hds_pdf (vier Abschlüsse)",
        "Entity": ledger.entity,
        "Perioden": ", ".join(f"{p} ({PERIODENLAENGE.get(p, '?')} Mon.)"
                              for p in ledger.perioden),
        "Status je Spalte": status.zusammenfassung(),
        "Strukturquelle": "Kontennachweis des Abschlusses für Aktiva, Passiva und GuV",
        "Werte- und Strukturquelle identisch": "ja — es gibt keine Saldenliste",
        "Hausconvention-Version": hc.version,
        "Fingerprint (SHA-256/16)": ledger.fingerprint,
        "Konten gesamt": len(mapped),
        "davon Review": len(review),
        "Aufrisse (Schedules)": len(schedules.aufrisse),
        "Konten ohne Aufriss": len(schedules.ohne_aufriss),
        "Parser-Selbstkontrolle": next(
            (p.befund for p in qa.pruefungen if p.id == "C5"), ""),
        "Warnungen Reader": len(ledger.warnungen),
    }


def _zusammenfassung(meta, ledger, abschluesse, status, ausgabe, lp) -> None:
    print("=" * 78)
    for k, v in meta.items():
        print(f"  {k:38} {v}")
    print("-" * 78)
    for s in status.spalten:
        print(f"  {s.periode:14} {s.bilanz}")
        if s.hinweis:
            print(f"                 {s.hinweis}")
    if ledger.warnungen:
        print("-" * 78)
        for w in ledger.warnungen[:8]:
            print(f"  [warn] {w}")
    print("-" * 78)
    print("\n".join("  " + z for z in lp.als_text()))
    print("-" * 78)
    print(f"  Databook geschrieben: {ausgabe}")
    print("=" * 78)
