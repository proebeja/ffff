"""Projekt Kitchenstories (AJNS New Media GmbH) — Aufbau des Databooks.

Die Kette weicht in einem Punkt von den übrigen Mandanten ab: es gibt
**mehrere Quellen mit unterschiedlicher Reichweite**, und jede darf nur das
tun, wofür sie belastbar ist.

* SuSa (Excel)         — Werteseite, alle vier Perioden.
* Kontennachweis 2023  — Strukturquelle, aber nur für die **Bilanz** und nur
                         für FY2023 und dessen Vorjahresspalte FY2022.
* Kontenplan 2025      — keine HGB-Zuordnung; liefert Bezeichnungen und die
                         DATEV-Funktionscodes und schließt damit nur Lücken,
                         die der Kontennachweis lässt.
* Prüfbericht 2022     — Abstimmziel auf Gliederungsebene, ohne Kontenebene.

Daraus folgt der spaltenweise Status: FY2022/FY2023 abschlusstreu in der
Bilanz, FY2024 und YTD Jul25 vorläufig, die GuV durchgehend abgeleitet.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Optional

from .core.hausconvention import Hausconvention
from .engine.cascade import Engine
from .engine.recon_abschluss import (Ueberleitung, reconcile_aggregiert,
                                     reconcile_gegen_kontennachweis)
from .engine.einfrierung import (lade_snapshot, schreibe_snapshot,
                                 wende_einfrierung_an)
from .engine.laufprotokoll import Laufprotokoll
from .engine.qa import baue_qa_report
from .engine.spalten_status import baue_status
from .engine.v28 import (loese_saldenvortraege, pruefe_verhalten,
                         setze_vorlaeufige_pfade, wende_seitenwechsel_an)
from .engine.kontennachweis_apply import wende_kontennachweis_an
from .export import excel
from .readers.datev_ja_pdf import lies_datev_ja
from .readers.datev_kontenplan_pdf import lies_kontenplan, wende_kontenplan_an
from .readers.susa_databook import SusaDatabookReader
from .views.benchmark import baue_benchmark
from .views.leads import baue_lead_na, baue_lead_pl
from .views.net_debt import baue_net_debt
from .views.review_queue import baue_review_queue
from .views.schedules import baue_schedules
from .views.working_capital import baue_working_capital

#: Bilanzsumme laut EY-Prüfbericht 2022, in Databook-Vorzeichen
#: (Aktiva positiv, Passiva negativ). Aus Anlage "BILANZ zum 31. Dezember 2022".
PRUEFBERICHT_2022: dict[str, float] = {
    "I. Immaterielle Vermögensgegenstände": 394_131.00,
    "II. Sachanlagen": 375_689.00,
    "III. Finanzanlagen": 48_138.76,
    "I. Vorräte": 86_611.40,
    "II. Forderungen und sonstige Vermögensgegenstände": 491_012.03,
    "III. Kassenbestand und Guthaben bei Kreditinstituten": 43_840.64,
    "C. Rechnungsabgrenzungsposten": 107_487.63,
    "A. Eigenkapital": 0.00,
    "B. Rückstellungen": -705_354.12,
    "C. Verbindlichkeiten": -3_743_978.25,
}

PERIODE_JA = "FY2023"
PERIODE_VJ = "FY2022"


@dataclass
class Quellen:
    susa: str
    jahresabschluss: Optional[str] = None
    kontenplan: Optional[str] = None
    pruefbericht: Optional[str] = None


def run(quellen: Quellen, ausgabe: str, verbose: bool = True,
        snapshot: Optional[str] = None) -> dict:
    """``snapshot`` ist das Entscheidungsprotokoll eines früheren Laufs. Ist es
    gesetzt, werden nur die Konten neu hergeleitet, die eine der benannten
    Änderungen berührt — alle übrigen behalten ihr eingefrorenes Ergebnis."""
    lp = Laufprotokoll()
    with lp.phase("Setup") as d:
        hc = Hausconvention.laden()
        d["detail"] = f"Hausconvention v{hc.version}"

    with lp.phase("Einlesen SuSa") as d:
        ledger, diagnose = SusaDatabookReader().lesen_mit_diagnose(quellen.susa)
        d["detail"] = f"{diagnose.kontozeilen} Kontozeilen, {len(ledger.accounts)} Konten"

    ja = kn = None
    with lp.phase("Einlesen Jahresabschluss (PDF)") as d:
        if quellen.jahresabschluss:
            ja = lies_datev_ja(quellen.jahresabschluss)
            kn = ja.als_kontennachweis(PERIODE_JA, PERIODE_VJ)
            ledger = wende_kontennachweis_an(ledger, kn)
            d["detail"] = f"{len(ja.eintraege)} Kontennachweis-Einträge, 9 Seiten"

    plan = None
    with lp.phase("Einlesen Kontenplan (PDF)") as d:
        if quellen.kontenplan:
            plan = lies_kontenplan(quellen.kontenplan)
            ledger = wende_kontenplan_an(ledger, plan)
            ledger.warnungen.extend(_hinweise_zu_duplikaten(diagnose, plan))
            d["detail"] = f"{len(plan.eintraege)} Einträge, 4 Seiten"

    protokoll_ledger = ledger
    with lp.phase("Mapping (Kaskade)") as d:
        engine = Engine(hc)
        mapped = engine.map_ledger(ledger)
        d["detail"] = f"{len(mapped)} Konten durch 4 Kaskadenstufen"

    with lp.phase("v2.8-Nachlauf") as d:
        mapped, saldenvortrag = loese_saldenvortraege(mapped, ledger.perioden, hc)
        mapped, seitenwechsel = wende_seitenwechsel_an(
            mapped, ledger.perioden, hc, _nachgewiesene_seiten(ja))
        mapped, ungeloest = setze_vorlaeufige_pfade(mapped, ledger.perioden, hc)
        schwelle = _wesentlichkeitsschwelle(mapped, ledger.perioden, hc)
        verhalten = pruefe_verhalten(mapped, ledger.perioden, hc, schwelle)
        d["detail"] = (f"{len(seitenwechsel)} Seitenwechsel, "
                       f"{len(saldenvortrag)} Saldenvorträge, "
                       f"{len(ungeloest)} vorläufige Pfade, "
                       f"{len(verhalten)} Verhaltensbefunde")

    einfrierung = None
    with lp.phase("Einfrierung / Delta") as d:
        if snapshot and os.path.exists(snapshot):
            einfrierung = wende_einfrierung_an(mapped, lade_snapshot(snapshot), hc)
            mapped = einfrierung.mapped
            d["detail"] = (f"{einfrierung.eingefroren} eingefroren, "
                           f"{einfrierung.neu_hergeleitet} neu hergeleitet, "
                           f"{len(einfrierung.defekte)} Defekte")
        else:
            d["detail"] = "kein Snapshot — vollständige Herleitung"

    # --- Spaltenweiser Status ---------------------------------------------
    quellen_je_periode = {
        PERIODE_VJ: "Kontennachweis 2023 (Vorjahresspalte) + Prüfbericht 2022",
        PERIODE_JA: f"Jahresabschluss {PERIODE_JA} inkl. Kontennachweis",
    }
    status = baue_status(
        perioden=ledger.perioden,
        kontennachweis_perioden={PERIODE_JA, PERIODE_VJ} if kn else set(),
        aggregiert_perioden=set(),
        quellen=quellen_je_periode)

    # --- Views -------------------------------------------------------------
    with lp.phase("Views (Leads, Net Debt, WC, Benchmark)") as d:
        nd = baue_net_debt(mapped, ledger.perioden, ledger.entity)
        wc = baue_working_capital(mapped, ledger.perioden, ledger.entity)
        lead_na = baue_lead_na(mapped, ledger.perioden, ledger.entity)
        lead_pl = baue_lead_pl(mapped, ledger.perioden, ledger.entity)
        review = baue_review_queue(mapped, ledger.perioden)
        benchmark = baue_benchmark(mapped, diagnose.benchmark, ledger.entity)
        d["detail"] = f"{len(lead_na.bloecke)} NA-Blöcke, {len(review)} Review"

    with lp.phase("Aufrisse (Schedules)") as d:
        schedules = baue_schedules(mapped, ledger.perioden)
        d["detail"] = f"{len(schedules.aufrisse)} Aufrisse"

    # --- Abstimmungen ------------------------------------------------------
    with lp.phase("Abstimmungen") as d:
        recon_kn = reconcile_gegen_kontennachweis(mapped, ja, PERIODE_JA) if ja else None
        d["detail"] = ("FY2023 Kontennachweis + FY2022 Prüfbericht"
                       if recon_kn else "keine")
    recon_agg = None
    if quellen.pruefbericht:
        recon_agg = reconcile_aggregiert(
            mapped, PERIODE_VJ, PRUEFBERICHT_2022,
            quelle=os.path.basename(quellen.pruefbericht),
            hinweise=["Der Prüfbericht weist keine Kontenebene aus; abgestimmt "
                      "wird deshalb nur bis zur Gliederungsebene.",
                      "Die Überleitungsspalte trägt die Posten, die Databook "
                      "und Bericht systematisch trennen — sie sind erklärt, "
                      "nicht bereinigt."],
            ueberleitung=_ueberleitung_2022(mapped, ja))

    with lp.phase("QA-Diagnose") as d:
        qa = baue_qa_report(
            diagnose=diagnose, mapped=mapped, perioden=ledger.perioden, ja=ja,
            plan=plan, recon_kn=recon_kn, benchmark=benchmark, status=status,
            ungeloest=ungeloest, seitenwechsel=seitenwechsel,
            saldenvortrag=saldenvortrag, ja_pdf_pfad=quellen.jahresabschluss)
        d["detail"] = (f"{len(qa.pruefungen)} Einzelprüfungen, "
                       f"{len(qa.durchgefallen)} nicht bestanden")

    meta = _meta(quellen, ledger, mapped, diagnose, status, plan, kn,
                 review, schedules, benchmark, hc)
    with lp.phase("Excel-Ausgabe") as d:
        excel.schreibe_databook(
            ausgabe, mapped, nd, review, ledger.perioden, ledger.entity, meta=meta,
            wc=wc, schedules=schedules, lead_na=lead_na, lead_pl=lead_pl,
            status=status, benchmark=benchmark, recon_abschluss=recon_kn,
            recon_aggregiert=recon_agg, qa=qa, verhalten=verhalten,
            einfrierung=einfrierung)
        d["detail"] = os.path.basename(ausgabe)
    lp.zaehle_arbeitsmappe(ausgabe)
    lp.ki_aufrufe.extend(getattr(engine, "ki_aufrufe", []))

    if verbose:
        _zusammenfassung(meta, ledger, diagnose, status, ausgabe)
    return {"laufprotokoll": lp, "qa": qa, "einfrierung": einfrierung,
            "seitenwechsel": seitenwechsel, "saldenvortrag": saldenvortrag,
            "ungeloest": ungeloest, "verhalten": verhalten,
            "ledger": protokoll_ledger, "mapped": mapped, "diagnose": diagnose,
            "kn": kn, "ja": ja, "plan": plan, "status": status, "nd": nd,
            "wc": wc, "lead_na": lead_na, "lead_pl": lead_pl,
            "schedules": schedules, "review": review, "benchmark": benchmark,
            "recon_kn": recon_kn, "recon_agg": recon_agg, "meta": meta}


def _nachgewiesene_seiten(ja) -> dict[str, dict[str, str]]:
    """Welche Bilanzseite weist der Kontennachweis je Konto und Periode aus?

    Nur für die beiden Perioden, die er abdeckt. Für FY2024 und YTD Jul25
    liegt kein Abschluss vor — dort bleibt die Vorzeichenableitung."""
    if ja is None:
        return {}
    je_konto: dict[str, dict[str, str]] = {}
    eintraege: dict[str, list] = {}
    for e in ja.eintraege:
        eintraege.setdefault(e.konto, []).append(e)
    for konto, es in eintraege.items():
        netto_gj = sum(e.vorzeichenrichtig() for e in es)
        netto_vj = sum(e.vorzeichenrichtig(vorjahr=True) for e in es)
        seiten = {}
        for periode, netto, feld in ((PERIODE_JA, netto_gj, "gj"),
                                     (PERIODE_VJ, netto_vj, "vj")):
            relevant = [e for e in es if abs(getattr(e, feld)) > 0.005]
            if not relevant:
                continue
            # Bei Saldenspaltung entscheidet die Seite des Nettosaldos.
            seiten[periode] = ("AKTIVA" if netto >= 0 else "PASSIVA") \
                if len({e.sektion for e in relevant}) > 1 else relevant[0].sektion
        if seiten:
            je_konto[konto] = seiten
    return je_konto


def _wesentlichkeitsschwelle(mapped, perioden, hc) -> float:
    """Schwelle für die Verhaltensprüfung: Anteil der Bilanzsumme laut
    Hausconvention, gemessen an der größten Periode."""
    prozent = hc.wesentlichkeit.get("bilanzsumme_prozent", 2) / 100.0
    summen = [sum(m.saldo(p) for m in mapped if m.hgb_pfad.startswith("/Aktiva"))
              for p in perioden]
    return max(summen or [0.0]) * prozent


def _hinweise_zu_duplikaten(diagnose, plan) -> list[str]:
    """Kann der Kontenplan ein strittiges Duplikat auflösen?

    Trägt eine der beiden Zeilen die Bezeichnung eines **anderen** Kontos des
    Kontenplans, ist die Kontonummer in der SuSa schlicht falsch — und der
    Fall damit klärbar, statt nur offen."""
    hinweise = []
    nach_bezeichnung = {}
    for konto, e in plan.eintraege.items():
        if e.bezeichnung:
            nach_bezeichnung.setdefault(e.bezeichnung.strip().lower(), konto)
    for d in diagnose.duplikate:
        if d.konsolidierbar:
            continue
        for bez in d.bezeichnungen:
            treffer = nach_bezeichnung.get(bez.strip().lower())
            if treffer and treffer != d.konto:
                hinweise.append(
                    f"Konto {d.konto}: die Zeile '{bez[:40]}' trägt im "
                    f"Kontenplan die Nummer {treffer}. Die SuSa führt sie unter "
                    f"{d.konto} — vermutlich eine falsche Kontonummer, nicht "
                    "zwei Buchungen auf demselben Konto.")
    return hinweise


def _ueberleitung_2022(mapped, ja) -> dict[str, list[Ueberleitung]]:
    """Die drei systematischen Unterschiede zwischen SuSa-Stand 31.12.2022 und
    dem Prüfbericht.

    1. Der Prüfbericht zeigt den Abschluss **nach** Ergebnisverwendung: das
       Jahresergebnis steckt im Bilanzverlust. Die SuSa führt es noch in den
       GuV-Konten, das Eigenkapital trägt nur den Vortrag.
    2. Ein negatives Eigenkapital wird nach § 268 Abs. 3 HGB als "nicht durch
       Eigenkapital gedeckter Fehlbetrag" auf die **Aktivseite** umgestellt.
       Das Databook kennt diese Umstellung nicht — sie ist reine Darstellung.
    Der frühere dritte Posten — Konto 701 0 auf der falschen Bilanzseite — ist
    mit der Seitenwechsel-Regel (v2.8) entfallen: die Bilanzseite folgt jetzt je
    Periode dem Vorzeichen, damit steht das Konto FY2022 dort, wo es der
    Abschluss auch ausweist. Was bleibt, ist die Saldenspaltung von 1600 0."""
    if ja is None:
        return {}
    ergebnis = sum(m.saldo(PERIODE_VJ) for m in mapped
                   if m.hgb_pfad.startswith("/GuV"))
    fehlbetrag = 2_902_421.91

    return {
        "A. Eigenkapital": [
            Ueberleitung("Jahresergebnis 2022, in der SuSa noch in den "
                         "GuV-Konten statt im Bilanzverlust", -ergebnis),
            Ueberleitung("Nicht durch Eigenkapital gedeckter Fehlbetrag, vom "
                         "Bericht auf die Aktivseite umgestellt (§ 268 III HGB)",
                         fehlbetrag),
        ],
        "II. Forderungen und sonstige Vermögensgegenstände": [
            Ueberleitung("Saldenspaltung Konto 1600 0 im Vorjahr", -895.99),
        ],
        "C. Verbindlichkeiten": [
            Ueberleitung("Gegenposten zur Saldenspaltung Konto 1600 0", 895.99),
        ],
    }


def _meta(quellen, ledger, mapped, diagnose, status, plan, kn, review,
          schedules, benchmark, hc) -> dict:
    ab = plan.abdeckung({a.konto for a in ledger.accounts}) if plan else (0, 0)
    strittig = [d for d in diagnose.duplikate if not d.konsolidierbar]
    return {
        "Eingabedatei": os.path.basename(quellen.susa),
        "Reader": "susa_databook",
        "Entity": ledger.entity,
        "Perioden": ", ".join(ledger.perioden),
        "Status je Spalte": status.zusammenfassung(),
        "Strukturquelle Bilanz": (f"Kontennachweis {PERIODE_JA} "
                                  f"({len(kn.konten)} Konten)" if kn else "keine"),
        "Strukturquelle GuV": "keine — SKR03-Default + Hausconvention (nicht abschlusstreu)",
        "Kontenplan": (f"{ab[0]} von {ab[1]} Konten abgedeckt, "
                       f"{len(plan.mit_funktion())} mit Funktionscode" if plan else "nicht vorhanden"),
        "Hausconvention-Version": hc.version,
        "Fingerprint (SHA-256/16)": ledger.fingerprint,
        "Kontozeilen gelesen": diagnose.kontozeilen,
        "Konten nach Konsolidierung": len(ledger.accounts),
        "Trial Balance ausgeglichen": "ja" if diagnose.ist_ausgeglichen else "NEIN",
        "Doppelte Kontonummern": (f"{len(diagnose.duplikate)}, davon "
                                  f"{len(strittig)} strittig -> Review"),
        "davon Review": len(review),
        "davon technisch (TECH)": sum(1 for m in mapped if m.klasse.value == "TECH"),
        "Aufrisse (Schedules)": len(schedules.aufrisse),
        "Konten ohne Aufriss": len(schedules.ohne_aufriss),
        "Benchmark gegen manuelle Klassifizierung": (
            f"{len(benchmark.zeilen)} Konten, {len(benchmark.abweichungen)} Abweichungen, "
            f"{len(benchmark.unuebersetzbar)} nicht übersetzbar"),
        "Warnungen Reader": len(ledger.warnungen),
    }


def _zusammenfassung(meta, ledger, diagnose, status, ausgabe) -> None:
    print("=" * 74)
    for k, v in meta.items():
        print(f"  {k:42} {v}")
    print("-" * 74)
    print("  Spaltensummen der Trial Balance (muessen 0 sein):")
    for p, v in diagnose.spaltensummen.items():
        print(f"    {p:12} {v:>22.10f}")
    print("-" * 74)
    for s in status.spalten:
        print(f"  {s.periode:12} Bilanz: {s.bilanz:32} GuV: {s.guv}")
    if ledger.warnungen:
        print("-" * 74)
        for w in ledger.warnungen[:10]:
            print(f"  [warn] {w}")
        if len(ledger.warnungen) > 10:
            print(f"  … und {len(ledger.warnungen) - 10} weitere")
    print("-" * 74)
    print(f"  Databook geschrieben: {ausgabe}")
    print("=" * 74)
