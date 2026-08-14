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
from .engine.spalten_status import (ABGELEITET, ABSCHLUSSTREU, AGGREGIERT,
                                    VORLAEUFIG, baue_status)
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


def run(quellen: Quellen, ausgabe: str, verbose: bool = True) -> dict:
    hc = Hausconvention.laden()
    ledger, diagnose = SusaDatabookReader().lesen_mit_diagnose(quellen.susa)
    susa_konten = {a.konto for a in ledger.accounts}

    # --- Strukturquellen, in der Reihenfolge ihrer Belastbarkeit -----------
    ja = kn = None
    if quellen.jahresabschluss:
        ja = lies_datev_ja(quellen.jahresabschluss)
        kn = ja.als_kontennachweis(PERIODE_JA, PERIODE_VJ)
        ledger = wende_kontennachweis_an(ledger, kn)

    plan = None
    if quellen.kontenplan:
        plan = lies_kontenplan(quellen.kontenplan)
        ledger = wende_kontenplan_an(ledger, plan)
        ledger.warnungen.extend(_hinweise_zu_duplikaten(diagnose, plan))

    protokoll_ledger = ledger
    mapped = Engine(hc).map_ledger(ledger)

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
    nd = baue_net_debt(mapped, ledger.perioden, ledger.entity)
    wc = baue_working_capital(mapped, ledger.perioden, ledger.entity)
    lead_na = baue_lead_na(mapped, ledger.perioden, ledger.entity)
    lead_pl = baue_lead_pl(mapped, ledger.perioden, ledger.entity)
    schedules = baue_schedules(mapped, ledger.perioden)
    review = baue_review_queue(mapped, ledger.perioden)
    benchmark = baue_benchmark(mapped, diagnose.benchmark, ledger.entity)

    # --- Abstimmungen ------------------------------------------------------
    recon_kn = reconcile_gegen_kontennachweis(mapped, ja, PERIODE_JA) if ja else None
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

    meta = _meta(quellen, ledger, mapped, diagnose, status, plan, kn,
                 review, schedules, benchmark, hc)
    excel.schreibe_databook(
        ausgabe, mapped, nd, review, ledger.perioden, ledger.entity, meta=meta,
        wc=wc, schedules=schedules, lead_na=lead_na, lead_pl=lead_pl,
        status=status, benchmark=benchmark, recon_abschluss=recon_kn,
        recon_aggregiert=recon_agg)

    if verbose:
        _zusammenfassung(meta, ledger, diagnose, status, ausgabe)
    return {"ledger": protokoll_ledger, "mapped": mapped, "diagnose": diagnose,
            "kn": kn, "ja": ja, "plan": plan, "status": status, "nd": nd,
            "wc": wc, "lead_na": lead_na, "lead_pl": lead_pl,
            "schedules": schedules, "review": review, "benchmark": benchmark,
            "recon_kn": recon_kn, "recon_agg": recon_agg, "meta": meta}


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
    3. Konto 701 0 (Cash-Pool) steht 2022 auf der Passivseite und 2023 auf der
       Aktivseite; das Mastersheet führt einen Pfad je Konto und folgt dem
       jüngeren Abschluss."""
    if ja is None:
        return {}
    ergebnis = sum(m.saldo(PERIODE_VJ) for m in mapped
                   if m.hgb_pfad.startswith("/GuV"))
    fehlbetrag = 2_902_421.91

    # Ein Seitenwechsel liegt vor, wenn der **Nettosaldo** die Seite wechselt.
    # Dass ein Konto in einem Jahr gespalten ist und im anderen nicht, ist
    # dagegen bloße Darstellung und kein Wechsel.
    seitenwechsel = 0.0
    je_konto: dict[str, list] = {}
    for e in ja.eintraege:
        je_konto.setdefault(e.konto, []).append(e)
    for konto, eintraege in je_konto.items():
        # Nur Konten, die der Abschluss überhaupt auf beiden Seiten führt,
        # können die Seite wechseln. Ein Konto mit nur einer Seite ändert
        # allenfalls sein Saldovorzeichen (1789 0 Umsatzsteuer) und bleibt
        # dabei in derselben Position — das ist kein Wechsel.
        if len({e.sektion for e in eintraege}) < 2:
            continue
        netto_gj = sum(e.vorzeichenrichtig() for e in eintraege)
        netto_vj = sum(e.vorzeichenrichtig(vorjahr=True) for e in eintraege)
        if abs(netto_gj) < 0.005 or abs(netto_vj) < 0.005:
            continue
        if (netto_gj >= 0) != (netto_vj >= 0):
            seitenwechsel += netto_vj

    return {
        "A. Eigenkapital": [
            Ueberleitung("Jahresergebnis 2022, in der SuSa noch in den "
                         "GuV-Konten statt im Bilanzverlust", -ergebnis),
            Ueberleitung("Nicht durch Eigenkapital gedeckter Fehlbetrag, vom "
                         "Bericht auf die Aktivseite umgestellt (§ 268 III HGB)",
                         fehlbetrag),
        ],
        "II. Forderungen und sonstige Vermögensgegenstände": [
            Ueberleitung("Konto 701 0 Cash-Pool: 2022 Verbindlichkeit, 2023 "
                         "Forderung — das Mastersheet folgt dem jüngeren "
                         "Abschluss", seitenwechsel),
            Ueberleitung("Saldenspaltung Konto 1600 0 im Vorjahr", -895.99),
        ],
        "C. Verbindlichkeiten": [
            Ueberleitung("Gegenposten zum Seitenwechsel Konto 701 0",
                         -seitenwechsel),
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
