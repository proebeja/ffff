"""Projekt Luma (BT Imaging, Australien) — Aufbau des Databooks.

Quellen und ihre Reichweite:

* **Bilanz, quartalsweise.** Vier Jahresblätter und fünfzehn Quartalsblätter,
  FY2023 bis FY2026. Der Bestand steht in der Spalte ``Year``.
* **GuV, jährlich und quartalsweise.** Dieselben Perioden. Hier ist ``Year``
  kumuliert; gelesen wird ``PeriodMovementTYD``.

Kein Kontennachweis und kein Abschluss. Alle vier Jahre sind damit
**vorläufig und nicht abschlusstreu** — die Positionen stammen aus dem
Kontenrahmen AASB, nicht aus einem testierten Abschluss.

Zwei Dinge, die dieses Mandat von den bisherigen unterscheidet:

**Das Mastersheet führt Quartale, alle übrigen Blätter Jahre.** Der Aufriss
einer Position über fünfzehn Quartalsspalten ist nicht lesbar, und die
Net-Asset-Sicht einer FDD ist eine Sicht auf Stichtage. Die Quartale bleiben
deshalb im Mastersheet, eingeklappt neben ihrer Jahresspalte, und die
Jahresspalte rechnet sich aus ihnen: Bilanzpositionen als Schlussbestand des
letzten Quartals, GuV-Positionen als Summe der vier.

**FY2023 ist ein Rumpfjahr.** Der Export beginnt im Juli 2022, das
Geschäftsjahr im April. FY2023 trägt drei statt vier Quartale; die Bilanz zum
31.03.2023 ist davon unberührt, die GuV FY2023 deckt aber nur neun Monate ab
und ist mit den übrigen Jahren nicht vergleichbar. Das steht als Warnung im
QA-Blatt und in der Zusammenfassung, nicht im Kleingedruckten.

**Der Kontenrahmen ist AASB, nicht HGB.** Die Zuordnung läuft über den
Kontonamen (Stichwortregeln, Kontenbibliothek, Kontogruppe), nicht über die
Kontonummer. Die SKR-Bereichstabelle ist hier bedeutungslos: ``1-10100`` ist
ein Bankkonto und kein SKR03-Kassenbestand.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Optional

from .core.hausconvention import Hausconvention
from .core.model import Klasse
from .engine.cascade import Engine
from .engine.laufprotokoll import Laufprotokoll
from .engine.qa import ABBRUCH, FLAG, QAReport, _r2
from .engine.setup import setup as setup_dialog
from .engine.spalten_status import baue_status
from .export import excel
from .readers.myob_export import lies_myob_export
from .views.leads import baue_lead_na, baue_lead_pl
from .views.net_debt import baue_net_debt
from .views.review_queue import baue_review_queue
from .views.schedules import baue_schedules
from .views.working_capital import baue_working_capital

ENTITY = "Projekt Luma (BT Imaging Pty Ltd)"

#: Namen der verbundenen Unternehmen (Setup-Frage 5). Ohne sie liest die
#: Software ``Accrued Interest - Aurora`` als Zinsabgrenzung und damit als Net
#: Debt, statt als Konzernforderung.
KONZERN = ["Aurora"]

#: Der Mandant berichtet in australischen Dollar. Reine Anzeige in den
#: Kopfzeilen — es wird nichts umgerechnet.
WAEHRUNG = "AUD"


@dataclass
class Quellen:
    bilanz: list[str]
    guv: list[str] = field(default_factory=list)


def run(quellen: Quellen, ausgabe: str, verbose: bool = True) -> dict:
    lp = Laufprotokoll()

    with lp.phase("Setup") as d:
        hc = Hausconvention.laden()
        s = setup_dialog(None, kontenrahmen="aasb", konzernnamen=KONZERN)
        d["detail"] = f"Hausconvention v{hc.version}; {s.rahmen_meldung[:80]}"

    with lp.phase("Einlesen MYOB-Export") as d:
        ledger, raster, diag = lies_myob_export(quellen.bilanz, quellen.guv,
                                                entity=ENTITY)
        d["detail"] = (f"{len(raster.jahre)} Jahre, {len(raster.quartale)} "
                       f"Quartale, {len(ledger.accounts)} Konten")

    jahre = list(raster.jahre)

    with lp.phase("Mapping (Kaskade AASB)") as d:
        eng = Engine(hc, kontenrahmen=s.kontenrahmen, konzernnamen=s.konzernnamen)
        mapped = eng.map_ledger(ledger)
        d["detail"] = f"{len(mapped)} Konten"

    # Der v2.8-Nachlauf (Saldenvorträge, Seitenwechsel, vorläufige HGB-Pfade)
    # bleibt aus: er arbeitet auf SKR-Kontonummern und HGB-Pfaden und hätte
    # hier nichts zu greifen. Ihn trotzdem laufen zu lassen, hieße, seine
    # Befunde als "keine Auffälligkeiten" zu lesen, obwohl er nur nicht
    # zuständig war.

    # Ohne Kontennachweis und ohne Abschluss ist keine Spalte abschlusstreu.
    status = baue_status(
        jahre, kontennachweis_perioden=set(), aggregiert_perioden=set(),
        quellen={p: (f"MYOB-Export, Blatt '{raster.blaetter.get(p, '?')}'"
                     + (f" — Rumpfjahr, nur "
                        f"{len(raster.quartale_je_jahr.get(p, []))} Quartale"
                        if p in diag.unvollstaendige_jahre else ""))
                 for p in jahre})

    with lp.phase("Views (jahresweise)") as d:
        nd = baue_net_debt(mapped, jahre, ENTITY)
        wc = baue_working_capital(mapped, jahre, ENTITY)
        lead_na = baue_lead_na(mapped, jahre, ENTITY)
        lead_pl = baue_lead_pl(mapped, jahre, ENTITY)
        review = baue_review_queue(mapped, jahre)
        d["detail"] = f"{len(lead_na.bloecke)} NA-Blöcke, {len(review)} Review"

    with lp.phase("Aufrisse (Schedules)") as d:
        schedules = baue_schedules(mapped, jahre)
        d["detail"] = f"{len(schedules.aufrisse)} Aufrisse"

    with lp.phase("QA-Diagnose") as d:
        qa = _qa(mapped, raster, diag, jahre)
        d["detail"] = (f"{len(qa.pruefungen)} Prüfungen, "
                       f"{len(qa.durchgefallen)} nicht bestanden")

    im_ek, ek_probe = _ergebnis_im_eigenkapital(mapped, jahre)
    qa.add("L11", "Eigenkapitalbewegung geht im Periodenergebnis auf",
           im_ek and all("deckungsgleich" in z for z in ek_probe), FLAG,
           ("Das Eigenkapital traegt das Periodenergebnis bereits; die Zeile "
            "'Jahresergebnis der Periode' entfaellt deshalb im Lead NA, sie "
            "zaehlte es sonst doppelt."
            if im_ek else
            "Das Eigenkapital traegt das Periodenergebnis nicht; es wird im "
            "Lead NA als eigene Zeile ergaenzt.")
           + (" Nicht jedes Jahr geht auf — die Reste unten sind zu klaeren."
              if not all("deckungsgleich" in z for z in ek_probe) else ""),
           ek_probe)

    meta = _meta(quellen, ledger, raster, mapped, diag, review, hc, s)
    with lp.phase("Excel-Ausgabe") as d:
        excel.schreibe_databook(
            ausgabe, mapped, nd, review, jahre, ENTITY, meta=meta, wc=wc,
            schedules=schedules, lead_na=lead_na, lead_pl=lead_pl,
            setup=s, status=status, qa=qa, raster=raster,
            waehrung=WAEHRUNG, ergebnis_im_eigenkapital=im_ek)
        d["detail"] = os.path.basename(ausgabe)
    lp.zaehle_arbeitsmappe(ausgabe)

    if verbose:
        _zusammenfassung(meta, raster, diag, mapped, review, qa, ausgabe, lp)
    return {"laufprotokoll": lp, "ledger": ledger, "raster": raster,
            "diagnose": diag, "mapped": mapped, "qa": qa, "status": status,
            "nd": nd, "wc": wc, "lead_na": lead_na, "lead_pl": lead_pl,
            "schedules": schedules, "review": review, "meta": meta, "setup": s}


# --------------------------------------------------------------------------

def _qa(mapped, raster, diag, jahre) -> QAReport:
    """Eingangsdiagnose. Geprüft wird das Einlesen und das Mapping, nicht die
    Buchführung des Mandanten.

    Die drei ersten Prüfblöcke belegen den Reader rechnerisch: eine
    vertauschte Wertespalte, ein übersehenes Blatt oder eine verlorene
    Kostenstelle fällt hier auf und nicht erst im fertigen Buch.
    """
    qa = QAReport()

    schlecht = {p: w for p, w in diag.bilanzidentitaet.items() if abs(w) > 0.05}
    qa.add("L1", "Bilanzidentität je Periodenspalte (Toleranz 0,05)",
           not schlecht, ABBRUCH,
           f"{len(diag.bilanzidentitaet)} Perioden geprüft, "
           f"{len(schlecht)} außerhalb der Toleranz",
           [f"{p}: {w:,.2f}" for p, w in sorted(schlecht.items())])

    guv_ab = {j: (q, g) for j, (q, g) in diag.guv_aufriss.items()
              if abs(q - g) > 0.05}
    qa.add("L2", "GuV: Summe der Quartale trifft das Jahresblatt",
           not guv_ab, ABBRUCH,
           f"{len(diag.guv_aufriss)} Geschäftsjahre geprüft, "
           f"{len(guv_ab)} abweichend",
           [f"{j}: Quartale {q:,.2f} gegen Jahr {g:,.2f}"
            for j, (q, g) in sorted(guv_ab.items())])

    bil_ab = {j: (q, g) for j, (q, g) in diag.bilanz_jahresende.items()
              if abs(q - g) > 0.05}
    qa.add("L3", "Bilanz: letztes Quartal trifft das Jahresblatt",
           not bil_ab, ABBRUCH,
           f"{len(diag.bilanz_jahresende)} Geschäftsjahre geprüft, "
           f"{len(bil_ab)} abweichend",
           [f"{j}: Q4 {q:,.2f} gegen Jahr {g:,.2f}"
            for j, (q, g) in sorted(bil_ab.items())])

    qa.add("L4", "Jedes Geschäftsjahr trägt vier Quartale",
           not diag.unvollstaendige_jahre, FLAG,
           ("alle Jahre vollständig" if not diag.unvollstaendige_jahre else
            f"{len(diag.unvollstaendige_jahre)} Rumpfjahr(e) — die GuV deckt "
            f"dort nur einen Teil des Jahres ab und ist mit den übrigen "
            f"Jahren nicht vergleichbar"),
           [f"{j}: {n} statt vier Quartale "
            f"({', '.join(raster.quartale_je_jahr.get(j, []))})"
            for j, n in sorted(diag.unvollstaendige_jahre.items())])

    qa.add("L5", "Trennblätter des Exports bewusst übersprungen", True, FLAG,
           f"{len(diag.uebersprungene_blaetter)} Blätter ohne Kopfzeile",
           list(diag.uebersprungene_blaetter))

    letzte = jahre[-1]
    ohne = [m for m in mapped if m.review]
    betrag = sum(abs(m.saldo(letzte)) for m in ohne)
    qa.add("L6", "Review-Queue unter 5 % der Konten",
           len(ohne) <= 0.05 * len(mapped), FLAG,
           f"{len(ohne)} von {len(mapped)} Konten, {betrag:,.2f} absolut "
           f"in {letzte}",
           [f"{m.konto} {m.bezeichnung}: {m.saldo(letzte):,.2f}"
            for m in sorted(ohne, key=lambda m: -abs(m.saldo(letzte)))[:20]])

    guv = [m for m in mapped if m.klasse == Klasse.PL]
    qa.add("L7", "GuV-Konten als solche erkannt", len(guv) > 0, ABBRUCH,
           f"{len(guv)} Konten der Klasse PL von {len(mapped)}")

    schief = []
    for jahr in jahre:
        aktiva = _r2(sum(m.saldo(jahr) for m in mapped
                         if m.bilanzseite == "AKTIVA"))
        passiva = _r2(sum(m.saldo(jahr) for m in mapped
                          if m.bilanzseite == "PASSIVA"))
        qa.bilanzsumme[jahr] = aktiva
        if abs(aktiva + passiva) > 0.05:
            schief.append(f"{jahr}: Aktiva {aktiva:,.2f}, Passiva {passiva:,.2f}")
    qa.add("L8", "Aktiva und Passiva gleichen sich nach dem Mapping aus",
           not schief, ABBRUCH,
           f"{len(jahre)} Jahre geprüft, {len(schief)} abweichend", schief)

    unbestimmt = [m for m in mapped
                  if m.klasse != Klasse.PL and m.bilanzseite is None]
    qa.add("L9", "Jedes Bilanzkonto hat eine Bilanzseite", not unbestimmt, FLAG,
           f"{len(unbestimmt)} Konten ohne bestimmbare Seite",
           [f"{m.konto} {m.bezeichnung}" for m in unbestimmt[:20]])

    aus_gruppe = [m for m in mapped if m.klasse == Klasse.PL
                  and m.regel_id == "gruppe:mandant"]
    qa.add("L10", "GuV-Positionen stammen aus dem Kontenrahmen",
           not aus_gruppe, FLAG,
           (f"{len(aus_gruppe)} von {len(guv)} GuV-Konten sind nach der "
            f"Kontogruppe des Mandanten gegliedert, nicht nach dem Rahmen — "
            f"der AASB-Rahmen v1.0 deckt mit seinen 67 FS Line Items nur die "
            f"Bilanz ab" if aus_gruppe else "alle Positionen aus dem Rahmen"),
           sorted({m.na_de for m in aus_gruppe})[:20])

    qa.annahmen.append(
        "Bilanz aus der Spalte 'Year' (Schlussbestand), GuV aus "
        "'PeriodMovementTYD' (Periodenbewegung). In der GuV ist 'Year' "
        "kumuliert und deshalb für Quartale unbrauchbar.")
    qa.annahmen.append(
        "Kontenrahmen AASB: die Position folgt dem Kontonamen, nicht der "
        "Kontonummer. Kein Kontennachweis vorhanden — alle Spalten vorläufig.")
    return qa


def _ergebnis_im_eigenkapital(mapped, jahre) -> tuple[bool, list[str]]:
    """Gemessen, nicht angenommen: waechst das Eigenkapital je Jahr um genau
    das Periodenergebnis?

    Eine DATEV-SuSa fuehrt das Ergebnis vor dem Abschluss NICHT im Kapital;
    ein MYOB-Export tut es. Wer das verwechselt, zaehlt das Ergebnis in der
    Net-Asset-Bruecke doppelt — und zwar in genau der Kontrollzeile, die
    belegen soll, dass die Bilanz schliesst.

    Entschieden wird ueber ``any`` und nicht ueber ``all``: eine Uebereinstimmung
    auf den Cent ist kein Zufall, waehrend eine Abweichung viele harmlose
    Gruende hat — eine Kapitalerhoehung, eine Ausschuettung, eine Umbuchung in
    den Ruecklagen. ``all`` zu verlangen hiesse, ein einziges Jahr mit einer
    Kapitalmassnahme entscheiden zu lassen. Die Jahre, die abweichen, stehen
    einzeln im QA-Blatt: dort ist die Abweichung ein eigener Befund und keine
    Nebenwirkung.
    """
    def ek(jahr):
        return _r2(sum(m.saldo(jahr) for m in mapped if m.klasse == Klasse.EQ))

    def pl(jahr):
        return _r2(sum(m.saldo(jahr) for m in mapped if m.klasse == Klasse.PL))

    probe, treffer = [], []
    for vorjahr, jahr in zip(jahre, jahre[1:]):
        delta = _r2(ek(jahr) - ek(vorjahr))
        ergebnis = pl(jahr)
        rest = _r2(delta - ergebnis)
        passt = abs(rest) <= 0.05
        treffer.append(passt)
        probe.append(
            f"{jahr}: Eigenkapital {ek(vorjahr):,.2f} -> {ek(jahr):,.2f} "
            f"(Delta {delta:,.2f}), Periodenergebnis {ergebnis:,.2f}"
            + (" — deckungsgleich" if passt else
               f" — Rest {rest:,.2f}, also eine Eigenkapitalbewegung "
               f"ausserhalb des Ergebnisses (Kapitalmassnahme, Ausschuettung "
               f"oder Umbuchung); zu klaeren"))
    return (any(treffer)), probe


def _meta(quellen, ledger, raster, mapped, diag, review, hc, s) -> dict:
    return {
        "Mandat": ENTITY,
        "Kontenrahmen": f"{s.kontenrahmen.name} {s.kontenrahmen.version}",
        "Hausconvention": f"v{hc.version}",
        "Quelldateien": ", ".join(diag.dateien),
        "Fingerprint": ledger.fingerprint,
        "Geschäftsjahre": ", ".join(raster.jahre),
        "Quartale im Mastersheet": ", ".join(raster.quartale),
        "Stichtage": ", ".join(f"{p}: {raster.stichtage[p]:%d.%m.%Y}"
                               for p in raster.jahre),
        "Konten": str(len(mapped)),
        "Review-Queue": str(len(review)),
        "Status": "VORLÄUFIG — kein Abschluss, kein Kontennachweis",
        "Verbundene Unternehmen": ", ".join(KONZERN) or "(keine erfasst)",
        "Währung": f"{WAEHRUNG} (Anzeige; keine Umrechnung)",
    }


def _zusammenfassung(meta, raster, diag, mapped, review, qa, ausgabe, lp) -> None:
    import collections

    print("=" * 100)
    print(f"{ENTITY} — Databook")
    print("=" * 100)
    for k, v in meta.items():
        print(f"  {k:26} {v}")

    print(f"\n  Mastersheet: {len(raster.spalten)} Spalten "
          f"({len(raster.quartale)} Quartale + {len(raster.jahre)} Jahre). "
          f"Alle übrigen Blätter: {len(raster.jahre)} Jahresspalten.")

    letzte = raster.jahre[-1]
    print(f"\nKlassenverteilung ({letzte})")
    kl = collections.Counter(m.klasse.value for m in mapped)
    ks = collections.defaultdict(float)
    for m in mapped:
        ks[m.klasse.value] += m.saldo(letzte)
    for name, n in sorted(kl.items(), key=lambda x: -x[1]):
        print(f"  {name:10}{n:>6} Konten{ks[name]:>20,.2f}")

    print(f"\nQuellen der Zuordnung")
    q = collections.Counter(m.quelle.value for m in mapped)
    for name, n in q.most_common():
        print(f"  {name:28}{n:>6}{n/len(mapped):>8.1%}")

    print(f"\nQA: {len(qa.pruefungen)} Prüfungen, "
          f"{len(qa.durchgefallen)} nicht bestanden")
    for p in qa.pruefungen:
        zeichen = "OK  " if p.bestanden else f"{p.schwere:4}"
        print(f"  {zeichen} {p.id}  {p.titel}")
        print(f"        {p.befund}")
        for zeile in p.details[:6]:
            print(f"          - {zeile}")

    print(f"\nReview-Queue: {len(review)} Konten")
    print(f"\nDatei: {ausgabe}")
    for zeile in lp.als_text():
        print(zeile)
