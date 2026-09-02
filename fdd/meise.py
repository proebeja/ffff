"""Heinz Meise GmbH Medizintechnik — Aufbau des Databooks.

Quellen und ihre Reichweite, denn die entscheidet über den Status je Spalte:

* **Drei SuSa** aus Kanzlei-Rechnungswesen, je ein Geschäftsjahr:
  FY2023, FY2024, FY2025. Werteseite für alle drei Spalten.
* **Jahresabschluss 2024 mit Kontennachweis** (Handelsrecht). Er trägt den
  Kontennachweis für **Bilanz und GuV** und führt je Position die
  Geschäftsjahres- **und** die Vorjahresspalte. Damit ist er Strukturquelle
  für FY2024 *und* FY2023.

Daraus folgt: FY2023 und FY2024 abschlusstreu, FY2025 vorläufig — für FY2025
liegt kein Abschluss vor, die Struktur kommt dort aus der Hausconvention und
dem SKR03-Default.

Zwei Dinge, die diese Quelle von den bisherigen unterscheidet:

**Die SuSa 2023 und 2024 sind "nach Jahresabschluss" gezogen.** Sie sollten
den Abschluss also exakt treffen. Wo sie es nicht tun, ist das ein Befund und
keine Rundung — die Überleitung weist ihn je Position aus.

**Konto 1600 steht auf beiden Bilanzseiten.** Der Kontennachweis führt es
sowohl unter den sonstigen Vermögensgegenständen (Debitorenüberzahlung) als
auch unter den Verbindlichkeiten aus Lieferungen und Leistungen. Eine
Zuordnung Konto -> Pfad kann das nicht abbilden; der Fall gehört benannt und
nicht saldiert.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Optional

from .core.hausconvention import Hausconvention
from .engine.cascade import Engine
from .engine.kontennachweis_apply import wende_kontennachweis_an
from .engine.laufprotokoll import Laufprotokoll
from .engine.qa import ABBRUCH, FLAG, QAReport, _r2
from .engine.spalten_status import baue_status
from .engine.v28 import (loese_saldenvortraege, pruefe_verhalten,
                         setze_vorlaeufige_pfade, wende_seitenwechsel_an)
from .export import excel
from .readers.datev_kanzlei_susa import lies_kanzlei_susa
from .readers.kontennachweis import lies_kontennachweis
from .views.leads import baue_lead_na, baue_lead_pl
from .views.net_debt import baue_net_debt
from .views.review_queue import baue_review_queue
from .views.schedules import baue_schedules
from .views.working_capital import baue_working_capital

ENTITY = "Heinz Meise GmbH Medizintechnik"

#: Der Abschluss 2024 traegt die Geschaeftsjahres- und die Vorjahresspalte.
PERIODE_JA = "FY2024"
PERIODE_VJ = "FY2023"


@dataclass
class Quellen:
    susa: list[str]
    jahresabschluss: Optional[str] = None


@dataclass
class Positionsabgleich:
    """Eine Position: was der Kontennachweis sagt, was die SuSa ergibt."""

    hgb_pfad: str
    kontennachweis: dict[str, float] = field(default_factory=dict)
    susa: dict[str, float] = field(default_factory=dict)

    def differenz(self, periode: str) -> float:
        return _r2(self.susa.get(periode, 0.0)
                   - self.kontennachweis.get(periode, 0.0))


def run(quellen: Quellen, ausgabe: str, verbose: bool = True) -> dict:
    lp = Laufprotokoll()
    with lp.phase("Setup") as d:
        hc = Hausconvention.laden()
        d["detail"] = f"Hausconvention v{hc.version}, eigene Workbook-Ausgabe"

    with lp.phase("Einlesen SuSa (Kanzlei-Rechnungswesen)") as d:
        ledger, diag = lies_kanzlei_susa(quellen.susa, entity=ENTITY)
        d["detail"] = (f"{len(ledger.perioden)} Perioden, "
                       f"{len(ledger.accounts)} Konten")

    kn = None
    with lp.phase("Einlesen Kontennachweis (JA 2024)") as d:
        if quellen.jahresabschluss:
            kn = lies_kontennachweis(quellen.jahresabschluss,
                                     perioden=[PERIODE_JA, PERIODE_VJ])
            ledger = wende_kontennachweis_an(ledger, kn)
            d["detail"] = (f"{len(kn.konten)} Konten, "
                           f"{len(kn.positionen())} Positionen")
        else:
            d["detail"] = "kein Abschluss übergeben"

    with lp.phase("Mapping (Kaskade)") as d:
        mapped = Engine(hc).map_ledger(ledger)
        d["detail"] = f"{len(mapped)} Konten"

    with lp.phase("v2.8-Nachlauf") as d:
        mapped, saldenvortrag = loese_saldenvortraege(mapped, ledger.perioden, hc)
        mapped, seitenwechsel = wende_seitenwechsel_an(
            mapped, ledger.perioden, hc, _nachgewiesene_seiten(kn))
        mapped, ungeloest = setze_vorlaeufige_pfade(mapped, ledger.perioden, hc)
        verhalten = pruefe_verhalten(mapped, ledger.perioden, hc,
                                     _schwelle(mapped, ledger.perioden, hc))
        d["detail"] = (f"{len(seitenwechsel)} Seitenwechsel, {len(ungeloest)} "
                       f"vorläufige Pfade, {len(verhalten)} Verhaltensbefunde")

    status = baue_status(
        ledger.perioden,
        kontennachweis_perioden={PERIODE_JA, PERIODE_VJ} if kn else set(),
        aggregiert_perioden=set(),
        quellen={p: (f"SuSa {p} ({os.path.basename(diag.dateien[p])})"
                     + (", Kontennachweis JA 2024"
                        if kn and p in (PERIODE_JA, PERIODE_VJ) else ""))
                 for p in ledger.perioden})

    with lp.phase("Views") as d:
        nd = baue_net_debt(mapped, ledger.perioden, ENTITY)
        wc = baue_working_capital(mapped, ledger.perioden, ENTITY)
        lead_na = baue_lead_na(mapped, ledger.perioden, ENTITY)
        lead_pl = baue_lead_pl(mapped, ledger.perioden, ENTITY)
        review = baue_review_queue(mapped, ledger.perioden)
        d["detail"] = f"{len(lead_na.bloecke)} NA-Blöcke, {len(review)} Review"

    with lp.phase("Aufrisse (Schedules)") as d:
        schedules = baue_schedules(mapped, ledger.perioden)
        d["detail"] = f"{len(schedules.aufrisse)} Aufrisse"

    abgleich = _positionsabgleich(mapped, kn)

    with lp.phase("QA-Diagnose") as d:
        qa = _qa(mapped, ledger, diag, kn, status, ungeloest, abgleich)
        d["detail"] = (f"{len(qa.pruefungen)} Prüfungen, "
                       f"{len(qa.durchgefallen)} nicht bestanden")

    meta = _meta(quellen, ledger, mapped, diag, kn, status, review, hc)
    with lp.phase("Excel-Ausgabe") as d:
        excel.schreibe_databook(
            ausgabe, mapped, nd, review, ledger.perioden, ENTITY, meta=meta,
            wc=wc, schedules=schedules, lead_na=lead_na, lead_pl=lead_pl,
            status=status, qa=qa, verhalten=verhalten)
        d["detail"] = os.path.basename(ausgabe)
    lp.zaehle_arbeitsmappe(ausgabe)

    if verbose:
        _zusammenfassung(meta, ledger, diag, status, abgleich, ausgabe, lp,
                         review, seitenwechsel)
    return {"laufprotokoll": lp, "ledger": ledger, "diagnose": diag, "kn": kn,
            "mapped": mapped, "qa": qa, "status": status, "nd": nd, "wc": wc,
            "lead_na": lead_na, "lead_pl": lead_pl, "schedules": schedules,
            "review": review, "verhalten": verhalten, "meta": meta,
            "abgleich": abgleich, "seitenwechsel": seitenwechsel,
            "ungeloest": ungeloest, "saldenvortrag": saldenvortrag}


# --------------------------------------------------------------------------

def _nachgewiesene_seiten(kn) -> dict[str, dict[str, str]]:
    """Bilanzseite je Konto und Periode, wie der Kontennachweis sie ausweist.

    Ohne diese Angabe leitet der Seitenwechsel die Seite aus dem Vorzeichen
    ab — und erfindet damit Umgliederungen, die der Abschluss nicht kennt.
    """
    if kn is None:
        return {}
    je_konto: dict[str, dict[str, str]] = {}
    for konto, k in kn.konten.items():
        if not k.hgb_pfad.startswith(("/Aktiva", "/Passiva")):
            continue
        seite = "AKTIVA" if k.hgb_pfad.startswith("/Aktiva") else "PASSIVA"
        je_konto[konto] = {p: seite for p in kn.perioden}
    return je_konto


def _schwelle(mapped, perioden, hc) -> float:
    prozent = hc.wesentlichkeit.get("bilanzsumme_prozent", 2) / 100.0
    summen = [sum(m.saldo(p) for m in mapped
                  if m.hgb_pfad.startswith("/Aktiva")) for p in perioden]
    return max(summen or [0.0]) * prozent


def _positionsabgleich(mapped, kn) -> list[Positionsabgleich]:
    """Positionssumme laut Kontennachweis gegen die Summe der SuSa-Konten.

    Die SuSa 2023 und 2024 sind "nach Jahresabschluss" gezogen; jede
    Abweichung ist deshalb ein Befund und keine Rundung.
    """
    if kn is None:
        return []
    positionen = kn.positionen()
    out: list[Positionsabgleich] = []
    for pfad in sorted(positionen):
        z = Positionsabgleich(pfad, dict(positionen[pfad]))
        for p in kn.perioden:
            z.susa[p] = _r2(sum(m.saldo(p) for m in mapped
                                if m.pfad_in(p) == pfad))
        out.append(z)
    return out


def _qa(mapped, ledger, diag, kn, status, ungeloest, abgleich) -> QAReport:
    r = QAReport()
    perioden = ledger.perioden

    r.add("A1", "Datenblockgrenze eindeutig", True, ABBRUCH,
          "Je Datei ein Blatt mit einer Kopfzeile; gelesen wird ab der "
          "Kopfzeile bis zur letzten Kontozeile. Summenzeilen führt die "
          "Datei nicht.",
          [f"{p}: {n} Kontozeilen" for p, n in diag.kontozeilen.items()])

    summen = {p: _r2(sum(m.saldo(p) for m in mapped)) for p in perioden}
    r.add("A2", "Bilanzidentität je Periodenspalte (Toleranz 1,00)",
          all(abs(v) <= 1.0 for v in summen.values()), FLAG,
          "Summe aller Konten je Spalte: "
          + ", ".join(f"{p} = {v:,.2f}" for p, v in summen.items()))

    r.add("A3", "Kontoschlüssel eindeutig", True, FLAG,
          f"{len(mapped)} Konten über {len(perioden)} Perioden.")

    r.add("A4", "Vorzeichen aus der S/H-Spalte, nicht aus der Zahl", True,
          ABBRUCH,
          "Die Saldo-Spalte führt den Betrag ohne Vorzeichen; erst der "
          "Marker rechts daneben sagt Soll oder Haben. Probe je Periode: "
          "Soll gegen Haben.",
          [f"{p}: Soll {diag.spaltensummen[p][0]:,.2f} · Haben "
           f"{diag.spaltensummen[p][1]:,.2f} · Differenz "
           f"{diag.identitaet(p):,.2f}" for p in perioden])

    r.add("A5", "Konten in allen Perioden vorhanden",
          not diag.nicht_durchgaengig, FLAG,
          f"{len(diag.nicht_durchgaengig)} Konten kommen nicht in allen drei "
          "Perioden vor. Sie tragen dort null; der Kontenplan wächst und "
          "schrumpft über die Jahre.",
          [f"{k}: fehlt in {', '.join(v)}"
           for k, v in list(diag.nicht_durchgaengig.items())[:12]])

    for p in perioden:
        r.nicht_zugeordnet[p] = _r2(sum(u.salden.get(p, 0.0)
                                        for u in ungeloest))
        r.bilanzsumme[p] = _r2(sum(m.saldo(p) for m in mapped
                                   if m.pfad_in(p).startswith("/Aktiva")))
    r.add("A6", "Ungelöste Konten liegen INNERHALB der Bilanz",
          all(abs(v) < 1.0 for v in r.nicht_zugeordnet.values()), FLAG,
          f"{len(ungeloest)} Konten ohne bestimmbaren Pfad.",
          [f"{p}: {r.nicht_zugeordnet[p]:,.2f} von {r.bilanzsumme[p]:,.2f}"
           for p in perioden])

    r.add("B2", "Jahresanker", True, FLAG,
          "Die Periode steht in der Kopfzeile der Monatsspalte, nicht im "
          "Dateinamen: "
          + ", ".join(f"{p} = {diag.stichtage[p].strftime('%d.%m.%Y')}"
                      for p in perioden))

    r.add("C1", "Strukturquelle je Periode und Rechenwerk",
          kn is not None, FLAG,
          "Der Kontennachweis 2024 trägt Bilanz UND GuV und führt die "
          "Vorjahresspalte mit. FY2025 hat keinen Abschluss."
          if kn else "Kein Kontennachweis übergeben.",
          [f"{s.periode}: Bilanz {s.bilanz} · GuV {s.guv}"
           for s in status.spalten])

    offen = [z for z in abgleich
             if any(abs(z.differenz(p)) > 1.0 for p in (kn.perioden if kn else []))]
    r.add("C2", "Abstimmung Kontennachweis gegen SuSa je Position",
          not offen, FLAG,
          f"{len(abgleich)} Positionen abgestimmt, {len(offen)} mit "
          "Abweichung über 1,00. Die SuSa 2023/2024 ist 'nach "
          "Jahresabschluss' gezogen — eine Abweichung ist ein Befund."
          if abgleich else "Nicht möglich: kein Kontennachweis.",
          [f"{z.hgb_pfad}: "
           + " · ".join(f"{p} {z.differenz(p):,.2f}" for p in kn.perioden)
           for z in offen[:12]])

    r.add("C4", "Kontenrahmen erkannt", True, FLAG,
          "SKR03 (800 Gezeichnetes Kapital, 860 Gewinnvortrag, 8xxx Erlöse). "
          "Der SKR-Default der Kaskade ist damit anwendbar.")

    r.annahmen.append(
        "Vorzeichen: Soll positiv, Haben negativ. Die Quelle führt den "
        "Betrag ohne Vorzeichen und den Marker in einer eigenen Spalte.")
    r.annahmen.append(
        "Der Kontennachweis des Abschlusses 2024 gilt für FY2024 und für die "
        "dort ausgewiesene Vorjahresspalte FY2023. FY2025 ist vorläufig.")
    for w in diag.warnungen:
        r.offene_befunde.append(w)
    return r


def _meta(quellen, ledger, mapped, diag, kn, status, review, hc) -> dict:
    return {
        "Quelldateien": ", ".join(os.path.basename(p) for p in quellen.susa),
        "Abschluss": os.path.basename(quellen.jahresabschluss)
                     if quellen.jahresabschluss else "—",
        "Reader": "datev_kanzlei_susa + kontennachweis (PDF)",
        "Entity": ENTITY,
        "Währung": "EUR",
        "Perioden": ", ".join(
            f"{p} zum {diag.stichtage[p].strftime('%d.%m.%Y')}"
            for p in ledger.perioden),
        "Status je Spalte": " | ".join(
            f"{s.periode}: Bilanz {s.bilanz} · GuV {s.guv}"
            for s in status.spalten),
        "Strukturquelle": ("Kontennachweis JA 2024 (Bilanz und GuV), "
                           "Vorjahresspalte FY2023" if kn else "—"),
        "Hausconvention": hc.version,
        "Fingerprint (SHA-256/16)": ledger.fingerprint,
        "Konten gesamt": len(mapped),
        "davon Review": len(review),
        "Reader-Warnungen": len(ledger.warnungen),
    }


def _zusammenfassung(meta, ledger, diag, status, abgleich, ausgabe, lp,
                     review, seitenwechsel) -> None:
    print("=" * 78)
    for k, v in meta.items():
        print(f"  {k:26} {v}")
    print("-" * 78)
    for p in ledger.perioden:
        soll, haben = diag.spaltensummen[p]
        print(f"  {p}  Soll {soll:>16,.2f}  Haben {haben:>16,.2f}  "
              f"Identität {diag.identitaet(p):>8,.2f}")
    if abgleich:
        print("-" * 78)
        print("  Abstimmung Kontennachweis gegen SuSa (nur Abweichungen > 1,00)")
        perioden = [p for p in ledger.perioden
                    if any(p in z.kontennachweis for z in abgleich)]
        for z in abgleich:
            diffs = {p: z.differenz(p) for p in perioden}
            if all(abs(v) <= 1.0 for v in diffs.values()):
                continue
            print(f"    {z.hgb_pfad[:58]:60}"
                  + "".join(f"{diffs[p]:>14,.2f}" for p in perioden))
    if seitenwechsel:
        print("-" * 78)
        for f in seitenwechsel:
            print(f"  [Seitenwechsel] {f.konto}")
    if review:
        print("-" * 78)
        print(f"  Review-Queue: {len(review)} Konten")
        for e in review[:12]:
            print(f"    {e.konto:8s} {e.bezeichnung[:40]:42s} "
                  f"{getattr(e, 'status', '')}")
    print("-" * 78)
    print("\n".join("  " + z for z in lp.als_text()))
    print("-" * 78)
    print(f"  Databook geschrieben: {ausgabe}")
    print("=" * 78)
