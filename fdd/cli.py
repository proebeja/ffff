"""End-to-end-CLI der ersten Scheibe: Eingabedatei -> Databook-Excel.

    python -m fdd.cli testdata/Testdaten_Eckart_SuSa_2022-2025-03.xlsx -o out.xlsx

Kette: Reader (formaterkannt) -> Engine-Kaskade -> Net-Debt-View + Review-Queue
-> Excel-Export im Hausformat. Gibt eine kurze Zusammenfassung auf stdout.
"""

from __future__ import annotations

import argparse
import os
import sys

from .core.hausconvention import Hausconvention
from .engine.cascade import Engine
from .engine.decision_log import Entscheidungsprotokoll
from .readers.detect import waehle_reader
from .views.net_debt import baue_net_debt
from .views.review_queue import baue_review_queue
from .views.schedules import baue_schedules
from .views.working_capital import baue_working_capital
from .export import excel


def run(eingabe: str, ausgabe: str, hc_pfad: str | None = None,
        verbose: bool = True) -> dict:
    hc = Hausconvention.laden(hc_pfad) if hc_pfad else Hausconvention.laden()
    reader = waehle_reader(eingabe)
    ledger = reader.lesen(eingabe)

    protokoll = Entscheidungsprotokoll()
    engine = Engine(hc, protokoll=protokoll)
    mapped = engine.map_ledger(ledger)

    nd = baue_net_debt(mapped, ledger.perioden, ledger.entity)
    wc = baue_working_capital(mapped, ledger.perioden, ledger.entity)
    schedules = baue_schedules(mapped, ledger.perioden)
    review = baue_review_queue(mapped, ledger.perioden)

    meta = {
        "Eingabedatei": os.path.basename(eingabe),
        "Reader": reader.name,
        "Entity": ledger.entity,
        "Perioden": ", ".join(ledger.perioden),
        "Kontennachweis vorhanden": "ja" if ledger.hat_kontennachweis else "nein",
        "Hausconvention-Version": hc.version,
        "Fingerprint (SHA-256/16)": ledger.fingerprint,
        "Konten gesamt": len(mapped),
        "davon Review": len(review),
        "davon technisch (TECH)": sum(1 for m in mapped if m.klasse.value == "TECH"),
        "Warnungen Reader": len(ledger.warnungen),
        "WC-Definition": (
            "Über alle Perioden identisch — jede Periode läuft durch dieselbe "
            "Klassifizierung (Klasse=TWC/OWC, abgeleitet aus dem HGB-Pfad via "
            f"Reklassifizierung, Hausconvention v{hc.version}). Voraussetzung "
            "für die spätere Kaufpreisanpassung: die WC-Definition am "
            "Completion-Stichtag muss exakt der des Referenz-WC entsprechen."
        ),
        "WC-Konten ohne NA-Zeile (Raster-Löcher)": len(wc.ohne_na_zeile),
        "Aufrisse (Schedules)": len(schedules.aufrisse),
        "Konten ohne Aufriss": len(schedules.ohne_aufriss),
    }
    excel.schreibe_databook(ausgabe, mapped, nd, review,
                            ledger.perioden, ledger.entity, meta=meta, wc=wc,
                            schedules=schedules)

    if verbose:
        _zusammenfassung(ledger, mapped, nd, review, ausgabe, meta)
    return {"ledger": ledger, "mapped": mapped, "nd": nd, "wc": wc,
            "schedules": schedules, "review": review, "meta": meta}


def _zusammenfassung(ledger, mapped, nd, review, ausgabe, meta) -> None:
    print("=" * 66)
    for k, v in meta.items():
        print(f"  {k:32} {v}")
    print("-" * 66)
    print("  Net-Debt-Zeilen (Subtotal je Periode):")
    for p in ledger.perioden:
        print(f"    {p:14} {nd.subtotal(p):>16,.0f}")
    if ledger.warnungen:
        print("-" * 66)
        for w in ledger.warnungen[:8]:
            print(f"  [warn] {w}")
        if len(ledger.warnungen) > 8:
            print(f"  … und {len(ledger.warnungen) - 8} weitere Warnungen")
    print("-" * 66)
    print(f"  Databook geschrieben: {ausgabe}")
    print("=" * 66)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="FDD-Databook (erste Scheibe)")
    ap.add_argument("eingabe", help="SuSa/SAP-Export/PDF-Kontennachweis")
    ap.add_argument("-o", "--ausgabe", default=None, help="Ziel-Excel (.xlsx)")
    ap.add_argument("--hausconvention", default=None, help="Pfad zu hausconvention.json")
    args = ap.parse_args(argv)

    ausgabe = args.ausgabe or (os.path.splitext(os.path.basename(args.eingabe))[0]
                               + "_Databook.xlsx")
    try:
        run(args.eingabe, ausgabe, args.hausconvention)
    except Exception as e:
        print(f"FEHLER: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
