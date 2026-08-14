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
from .engine.kontennachweis_apply import wende_kontennachweis_an
from .engine.reconciliation import reconcile
from .engine.setup import setup
from .readers.detect import waehle_reader
from .readers.sap_bw_monate import lies_sap_jahre
from .readers.kontennachweis import lies_kontennachweis
from .views.leads import baue_lead_na, baue_lead_pl
from .views.net_debt import baue_net_debt
from .views.review_queue import baue_review_queue
from .views.schedules import baue_schedules
from .views.working_capital import baue_working_capital
from .export import excel


def run(eingabe, ausgabe: str, hc_pfad: str | None = None,
        verbose: bool = True, kontennachweis: str | None = None,
        jahresabschluss: str | None = None) -> dict:
    """``eingabe`` ist eine Datei — oder eine Liste von Jahresdateien, die zu
    einem mehrperiodigen Ledger verschmolzen werden (SAP-Monatsexporte)."""
    hc = Hausconvention.laden(hc_pfad) if hc_pfad else Hausconvention.laden()
    monats_diagnosen = []
    if isinstance(eingabe, (list, tuple)):
        ledger, monats_diagnosen = lies_sap_jahre(list(eingabe))
        reader_name = "sap_bw_monate (mehrjährig)"
        eingabe_label = ", ".join(os.path.basename(p) for p in eingabe)
    else:
        reader = waehle_reader(eingabe)
        ledger = reader.lesen(eingabe)
        reader_name = reader.name
        eingabe_label = os.path.basename(eingabe)
    susa_konten = {a.konto for a in ledger.accounts}

    # Setup-Dialog: der Kontennachweis ist die erste Frage und entscheidet
    # über den Modus (abschlusstreu vs. vorläufiger Default-Modus).
    kn = None
    if kontennachweis:
        kn = lies_kontennachweis(kontennachweis, ledger.perioden)
        ledger = wende_kontennachweis_an(ledger, kn)
    # Konten, die die Quelle selbst als nicht abschlussrelevant ausweist,
    # zählen nicht in die Abdeckung — sonst drückt die Aussonderung die Quote.
    relevante = [a for a in ledger.accounts if a.kontotyp != "technisch"]
    mit_kn_struktur = sum(1 for a in relevante if a.fs_pfad)
    # Bringt die Quelle ihre FS-Gliederung selbst mit (SAP-Hierarchie), ist das
    # eine vollwertige Strukturquelle — auch ohne separaten Kontennachweis.
    eingebettet = ("FS-Hierarchie des Exports"
                   if ledger.hat_kontennachweis and mit_kn_struktur else None)
    setup_ergebnis = setup(kontennachweis, len(relevante), mit_kn_struktur,
                           eingebettete_struktur=eingebettet)

    protokoll = Entscheidungsprotokoll()
    engine = Engine(hc, protokoll=protokoll)
    mapped = engine.map_ledger(ledger)
    recon = reconcile(mapped, kn, ledger.perioden, susa_konten) if kn else None

    ja_recon = None
    if jahresabschluss:
        from .engine.ja_reconciliation import reconcile_gegen_ja
        ja_recon = reconcile_gegen_ja(mapped, jahresabschluss, ledger.perioden)

    nd = baue_net_debt(mapped, ledger.perioden, ledger.entity)
    lead_na = baue_lead_na(mapped, ledger.perioden, ledger.entity)
    lead_pl = baue_lead_pl(mapped, ledger.perioden, ledger.entity)
    wc = baue_working_capital(mapped, ledger.perioden, ledger.entity)
    schedules = baue_schedules(mapped, ledger.perioden)
    review = baue_review_queue(mapped, ledger.perioden)

    meta = {
        "Eingabedatei": eingabe_label,
        "Reader": reader_name,
        "Entity": ledger.entity,
        "Perioden": ", ".join(ledger.perioden),
        "Kontennachweis vorhanden": "ja" if ledger.hat_kontennachweis else "nein",
        "Modus": setup_ergebnis.databook_kennzeichen,
        "Setup-Meldung": setup_ergebnis.meldung,
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
        "Struktur aus Abschluss": (
            f"{mit_kn_struktur} von {len(relevante)} abschlussrelevanten Konten"
            + (f" ({len(ledger.accounts) - len(relevante)} von der Quelle ausgesondert)"
               if len(relevante) != len(ledger.accounts) else "")),
        "Konten nur im Kontennachweis": len(recon.nur_im_kn) if recon else 0,
    }
    if monats_diagnosen:
        brueche = sum(len(d.ueberlappungs_bruch) for d in monats_diagnosen)
        meta["Monatsscheiben zusammengeführt"] = ", ".join(
            f"{d.jahr}: {len(d.scheiben)} kumulierte Scheiben" for d in monats_diagnosen)
        meta["Brüche in der kumulierten Reihe"] = brueche
    if ja_recon:
        meta["Reconciliation gegen Jahresabschluss"] = (
            f"{len(ja_recon.zeilen)} Positionen, "
            f"{len(ja_recon.zeilen_mit_differenz)} mit Differenz")
    if setup_ergebnis.anforderung:
        meta["Datenanforderung"] = setup_ergebnis.anforderung
    excel.schreibe_databook(ausgabe, mapped, nd, review,
                            ledger.perioden, ledger.entity, meta=meta, wc=wc,
                            schedules=schedules, recon=recon,
                            setup=setup_ergebnis, lead_na=lead_na,
                            lead_pl=lead_pl, ja_recon=ja_recon)

    if verbose:
        _zusammenfassung(ledger, mapped, nd, review, ausgabe, meta)
    return {"ledger": ledger, "mapped": mapped, "nd": nd, "wc": wc,
            "schedules": schedules, "review": review, "meta": meta,
            "recon": recon, "setup": setup_ergebnis, "kn": kn,
            "lead_na": lead_na, "lead_pl": lead_pl, "ja_recon": ja_recon,
            "monats_diagnosen": monats_diagnosen}


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
    ap.add_argument("eingabe", nargs="+",
                    help="SuSa/SAP-Export/PDF-Kontennachweis; mehrere "
                         "SAP-Jahresdateien werden zu einem Ledger verschmolzen")
    ap.add_argument("-o", "--ausgabe", default=None, help="Ziel-Excel (.xlsx)")
    ap.add_argument("--hausconvention", default=None, help="Pfad zu hausconvention.json")
    ap.add_argument("--kontennachweis", default=None,
                    help="Kontennachweis (PDF/Excel) — maßgebliche Strukturquelle")
    ap.add_argument("--jahresabschluss", default=None,
                    help="Bilanz/GuV nach Jahresabschluss — Reconciliation-Ziel")
    args = ap.parse_args(argv)

    eingabe = args.eingabe if len(args.eingabe) > 1 else args.eingabe[0]
    erste = args.eingabe[0]
    ausgabe = args.ausgabe or (os.path.splitext(os.path.basename(erste))[0]
                               + "_Databook.xlsx")
    try:
        run(eingabe, ausgabe, args.hausconvention,
            kontennachweis=args.kontennachweis,
            jahresabschluss=args.jahresabschluss)
    except Exception as e:
        print(f"FEHLER: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
