"""Excel-Export im Hausformat.

Hausformat (Handover Abschnitt 3): Arial, Teal-Palette, Zahlenformat
``#,##0;(#,##0);"-"``, Kontrollzeile. Sichtbare, prüfbare Formeln.

Single Source of Truth: das **Mastersheet** ist das eine Zuhause jeder
Kontozahl. Der Net-Debt-Tab rechnet nichts neu, sondern summiert die
Mastersheet-Werte über sichtbare ``SUMIFS``-Formeln, die auf die
Klasse-/NA-Spalte zeigen. Damit kann keine Zahl an zwei Stellen auseinander-
laufen, und der Prüfer sieht nachvollziehbare Formeln.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from ..core.model import Klasse, MappedAccount
from ..views.net_debt import NetDebtView
from ..views.review_queue import ReviewEintrag
from ..views.schedules import Aufriss, Schedules
from ..views.working_capital import WCView

# ---- Hausformat-Konstanten ------------------------------------------------
ZAHLENFORMAT = '#,##0;(#,##0);"-"'
FONT_NAME = "Arial"
TEAL = "1F6F6F"           # Kopf-/Akzentfarbe
TEAL_HELL = "D6E8E8"      # Zwischensummen-Hinterlegung
GRAU_HELL = "F2F2F2"

_kopf_fill = PatternFill("solid", fgColor=TEAL)
_sub_fill = PatternFill("solid", fgColor=TEAL_HELL)
_grau_fill = PatternFill("solid", fgColor=GRAU_HELL)
_kopf_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
_bold = Font(name=FONT_NAME, bold=True, size=10)
_normal = Font(name=FONT_NAME, size=10)
_duenn = Side(style="thin", color="BFBFBF")
_rahmen = Border(bottom=_duenn)
_top_double = Border(top=Side(style="double", color=TEAL))


@dataclass
class MastersheetLayout:
    """Merkt sich, wo im Mastersheet welche Spalte/Zeilen liegen, damit
    Aufrisse und Kontrollzeilen korrekt referenzieren."""

    sheetname: str
    erste_datenzeile: int
    letzte_datenzeile: int
    spalte_klasse: int
    spalte_na: int
    perioden_spalten: dict[str, int]
    # Schlüssel ist die Objekt-Identität des MappedAccount, NICHT die Kontonummer:
    # reale Charts (z.B. Namur) tragen dieselbe Kontonummer mehrfach als
    # verschiedene Unterkonten — ein konto->Zeile-Map würde die verwechseln.
    zeile_je_account: dict[int, int]

    def bereich(self, spalte: int) -> str:
        c = get_column_letter(spalte)
        return f"'{self.sheetname}'!${c}${self.erste_datenzeile}:${c}${self.letzte_datenzeile}"

    def wert_zelle(self, m: MappedAccount, periode: str) -> str:
        """Zellbezug auf den Saldo genau dieses Kontos je Periode."""
        c = get_column_letter(self.perioden_spalten[periode])
        return f"'{self.sheetname}'!${c}${self.zeile_je_account[id(m)]}"

    def klasse_zelle(self, m: MappedAccount) -> str:
        c = get_column_letter(self.spalte_klasse)
        return f"'{self.sheetname}'!${c}${self.zeile_je_account[id(m)]}"


def _ref_formel(aufriss: "Optional[AufrissRef]", periode: str, nd_teil: bool) -> str:
    """Formel einer Lead-Zeile: Verweis auf die passende Aufriss-Summenzelle.
    nd_teil=True -> Net-Debt-Lead (gemischt: thereof ND, sonst Summe);
    nd_teil=False -> Working-Capital-Lead (gemischt: operating, sonst Summe)."""
    if aufriss is None:
        return "0"
    if aufriss.is_mixed:
        zelle = aufriss.thereof_nd[periode] if nd_teil else aufriss.operating[periode]
    else:
        zelle = aufriss.total[periode]
    return "=" + zelle


def _krit(text: str) -> str:
    """Escaping für ein SUMIFS-Text-Kriterium: Doppelquotes verdoppeln (sonst
    bricht der Formel-String), Excel-Wildcards (* ? ~) mit ~ neutralisieren
    (sonst ändern sie stillschweigend die Match-Semantik)."""
    text = text.replace('"', '""')
    for w in ("~", "*", "?"):
        text = text.replace(w, "~" + w)
    return text


@dataclass
class AufrissRef:
    """Zellbezüge auf die Summenzeile eines Aufrisses (Quelle der Lead-Zeilen)."""

    sheetname: str
    is_mixed: bool
    total: dict[str, str]        # nicht-gemischt: Periode -> Summenzelle
    operating: dict[str, str]    # gemischt: Periode -> operating-Summenzelle
    thereof_nd: dict[str, str]   # gemischt: Periode -> thereof-ND-Summenzelle


def schreibe_databook(pfad: str, mapped: list[MappedAccount], nd: NetDebtView,
                      review: list[ReviewEintrag], perioden: list[str],
                      entity: str, meta: Optional[dict] = None,
                      wc: "Optional[WCView]" = None,
                      schedules: "Optional[Schedules]" = None) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    layout = _schreibe_mastersheet(wb, mapped, perioden, entity)
    refs = _schreibe_schedules(wb, schedules, layout, perioden, entity) if schedules else {}
    _schreibe_net_debt(wb, nd, layout, perioden, entity, refs)
    if wc is not None:
        _schreibe_working_capital(wb, wc, layout, perioden, entity, refs)
    _schreibe_review(wb, review, perioden)
    if meta:
        _schreibe_info(wb, meta)
    wb.save(pfad)


# ---- Aufriss-Schicht (Schedules) -----------------------------------------
def _schreibe_schedules(wb, schedules: Schedules, layout: MastersheetLayout,
                        perioden, entity) -> dict[str, AufrissRef]:
    """Schreibt je Net-Asset-Zeile einen Aufriss-Tab. Jede Kontozeile
    referenziert den Mastersheet-Einzelwert (SSOT); die Summenzeile summiert.
    Gemischte Aufrisse (NA_OA/OL/OP) führen zwei Spalten je Periode: operating
    (Klasse OWC) und thereof ND (Klasse ND), gesteuert durch die Klasse-Zelle
    des Mastersheets. Leere Aufrisse werden ausgeblendet. Gibt je NA-Zeile die
    Summenzell-Bezüge zurück, aus denen die Leads ziehen."""
    refs: dict[str, AufrissRef] = {}
    for a in schedules.aufrisse:
        refs[a.na_de] = (_schreibe_mixed_aufriss(wb, a, layout, perioden)
                         if a.is_mixed
                         else _schreibe_einfacher_aufriss(wb, a, layout, perioden))
    return refs


def _aufriss_kopf(ws, a: Aufriss) -> None:
    ws.cell(1, 2, f"Aufriss {a.sheetname}: {a.na_de} / {a.na_en}").font = Font(
        name=FONT_NAME, bold=True, size=12, color=TEAL)
    speist = {"ND": "Net-Debt-Lead", "WC": "Working-Capital-Lead",
              "beide": "WC-Lead (operating) + ND-Lead (thereof ND)"}[a.speist]
    ws.cell(2, 2, f"in EUR · Einzelkonten je Periode aus dem Mastersheet · speist {speist}").font = Font(
        name=FONT_NAME, italic=True, size=9)


def _schreibe_einfacher_aufriss(wb, a: Aufriss, layout, perioden) -> AufrissRef:
    ws = wb.create_sheet(a.sheetname)
    ws.sheet_view.showGridLines = False
    _aufriss_kopf(ws, a)
    hz = 4
    kopf = ["", "Konto", "Bezeichnung", "Klasse"] + list(perioden)
    for c, t in enumerate(kopf[1:], start=2):
        cell = ws.cell(hz, c, t); cell.font = _kopf_font; cell.fill = _kopf_fill
    p0 = 5
    r = hz + 1
    erste = r
    for m in a.konten:
        ws.cell(r, 2, m.konto).font = _normal
        ws.cell(r, 3, m.bezeichnung).font = _normal
        kc = ws.cell(r, 4, m.klasse.value); kc.font = _bold; kc.fill = _klasse_fill(m.klasse)
        for i, p in enumerate(perioden):
            cell = ws.cell(r, p0 + i, "=" + layout.wert_zelle(m, p))
            cell.number_format = ZAHLENFORMAT; cell.font = _normal
        r += 1
    letzte = r - 1
    total = {}
    tc = ws.cell(r, 3, "Summe (Quelle Lead)"); tc.font = _bold
    for i, p in enumerate(perioden):
        col = get_column_letter(p0 + i)
        cell = ws.cell(r, p0 + i, f"=SUM({col}{erste}:{col}{letzte})")
        cell.number_format = ZAHLENFORMAT; cell.font = _bold; cell.fill = _sub_fill
        cell.border = _top_double
        total[p] = f"'{a.sheetname}'!${col}${r}"
    _breiten(ws, {2: 12, 3: 40, 4: 8})
    for i in range(len(perioden)):
        ws.column_dimensions[get_column_letter(p0 + i)].width = 14
    if a.ist_leer:
        ws.sheet_state = "hidden"
    return AufrissRef(a.sheetname, False, total, {}, {})


def _schreibe_mixed_aufriss(wb, a: Aufriss, layout, perioden) -> AufrissRef:
    ws = wb.create_sheet(a.sheetname)
    ws.sheet_view.showGridLines = False
    _aufriss_kopf(ws, a)
    hz = 4
    ws.cell(hz, 2, "Konto").font = _kopf_font; ws.cell(hz, 2).fill = _kopf_fill
    ws.cell(hz, 3, "Bezeichnung").font = _kopf_font; ws.cell(hz, 3).fill = _kopf_fill
    ws.cell(hz, 4, "Klasse").font = _kopf_font; ws.cell(hz, 4).fill = _kopf_fill
    # je Periode zwei Spalten: operating | thereof ND
    p0 = 5
    op_col, nd_col = {}, {}
    for i, p in enumerate(perioden):
        oc, nc = p0 + 2 * i, p0 + 2 * i + 1
        op_col[p], nd_col[p] = oc, nc
        for col, lab in ((oc, f"{p} operating"), (nc, f"{p} thereof ND")):
            cell = ws.cell(hz, col, lab); cell.font = _kopf_font; cell.fill = _kopf_fill
    r = hz + 1
    erste = r
    for m in a.konten:
        ws.cell(r, 2, m.konto).font = _normal
        ws.cell(r, 3, m.bezeichnung).font = _normal
        kc = ws.cell(r, 4, m.klasse.value); kc.font = _bold; kc.fill = _klasse_fill(m.klasse)
        kz = layout.klasse_zelle(m)
        for p in perioden:
            wert = layout.wert_zelle(m, p)
            # Split ausschließlich über die Klasse-Zelle des Mastersheets:
            op = ws.cell(r, op_col[p], f'=IF({kz}="OWC",{wert},IF({kz}="TWC",{wert},0))')
            nd = ws.cell(r, nd_col[p], f'=IF({kz}="ND",{wert},0)')
            for cell in (op, nd):
                cell.number_format = ZAHLENFORMAT; cell.font = _normal
        r += 1
    letzte = r - 1
    tc = ws.cell(r, 3, "Summe (operating -> WC-Lead / thereof ND -> ND-Lead)"); tc.font = _bold
    operating, thereof_nd = {}, {}
    for p in perioden:
        for col, ziel in ((op_col[p], operating), (nd_col[p], thereof_nd)):
            L = get_column_letter(col)
            cell = ws.cell(r, col, f"=SUM({L}{erste}:{L}{letzte})")
            cell.number_format = ZAHLENFORMAT; cell.font = _bold; cell.fill = _sub_fill
            cell.border = _top_double
            ziel[p] = f"'{a.sheetname}'!${L}${r}"
    _breiten(ws, {2: 12, 3: 40, 4: 8})
    for i in range(len(perioden) * 2):
        ws.column_dimensions[get_column_letter(p0 + i)].width = 14
    if a.ist_leer:
        ws.sheet_state = "hidden"
    return AufrissRef(a.sheetname, True, {}, operating, thereof_nd)


# ---- Mastersheet (Single Source of Truth) --------------------------------
def _schreibe_mastersheet(wb, mapped, perioden, entity) -> MastersheetLayout:
    ws = wb.create_sheet("Mastersheet")
    kopf = ["Konto", "Bezeichnung", "Entity", "HGB-Pfad (DE)", "HGB-Path (EN)",
            "Klasse", "NA-Zeile (DE)", "NA-Zeile (EN)", "Quelle/Regel", "Review"]
    perioden_start = len(kopf) + 1
    for p in perioden:
        kopf.append(p)
    _schreibe_kopf(ws, kopf, zeile=1)

    r = 2
    zeile_je_account: dict[int, int] = {}
    for m in sorted(mapped, key=lambda x: x.konto):
        zeile_je_account[id(m)] = r
        ws.cell(r, 1, m.konto).font = _normal
        ws.cell(r, 2, m.bezeichnung).font = _normal
        ws.cell(r, 3, m.account.entity).font = _normal
        ws.cell(r, 4, m.hgb_pfad).font = _normal
        ws.cell(r, 5, m.hgb_pfad_en).font = _normal
        c6 = ws.cell(r, 6, m.klasse.value); c6.font = _bold
        c6.fill = _klasse_fill(m.klasse)
        ws.cell(r, 7, m.na_de).font = _normal
        ws.cell(r, 8, m.na_en).font = _normal
        quelle = m.quelle.value + (f" [{m.regel_id}]" if m.regel_id else "")
        ws.cell(r, 9, quelle).font = _normal
        rv = ws.cell(r, 10, "REVIEW" if m.review else ""); rv.font = _bold
        for i, p in enumerate(perioden):
            cell = ws.cell(r, perioden_start + i, round(m.saldo(p), 2))
            cell.number_format = ZAHLENFORMAT
            cell.font = _normal
        r += 1
    letzte = r - 1

    perioden_spalten = {p: perioden_start + i for i, p in enumerate(perioden)}
    _breiten(ws, {1: 10, 2: 34, 3: 18, 4: 46, 5: 46, 6: 8, 7: 30, 8: 30, 9: 26, 10: 9})
    ws.freeze_panes = "A2"
    return MastersheetLayout(
        sheetname="Mastersheet", erste_datenzeile=2, letzte_datenzeile=max(letzte, 2),
        spalte_klasse=6, spalte_na=7, perioden_spalten=perioden_spalten,
        zeile_je_account=zeile_je_account,
    )


# ---- Net-Debt-Tab (zieht aus Aufrissen, Kontrollzeile prüft Mastersheet) --
def _schreibe_net_debt(wb, nd: NetDebtView, layout: MastersheetLayout,
                       perioden, entity, refs: dict[str, AufrissRef]) -> None:
    ws = wb.create_sheet("Net Debt")
    ws.sheet_view.showGridLines = False
    p0 = 4  # erste Perioden-Spalte (nach Ref./NA-DE/NA-EN)

    # Titelblock
    t = ws.cell(1, 2, f"Net Debt — {entity}"); t.font = Font(name=FONT_NAME, bold=True, size=13, color=TEAL)
    ws.cell(2, 2, "in EUR · jede Zeile zieht aus genau einem Aufriss").font = Font(
        name=FONT_NAME, italic=True, size=9)

    kopf_zeile = 4
    ws.cell(kopf_zeile, 2, "Ref."); ws.cell(kopf_zeile, 3, "Net-Asset-Position")
    for i, p in enumerate(perioden):
        ws.cell(kopf_zeile, p0 + i, p)
    _style_kopf_row(ws, kopf_zeile, 2, p0 + len(perioden) - 1)

    r = kopf_zeile + 1
    ref = 1
    daten_zeilen: list[int] = []

    # Die SUMIFS-Formel filtert auf (Klasse=ND, NA-Zeile). Damit keine NA-Zeile
    # doppelt summiert wird, muss jede NA-Zeile im ganzen Tab genau einmal
    # erscheinen — direkte und Umgliederungs-Zeilen müssen disjunkt sein.
    direkt_namen = {z.na_de for z in nd.direkt}
    umg_namen = {z.na_de for z in nd.umgliederung}
    kollision = direkt_namen & umg_namen
    if kollision:
        raise ValueError(
            "Net-Debt-Export: NA-Zeile(n) in direkter UND Umgliederungs-Gruppe "
            f"({sorted(kollision)}) — SUMIFS würde doppelt zählen. Prüfe die "
            "Reklassifizierung/aus_mixed-Zuordnung."
        )

    def schreibe_gruppe(titel: str, zeilen) -> None:
        nonlocal r, ref
        if not zeilen:
            return
        c = ws.cell(r, 3, titel); c.font = _bold; c.fill = _grau_fill
        r += 1
        for z in zeilen:
            ws.cell(r, 2, ref).font = _normal
            aufriss = refs.get(z.na_de)
            quelle = f"  [{aufriss.sheetname}]" if aufriss else ""
            ws.cell(r, 3, f"{z.na_de} / {z.na_en}{quelle}").font = _normal
            for i, p in enumerate(perioden):
                col = p0 + i
                # Jede Zeile zieht aus GENAU EINEM Aufriss: gemischte Position ->
                # thereof-ND-Summe, sonst -> Aufriss-Summe. Kein direkter
                # Mastersheet-Zugriff mehr (nur die Kontrollzeile prüft dagegen).
                cell = ws.cell(r, col)
                cell.value = _ref_formel(aufriss, p, nd_teil=True)
                cell.number_format = ZAHLENFORMAT
                cell.font = _normal
            daten_zeilen.append(r)
            ref += 1
            r += 1

    schreibe_gruppe("Direkte Net-Debt-Positionen", nd.direkt)
    schreibe_gruppe("Umgliederung aus NWC in ND (thereof ND)", nd.umgliederung)

    # Zwischensumme Net Debt
    sub_zeile = r
    ws.cell(sub_zeile, 3, "Netto-Finanzvermögen / -Verbindlichkeiten (Net cash / Net debt)").font = _bold
    for i, p in enumerate(perioden):
        col = p0 + i
        if daten_zeilen:
            spalte = get_column_letter(col)
            formel = "=" + "+".join(f"{spalte}{z}" for z in daten_zeilen)
        else:
            formel = 0
        cell = ws.cell(sub_zeile, col, formel)
        cell.number_format = ZAHLENFORMAT
        cell.font = _bold
        cell.fill = _sub_fill
        cell.border = _top_double
    ws.cell(sub_zeile, 2).fill = _sub_fill

    # Kontrollzeile: unabhängig ALLE ND-Konten aus dem Mastersheet summieren und
    # gegen die aus den View-Zeilen gebaute Zwischensumme stellen. Differenz != 0
    # deckt auf, dass eine ND-Position in keiner View-Zeile repräsentiert ist.
    k = sub_zeile + 2
    ws.cell(k, 3, "Kontrollzeile (Σ ND Mastersheet − Zwischensumme, muss 0 sein)").font = Font(
        name=FONT_NAME, italic=True, size=9)
    for i, p in enumerate(perioden):
        col = p0 + i
        spalte = get_column_letter(col)
        gesamt_nd = (
            f"SUMIFS({layout.bereich(layout.perioden_spalten[p])},"
            f"{layout.bereich(layout.spalte_klasse)},\"ND\")"
        )
        cell = ws.cell(k, col, f"={gesamt_nd}-{spalte}{sub_zeile}")
        cell.number_format = ZAHLENFORMAT
        cell.font = Font(name=FONT_NAME, italic=True, size=9)

    _breiten(ws, {2: 6, 3: 58})
    for i in range(len(perioden)):
        ws.column_dimensions[get_column_letter(p0 + i)].width = 15
    ws.freeze_panes = ws.cell(kopf_zeile + 1, p0)


# ---- Working-Capital-Tab (zieht aus Aufrissen, Kontrollzeile prüft MS) ----
def _schreibe_working_capital(wb, wc: WCView, layout: MastersheetLayout,
                              perioden, entity, refs: dict[str, AufrissRef]) -> None:
    ws = wb.create_sheet("Working Capital")
    ws.sheet_view.showGridLines = False
    p0 = 4

    ws.cell(1, 2, f"Working Capital (Ist je Periode) — {entity}").font = Font(
        name=FONT_NAME, bold=True, size=13, color=TEAL)
    ws.cell(2, 2, "in EUR · jede Zeile zieht aus genau einem Aufriss · WC-Definition "
                  "über alle Perioden identisch · noch keine normalisierte Referenz").font = Font(
        name=FONT_NAME, italic=True, size=9)

    kopf_zeile = 4
    ws.cell(kopf_zeile, 2, "Ref."); ws.cell(kopf_zeile, 3, "Net-Asset-Position")
    for i, p in enumerate(perioden):
        ws.cell(kopf_zeile, p0 + i, p)
    _style_kopf_row(ws, kopf_zeile, 2, p0 + len(perioden) - 1)

    r = kopf_zeile + 1
    ref = 1
    zeilen_rows: dict[tuple[str, str], list[int]] = {}   # (seite,klasse) -> rows

    def summe_zeile(label: str, rows: list[int], fett=True, fill=None) -> int:
        nonlocal r
        c = ws.cell(r, 3, label); c.font = _bold if fett else _normal
        if fill:
            c.fill = fill
        for i, p in enumerate(perioden):
            col = p0 + i
            sp = get_column_letter(col)
            formel = ("=" + "+".join(f"{sp}{z}" for z in rows)) if rows else 0
            cell = ws.cell(r, col, formel)
            cell.number_format = ZAHLENFORMAT
            cell.font = _bold if fett else _normal
            if fill:
                cell.fill = fill
        zr = r
        r += 1
        return zr

    def klasse_block(seite: str, klasse: str) -> None:
        nonlocal r, ref
        zeilen = wc.zeilen_fuer(seite, klasse)
        if not zeilen:
            return
        label = "Trade Working Capital (TWC)" if klasse == "TWC" else "Other Working Capital (OWC)"
        h = ws.cell(r, 3, label); h.font = _bold; h.fill = _grau_fill
        r += 1
        rows: list[int] = []
        for z in zeilen:
            ws.cell(r, 2, ref).font = _normal
            aufriss = refs.get(z.na_de)
            quelle = f"  [{aufriss.sheetname}]" if aufriss else ""
            ws.cell(r, 3, f"{z.na_de} / {z.na_en}{quelle}").font = _normal
            for i, p in enumerate(perioden):
                # Jede Zeile zieht aus GENAU EINEM Aufriss: gemischte Position ->
                # operating-Summe, sonst -> Aufriss-Summe.
                cell = ws.cell(r, p0 + i, _ref_formel(aufriss, p, nd_teil=False))
                cell.number_format = ZAHLENFORMAT
                cell.font = _normal
            rows.append(r)
            zeilen_rows.setdefault((seite, klasse), []).append(r)
            ref += 1
            r += 1
        summe_zeile(f"  davon {label}", rows, fett=True)

    def seiten_block(seite: str, titel: str) -> int:
        nonlocal r
        t = ws.cell(r, 3, titel); t.font = Font(name=FONT_NAME, bold=True, color=TEAL)
        r += 1
        klasse_block(seite, "TWC")
        klasse_block(seite, "OWC")
        alle_rows = zeilen_rows.get((seite, "TWC"), []) + zeilen_rows.get((seite, "OWC"), [])
        return summe_zeile(titel + " gesamt", alle_rows, fett=True, fill=_sub_fill)

    oa_row = seiten_block("OA", "Operating Assets")
    ol_row = seiten_block("OL", "Operating Liabilities")

    # Net Working Capital = Operating Assets + Operating Liabilities (Passiva
    # sind vorzeichenrichtig negativ gespeichert => OA − |OL|).
    nwc_row = r
    ws.cell(nwc_row, 3, "Net Working Capital (Operating Assets − Operating Liabilities)").font = _bold
    for i, p in enumerate(perioden):
        col = p0 + i
        sp = get_column_letter(col)
        cell = ws.cell(nwc_row, col, f"={sp}{oa_row}+{sp}{ol_row}")
        cell.number_format = ZAHLENFORMAT
        cell.font = _bold
        cell.fill = _sub_fill
        cell.border = _top_double
    r += 2

    # Kontrollzeile: alle TWC+OWC im Mastersheet minus NWC = 0.
    ws.cell(r, 3, "Kontrollzeile (Σ TWC+OWC Mastersheet − NWC, muss 0 sein)").font = Font(
        name=FONT_NAME, italic=True, size=9)
    for i, p in enumerate(perioden):
        col = p0 + i
        sp = get_column_letter(col)
        rng = layout.bereich(layout.perioden_spalten[p])
        kl = layout.bereich(layout.spalte_klasse)
        gesamt = f"SUMIFS({rng},{kl},\"TWC\")+SUMIFS({rng},{kl},\"OWC\")"
        cell = ws.cell(r, col, f"={gesamt}-{sp}{nwc_row}")
        cell.number_format = ZAHLENFORMAT
        cell.font = Font(name=FONT_NAME, italic=True, size=9)

    _breiten(ws, {2: 6, 3: 58})
    for i in range(len(perioden)):
        ws.column_dimensions[get_column_letter(p0 + i)].width = 15
    ws.freeze_panes = ws.cell(kopf_zeile + 1, p0)


# ---- Review-Queue ---------------------------------------------------------
def _schreibe_review(wb, review: list[ReviewEintrag], perioden) -> None:
    ws = wb.create_sheet("Review-Queue")
    kopf = ["Konto", "Bezeichnung", "HGB-Pfad", "Klasse", "Status",
            "Quelle/Regel", "Grund"]
    p_start = len(kopf) + 1
    kopf += list(perioden)
    _schreibe_kopf(ws, kopf, zeile=1)
    r = 2
    for e in review:
        ws.cell(r, 1, e.konto).font = _normal
        ws.cell(r, 2, e.bezeichnung).font = _normal
        ws.cell(r, 3, e.hgb_pfad).font = _normal
        ws.cell(r, 4, e.klasse).font = _bold
        st = ws.cell(r, 5, e.status); st.font = _bold
        if e.status.startswith("Pflichtfrage"):
            st.fill = PatternFill("solid", fgColor="F9CB9C")
        elif "Verhaltensprüfung" in e.status:
            st.fill = PatternFill("solid", fgColor="FFF2CC")
        ws.cell(r, 6, (e.quelle + (f" [{e.regel_id}]" if e.regel_id else ""))).font = _normal
        ws.cell(r, 7, e.grund).font = _normal
        for i, p in enumerate(perioden):
            cell = ws.cell(r, p_start + i, round(e.salden.get(p, 0.0), 2))
            cell.number_format = ZAHLENFORMAT
            cell.font = _normal
        r += 1
    if not review:
        ws.cell(2, 1, "(keine offenen Fälle)").font = Font(name=FONT_NAME, italic=True)
    _breiten(ws, {1: 10, 2: 40, 3: 44, 4: 8, 5: 26, 6: 24, 7: 52})
    ws.freeze_panes = "A2"


def _schreibe_info(wb, meta: dict) -> None:
    ws = wb.create_sheet("Info")
    ws.cell(1, 1, "Lauf-Metadaten (Reproduzierbarkeit)").font = Font(
        name=FONT_NAME, bold=True, size=12, color=TEAL)
    r = 3
    for k, v in meta.items():
        ws.cell(r, 1, str(k)).font = _bold
        ws.cell(r, 2, str(v)).font = _normal
        r += 1
    _breiten(ws, {1: 28, 2: 70})


# ---- Stil-Helfer ----------------------------------------------------------
def _schreibe_kopf(ws, kopf, zeile) -> None:
    for c, text in enumerate(kopf, start=1):
        cell = ws.cell(zeile, c, text)
        cell.font = _kopf_font
        cell.fill = _kopf_fill
        cell.alignment = Alignment(vertical="center")


def _style_kopf_row(ws, zeile, c_von, c_bis) -> None:
    for c in range(c_von, c_bis + 1):
        cell = ws.cell(zeile, c)
        cell.font = _kopf_font
        cell.fill = _kopf_fill


def _klasse_fill(klasse: Klasse) -> PatternFill:
    farben = {
        Klasse.ND: "F4CCCC", Klasse.TWC: "D9EAD3", Klasse.OWC: "FFF2CC",
        Klasse.FA: "CFE2F3", Klasse.EQ: "D9D2E9", Klasse.DT: "EAD1DC",
        Klasse.REVIEW: "F9CB9C", Klasse.TECH: "EFEFEF", Klasse.PL: "FFFFFF",
        Klasse.MIXED: "FCE5CD",
    }
    return PatternFill("solid", fgColor=farben.get(klasse, "FFFFFF"))


def _breiten(ws, mapping: dict[int, int]) -> None:
    for col, w in mapping.items():
        ws.column_dimensions[get_column_letter(col)].width = w
