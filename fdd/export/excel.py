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
    """Merkt sich, wo im Mastersheet welche Spalte/Zeilen liegen, damit die
    Net-Debt-Formeln korrekt referenzieren."""

    sheetname: str
    erste_datenzeile: int
    letzte_datenzeile: int
    spalte_klasse: int
    spalte_na: int
    perioden_spalten: dict[str, int]

    def bereich(self, spalte: int) -> str:
        c = get_column_letter(spalte)
        return f"'{self.sheetname}'!${c}${self.erste_datenzeile}:${c}${self.letzte_datenzeile}"


def _krit(text: str) -> str:
    """Escaping für ein SUMIFS-Text-Kriterium: Doppelquotes verdoppeln (sonst
    bricht der Formel-String), Excel-Wildcards (* ? ~) mit ~ neutralisieren
    (sonst ändern sie stillschweigend die Match-Semantik)."""
    text = text.replace('"', '""')
    for w in ("~", "*", "?"):
        text = text.replace(w, "~" + w)
    return text


def schreibe_databook(pfad: str, mapped: list[MappedAccount], nd: NetDebtView,
                      review: list[ReviewEintrag], perioden: list[str],
                      entity: str, meta: Optional[dict] = None) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    layout = _schreibe_mastersheet(wb, mapped, perioden, entity)
    _schreibe_net_debt(wb, nd, layout, perioden, entity)
    _schreibe_review(wb, review, perioden)
    if meta:
        _schreibe_info(wb, meta)
    wb.save(pfad)


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
    for m in sorted(mapped, key=lambda x: x.konto):
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
    )


# ---- Net-Debt-Tab (sichtbare SUMIFS-Formeln, Kontrollzeile) --------------
def _schreibe_net_debt(wb, nd: NetDebtView, layout: MastersheetLayout,
                       perioden, entity) -> None:
    ws = wb.create_sheet("Net Debt")
    ws.sheet_view.showGridLines = False
    p0 = 4  # erste Perioden-Spalte (nach Ref./NA-DE/NA-EN)

    # Titelblock
    t = ws.cell(1, 2, f"Net Debt — {entity}"); t.font = Font(name=FONT_NAME, bold=True, size=13, color=TEAL)
    ws.cell(2, 2, "in EUR").font = Font(name=FONT_NAME, italic=True, size=9)

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
            ws.cell(r, 3, f"{z.na_de} / {z.na_en}").font = _normal
            for i, p in enumerate(perioden):
                col = p0 + i
                # Sichtbare SUMIFS-Formel auf das Mastersheet (Single Source of Truth).
                # Summiert alle ND-Konten der Klasse-Spalte mit passender NA-Zeile.
                cell = ws.cell(r, col)
                cell.value = (
                    f"=SUMIFS({layout.bereich(layout.perioden_spalten[p])},"
                    f"{layout.bereich(layout.spalte_klasse)},\"ND\","
                    f"{layout.bereich(layout.spalte_na)},\"{_krit(z.na_de)}\")"
                )
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


# ---- Review-Queue ---------------------------------------------------------
def _schreibe_review(wb, review: list[ReviewEintrag], perioden) -> None:
    ws = wb.create_sheet("Review-Queue")
    kopf = ["Konto", "Bezeichnung", "HGB-Pfad", "Klasse", "Quelle/Regel", "Grund"]
    p_start = len(kopf) + 1
    kopf += list(perioden)
    _schreibe_kopf(ws, kopf, zeile=1)
    r = 2
    for e in review:
        ws.cell(r, 1, e.konto).font = _normal
        ws.cell(r, 2, e.bezeichnung).font = _normal
        ws.cell(r, 3, e.hgb_pfad).font = _normal
        ws.cell(r, 4, e.klasse).font = _bold
        ws.cell(r, 5, (e.quelle + (f" [{e.regel_id}]" if e.regel_id else ""))).font = _normal
        ws.cell(r, 6, e.grund).font = _normal
        for i, p in enumerate(perioden):
            cell = ws.cell(r, p_start + i, round(e.salden.get(p, 0.0), 2))
            cell.number_format = ZAHLENFORMAT
            cell.font = _normal
        r += 1
    if not review:
        ws.cell(2, 1, "(keine offenen Fälle)").font = Font(name=FONT_NAME, italic=True)
    _breiten(ws, {1: 10, 2: 40, 3: 46, 4: 8, 5: 26, 6: 60})
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
