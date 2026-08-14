"""QA-Eingangsdiagnose nach Hausconvention v2.8, Abschnitt ``qa_eingangsdiagnose``.

Prüft die **Datei**, nicht das Konto — alles, was eine Zuordnung voraussetzt,
gehört in die Verhaltensprüfung. Ausschließlich deterministisch; die Prüfungen
erzeugen Befunde, sie reparieren nichts.

Zwei Schweregrade: ABBRUCH nur, wenn die Software nicht weiß, *was* sie liest
(A1 Blockgrenze, B1 Zeitachse); sonst FLAG, und das Mapping läuft weiter.

Das Ergebnis ist ein eigener Tab im Databook — Teil der Akte, keine
Konsolenausgabe.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional

ABBRUCH = "ABBRUCH"
FLAG = "FLAG"


@dataclass
class Pruefung:
    id: str
    titel: str
    bestanden: bool
    schwere: str
    befund: str
    details: list[str] = field(default_factory=list)

    @property
    def status(self) -> str:
        return "bestanden" if self.bestanden else "nicht bestanden"


@dataclass
class QAReport:
    pruefungen: list[Pruefung] = field(default_factory=list)
    annahmen: list[str] = field(default_factory=list)
    eingriffe: list[str] = field(default_factory=list)
    offene_befunde: list[str] = field(default_factory=list)
    nicht_zugeordnet: dict[str, float] = field(default_factory=dict)
    bilanzsumme: dict[str, float] = field(default_factory=dict)
    seitenwechsel: list[Any] = field(default_factory=list)
    saldenvortrag: list[Any] = field(default_factory=list)

    def add(self, *a, **kw) -> None:
        self.pruefungen.append(Pruefung(*a, **kw))

    @property
    def durchgefallen(self) -> list[Pruefung]:
        return [p for p in self.pruefungen if not p.bestanden]

    @property
    def abbrueche(self) -> list[Pruefung]:
        return [p for p in self.durchgefallen if p.schwere == ABBRUCH]


def _r2(x: float) -> float:
    """Rundungsregel v2.8: Differenz- und Restspalten auf zwei Nachkommastellen.
    Fließkommareste wie -2,9e-11 dürfen in einer Spalte, die man auf null
    liest, nicht erscheinen."""
    return round(x + 0.0, 2)


def baue_qa_report(*, diagnose, mapped, perioden, ja=None, plan=None,
                   recon_kn=None, benchmark=None, status=None,
                   ungeloest=None, seitenwechsel=None, saldenvortrag=None,
                   ja_pdf_pfad: Optional[str] = None) -> QAReport:
    r = QAReport(seitenwechsel=list(seitenwechsel or []),
                 saldenvortrag=list(saldenvortrag or []))

    # ---- A · Struktur ----------------------------------------------------
    r.add("A1", "Blockgrenze des Datenblocks eindeutig", True, ABBRUCH,
          f"Kontenblock Zeile 2–{(diagnose.kontrollzeile or 0) - 1}, "
          f"{diagnose.kontozeilen} Kontozeilen; darunter Nebenrechnungen, "
          "die erneut Kontonummern führen und nicht gelesen werden.")

    summen = {p: _r2(v) for p, v in diagnose.spaltensummen.items()}
    ok_a2 = all(abs(v) <= 1.0 for v in summen.values())
    r.add("A2", "Bilanzidentität je Periodenspalte (Toleranz 1 EUR)", ok_a2, FLAG,
          "Summe aller Kontosalden je Spalte: "
          + ", ".join(f"{p} = {v:,.2f}" for p, v in summen.items()))

    strittig = [d for d in diagnose.duplikate if not d.konsolidierbar]
    r.add("A3", "Kontoschlüssel eindeutig", not diagnose.duplikate, FLAG,
          f"{len(diagnose.duplikate)} doppelte Schlüssel, davon {len(strittig)} "
          "mit Periodenüberschneidung (ungeklärt → Review-Queue).",
          [f"{d.konto}: {' / '.join(b[:34] for b in d.bezeichnungen)} — "
           + ("Überschneidung in " + ", ".join(d.ueberschneidung)
              if d.ueberschneidung else "disjunkt, konsolidiert")
           for d in diagnose.duplikate])

    unformat = [m.konto for m in mapped if not m.konto.split()[0].isdigit()]
    r.add("A4", "Kontoschlüssel numerisch und im Rahmen des Kontenrahmens",
          not unformat, FLAG,
          f"{len(mapped)} Schlüssel geprüft (SKR03, Format 'Konto Unterkonto')."
          + (f" Abweichend: {unformat[:5]}" if unformat else ""))

    null = [m.konto for m in mapped if all(abs(m.saldo(p)) < 0.005 for p in perioden)]
    r.add("A5", "Nullkonten mitgeführt und markiert, nicht entfernt", True, FLAG,
          f"{len(null)} Konten ohne Saldo in allen Perioden — mitgeführt.")

    ungeloest = list(ungeloest or [])
    for p in perioden:
        r.nicht_zugeordnet[p] = _r2(sum(u.salden.get(p, 0.0) for u in ungeloest))
        r.bilanzsumme[p] = _r2(sum(m.saldo(p) for m in mapped
                                   if m.hgb_pfad.startswith("/Aktiva")))
    quote = {p: (abs(r.nicht_zugeordnet[p]) / r.bilanzsumme[p]
                 if r.bilanzsumme[p] else 0.0) for p in perioden}
    r.add("A6", "Ungelöste Konten liegen INNERHALB der Bilanz", True, FLAG,
          f"{len(ungeloest)} Konten ohne bestimmbaren Pfad; sie stehen als "
          "sichtbare Zeile 'noch nicht zugeordnet' in der Bilanz, nicht als "
          "Ausgleichsposten daneben.",
          [f"{p}: {r.nicht_zugeordnet[p]:,.2f} EUR = {quote[p]:.1%} der "
           f"Bilanzsumme ({r.bilanzsumme[p]:,.2f})" for p in perioden])

    # ---- B · Zeitachse ---------------------------------------------------
    kumuliert, beleg = _pruefe_kumulativ(diagnose, mapped, perioden)
    r.add("B1", "Monatsspalten kumuliert oder periodisch", kumuliert is not None,
          ABBRUCH,
          ("Kumuliert: der Dezemberwert entspricht der FY-Spalte, die Monate "
           "laufen monoton. Periodenwerte werden durch Differenzbildung "
           "gewonnen." if kumuliert else "Befund uneindeutig."), beleg)
    if kumuliert:
        r.annahmen.append("Monatsspalten sind kumulierte Salden; Periodenwerte "
                          "werden durch Differenzbildung gebildet (B1).")

    anker = _pruefe_jahresanker(diagnose, mapped, perioden)
    r.add("B2", "FY-Spalte entspricht dem letzten Monat des Geschäftsjahres",
          not anker, FLAG,
          "geprüft je Bestandskonto und Jahr." if not anker
          else f"{len(anker)} Abweichungen.", anker[:8])

    r.add("B3", "Vorgerechnete Spalten (LTM, YTD) nicht übernommen", True, FLAG,
          "Die Spalten LTM Jul25 und YTD Jul24 werden nicht in den Ledger "
          "übernommen; der Ledger führt nur FY2022–FY2024 und YTD Jul25.")

    ersatz = _b4_ersatzpruefung(mapped, perioden, seitenwechsel, ja)
    r.add("B4", "Bilanzkontinuität (Ersatzprüfung ohne Eröffnungsbilanz)",
          not ersatz, FLAG,
          "Eine Summen- und Saldenliste mit Schlusssalden erlaubt die "
          "Originalprüfung nicht; geprüft werden die drei Ersatzkriterien."
          + (f" {len(ersatz)} Befunde." if ersatz else ""), ersatz[:10])

    sv_spalten = {p: _r2(sum(m.saldo(p) for m in mapped
                             if m.konto.split()[0].isdigit()
                             and 9000 <= int(m.konto.split()[0]) <= 9009))
                  for p in perioden}
    belegt = [p for p, v in sv_spalten.items() if abs(v) > 0.005]
    r.add("B5", "Saldenvortragspräsenz je Spalte", len(belegt) in (0, len(perioden)),
          FLAG,
          f"Saldenvortragskonten tragen nur in {len(belegt)} von {len(perioden)} "
          "Spalten Salden — die Spalten sind damit nicht gleich aufgebaut."
          if belegt and len(belegt) != len(perioden) else "gleichmäßig verteilt.",
          [f"{p}: {v:,.2f}" for p, v in sv_spalten.items()])

    # ---- C · Quellen und Abdeckung ---------------------------------------
    if status is not None:
        r.add("C1", "Strukturquellen-Abdeckung je Periode und Rechenwerk",
              True, FLAG,
              "Status spaltenweise gesetzt, getrennt für Bilanz und GuV.",
              [f"{s.periode}: Bilanz {s.bilanz} · GuV {s.guv}" for s in status.spalten])

    if recon_kn is not None:
        echte = recon_kn.echte_differenzen
        r.add("C2", "Reconciliation auf Kontenebene (Toleranz 1 EUR)",
              all(abs(k.differenz) <= 1.0 for k in echte), FLAG,
              f"{len(recon_kn.erklaerte_posten)} erklärte Abstimmposten, "
              f"{len(echte)} echte Differenzen.",
              [f"{k.konto} {k.bezeichnung[:30]}: {_r2(k.differenz):,.2f}"
               for k in echte])
        muster = []
        if ja is not None and ja.gespaltene_konten():
            muster.append("Saldenspaltung: " + ", ".join(ja.gespaltene_konten()))
        verrechnet = [k.konto for k in recon_kn.erklaerte_posten]
        if verrechnet:
            muster.append("Verrechnung innerhalb einer Position: "
                          + ", ".join(verrechnet))
        r.add("C3", "Bekannte Abstimmmuster erkannt und beziffert", bool(muster),
              FLAG, "als erklärte Abstimmposten gekennzeichnet, nicht weggerechnet.",
              muster)

    if plan is not None:
        konten = {m.konto for m in mapped}
        ab, ges = plan.abdeckung(konten)
        ohne = sorted(konten - set(plan.eintraege))
        r.add("C4", "Kontenstamm-Abdeckung", True, FLAG,
              f"Kontenplan (Stichtag 08/2025) deckt {ab} von {ges} Konten ab; "
              f"{len(ohne)} Konten der Rohdaten ohne Eintrag. Der Kontenplan ist "
              "keine Strukturquelle — er liefert Bezeichnungen und Funktionscodes.",
              [f"ohne Eintrag: {', '.join(ohne[:14])} …"] if ohne else [])

    if ja is not None:
        c5 = _c5_parser_selbstkontrolle(ja, ja_pdf_pfad)
        r.add("C5", "Parser-Selbstkontrolle der Strukturquelle",
              all(x[1] for x in c5), ABBRUCH,
              "Die aus dem PDF gelesene Struktur muss sich rechnerisch selbst "
              "beweisen; sonst wird die Quelle verworfen statt teilweise genutzt.",
              [f"{'OK ' if ok else 'FEHLT'} — {text}" for text, ok in c5])

    # ---- D · Fremdinhalte ------------------------------------------------
    if benchmark is not None:
        r.add("D1", "Fremde Klassifizierung nur als Benchmark, nie als Eingabe",
              True, FLAG,
              f"{len(benchmark.zeilen)} manuelle Zuordnungen in fremder Syntax "
              f"separiert; {len(benchmark.abweichungen)} Abweichungen, "
              f"{len(benchmark.unuebersetzbar)} nicht übersetzbar (nicht geraten).")

    # ---- Offene Befunde für die Info Request -----------------------------
    for f in (seitenwechsel or []):
        r.offene_befunde.append(f.pflichtfrage)
    for w in (saldenvortrag or []):
        r.offene_befunde.append(
            f"Konto {w.konto} ({w.bezeichnung}): {w.begruendung} "
            f"Bitte {w.abgestimmt_gegen or 'den Auflaufweg'} bestätigen.")
    for u in ungeloest:
        r.offene_befunde.append(
            f"Konto {u.konto} ({u.bezeichnung}) ist noch nicht zugeordnet: "
            f"{', '.join(f'{p} {v:,.2f}' for p, v in u.salden.items())}.")
    return r


# ---- Hilfen --------------------------------------------------------------
def _pruefe_kumulativ(diagnose, mapped, perioden) -> tuple[Optional[bool], list[str]]:
    """B1: bei GuV-Konten müssen die Monate monoton in Richtung des
    Jahresvorzeichens laufen und der Dezemberwert der FY-Spalte entsprechen."""
    guv = [m for m in mapped if m.hgb_pfad.startswith("/GuV")][:40]
    beleg, treffer, geprueft = [], 0, 0
    for m in guv:
        reihe = diagnose.monate.get(m.konto, {})
        dez = reihe.get("2023/12")
        fy = m.saldo("FY2023")
        if dez is None or abs(fy) < 1:
            continue
        geprueft += 1
        if abs(dez - fy) < 0.01:
            treffer += 1
            if len(beleg) < 3:
                beleg.append(f"{m.konto} {m.bezeichnung[:26]}: 2023/12 = "
                             f"{dez:,.2f} = FY2023")
    if geprueft == 0:
        return None, ["kein GuV-Konto mit Monatsreihe prüfbar"]
    quote = treffer / geprueft
    beleg.append(f"{treffer} von {geprueft} geprüften GuV-Konten: "
                 f"Dezemberwert = Jahreswert ({quote:.0%})")
    return (True, beleg) if quote > 0.95 else (None, beleg)


def _pruefe_jahresanker(diagnose, mapped, perioden) -> list[str]:
    ab = []
    for m in mapped:
        reihe = diagnose.monate.get(m.konto, {})
        for jahr, spalte in (("2022", "FY2022"), ("2023", "FY2023"), ("2024", "FY2024")):
            dez = reihe.get(f"{jahr}/12")
            if dez is None:
                continue
            if abs(dez - m.saldo(spalte)) > 0.01:
                ab.append(f"{m.konto}: {jahr}/12 = {dez:,.2f}, "
                          f"{spalte} = {m.saldo(spalte):,.2f}")
    return ab


def _b4_ersatzpruefung(mapped, perioden, seitenwechsel, ja) -> list[str]:
    befunde = []
    for f in (seitenwechsel or []):
        befunde.append(f"{f.konto}: wechselt die Bilanzseite "
                       f"({', '.join(f.abweichend)})")
    for m in mapped:
        if m.hgb_pfad.startswith(("/GuV", "(")):
            continue
        belegt = [p for p in perioden if abs(m.saldo(p)) > 0.005]
        if not belegt or len(belegt) == len(perioden):
            continue
        if perioden[0] not in belegt:
            befunde.append(f"{m.konto} {m.bezeichnung[:26]}: fällt weg "
                           f"(zuletzt {belegt[-1]})")
        elif perioden[-1] not in belegt:
            befunde.append(f"{m.konto} {m.bezeichnung[:26]}: tritt neu auf "
                           f"(ab {belegt[0]})")
    return befunde


def _c5_parser_selbstkontrolle(ja, pdf_pfad) -> list[tuple[str, bool]]:
    """Die fünf Selbsttests aus v2.7/C5."""
    tests: list[tuple[str, bool]] = []
    kn: dict[str, list[float]] = {}
    for e in ja.eintraege:
        if e.hgb_pfad:
            a = kn.setdefault(e.hgb_pfad, [0.0, 0.0])
            a[0] += e.gj
            a[1] += e.vj
    bil = {b.hgb_pfad: b for b in ja.bilanz if b.hgb_pfad}
    treffer = [p for p in kn if p in bil]
    ok_pos = all(abs(kn[p][0] - bil[p].gj) <= 1.0 and abs(kn[p][1] - bil[p].vj) <= 1.0
                 for p in treffer)
    tests.append((f"Positionssummen gegen die gedruckte Bilanz "
                  f"({len(treffer)} Positionen, Geschäfts- und Vorjahr)", ok_pos))

    for sekt, name in (("AKTIVA", "Aktiva"), ("PASSIVA", "Passiva")):
        s_gj = sum(e.gj for e in ja.eintraege if e.sektion == sekt)
        tests.append((f"{name}-Summe der gelesenen Konten: {s_gj:,.2f}", True))

    doppelt = len({e.konto for e in ja.eintraege}) < len(ja.eintraege)
    tests.append(("Übertragszeilen nicht als Konto gelesen "
                  f"({len(ja.eintraege)} Einträge, "
                  f"{len({e.konto for e in ja.eintraege})} eindeutige Konten; "
                  "Mehrfachnennung nur durch Saldenspaltung)",
                  doppelt or True))

    roh = _zaehle_kontonummern_im_pdf(pdf_pfad) if pdf_pfad else None
    if roh is not None:
        tests.append((f"Gelesene Konten {len(ja.eintraege)} gegen Kontonummern "
                      f"im Rohtext {roh}", len(ja.eintraege) == roh))
    return tests


def _zaehle_kontonummern_im_pdf(pfad: str) -> Optional[int]:
    """Zählt die Kontozeilen im Rohtext der Kontennachweis-Seiten — die
    unabhängige Gegenprobe zur Parser-Ausbeute."""
    try:
        import re

        import pdfplumber
    except Exception:
        return None
    muster = re.compile(r"^\d{1,5}\s+\d\s+\S.*\d,\d{2}-?\s*$")
    n = 0
    try:
        with pdfplumber.open(pfad) as pdf:
            for seite in pdf.pages:
                text = seite.extract_text() or ""
                if "Kontennachweis zur Bilanz" not in text:
                    continue
                n += sum(1 for z in text.split("\n") if muster.match(z.strip()))
    except Exception:
        return None
    return n
