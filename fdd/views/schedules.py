"""Aufriss-Schicht (Schedules) — Architektur-Spec Abschnitt 3–5.

Sitzt zwischen Mastersheet (Single Source of Truth) und den Leads. Jede
Net-Asset-Zeile bekommt genau einen Aufriss; kein Lead zieht mehr direkt aus
dem Mastersheet, sondern jede Lead-Zeile holt ihren Wert aus genau einem
Aufriss (Uniformität — der Rödl-Konstruktionsfehler wird so vermieden).

Die drei gemischten Positionen (NA_OA/NA_OL/NA_OP) führen zwei Wertspalten:
`operating` (Klasse OWC) und `thereof ND` (Klasse ND). Der WC-Lead zieht die
operating-Spalte, der ND-Lead die thereof-ND-Spalte — beide aus demselben
Aufriss, die Aufteilung ist damit an genau einer Stelle definiert.

Auch die "trivialen" Aufrisse werden gebaut (Cash, Pensions, Vorräte, …); leere
Aufrisse werden beim Export ausgeblendet.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from ..core.model import Klasse, MappedAccount

# Die drei gemischten Positionen -> Haus-Aufrissname.
MIXED_NA: dict[str, str] = {
    "Sonstige Vermoegensgegenstaende": "NA_OA",
    "Sonstige Verbindlichkeiten": "NA_OL",
    "Sonstige Rueckstellungen": "NA_OP",
}

# Feste, kurze Aufriss-Sheetnamen je Net-Asset-Zeile (Excel-Namen <= 31 Zeichen).
_SHEETNAME: dict[str, str] = {
    # ND-Positionen
    "Liquide Mittel": "A_Cash",
    "Wertpapiere": "A_Wertpapiere",
    "Ausleihungen": "A_Ausleihungen",
    "Anleihen": "A_Anleihen",
    "Verbindlichkeiten ggue. Kreditinstituten": "A_VbKreditinstitute",
    "Pensionsrueckstellungen": "A_Pensionsrst",
    "Steuerrueckstellungen": "A_Steuerrst",
    "Forderungen ggue. verbundenen Unternehmen": "A_FordVerbUnt",
    "Verbindlichkeiten ggue. verbundenen Unternehmen": "A_VbVerbUnt",
    # WC-Positionen
    "Forderungen aus L+L": "A_FordLuL",
    "Vorraete": "A_Vorraete",
    "Geleistete Anzahlungen": "A_GelAnzahlungen",
    "Erhaltene Anzahlungen": "A_ErhAnzahlungen",
    "Verbindlichkeiten aus L+L": "A_VbLuL",
    "Aktive Rechnungsabgrenzung": "A_ARAP",
    "Passive Rechnungsabgrenzung": "A_PRAP",
    # gemischt
    "Sonstige Vermoegensgegenstaende": "NA_OA",
    "Sonstige Verbindlichkeiten": "NA_OL",
    "Sonstige Rueckstellungen": "NA_OP",
}

# Diese Klassen speisen ND- bzw. WC-Lead und bekommen daher einen Aufriss.
_LEAD_KLASSEN = (Klasse.ND, Klasse.TWC, Klasse.OWC)


@dataclass
class Aufriss:
    sheetname: str
    na_de: str
    na_en: str
    is_mixed: bool
    speist: str                      # "ND" | "WC" | "beide"
    konten: list[MappedAccount]
    perioden: list[str]

    def operating_summe(self, p: str) -> float:
        return sum(m.saldo(p) for m in self.konten if m.klasse in (Klasse.TWC, Klasse.OWC))

    def thereof_nd_summe(self, p: str) -> float:
        return sum(m.saldo(p) for m in self.konten if m.klasse == Klasse.ND)

    def summe(self, p: str) -> float:
        return sum(m.saldo(p) for m in self.konten)

    @property
    def ist_leer(self) -> bool:
        return all(abs(self.summe(p)) < 0.005 for p in self.perioden)


@dataclass
class Schedules:
    aufrisse: list[Aufriss]
    perioden: list[str]
    # Konten mit Lead-Klasse (ND/TWC/OWC), die in keinem Aufriss landen:
    ohne_aufriss: list[MappedAccount] = field(default_factory=list)

    def by_na(self, na_de: str) -> Aufriss | None:
        for a in self.aufrisse:
            if a.na_de == na_de:
                return a
        return None


def _sheetname(na_de: str) -> str:
    if na_de in _SHEETNAME:
        return _SHEETNAME[na_de]
    # Fallback: sicherer Slug
    slug = "".join(ch for ch in na_de if ch.isalnum() or ch == " ").strip()
    slug = slug.replace(" ", "")[:26]
    return "A_" + (slug or "Position")


def _speist(is_mixed: bool, konten: list[MappedAccount]) -> str:
    if is_mixed:
        return "beide"
    kl = konten[0].klasse
    return "ND" if kl == Klasse.ND else "WC"


def baue_schedules(mapped: list[MappedAccount], perioden: list[str]) -> Schedules:
    gruppen: dict[str, list[MappedAccount]] = {}
    ohne: list[MappedAccount] = []

    for m in mapped:
        if m.klasse not in _LEAD_KLASSEN:
            continue
        if not m.na_de or m.na_de.startswith("("):
            ohne.append(m)
            continue
        gruppen.setdefault(m.na_de, []).append(m)

    aufrisse: list[Aufriss] = []
    for na_de, konten in gruppen.items():
        is_mixed = na_de in MIXED_NA
        aufrisse.append(Aufriss(
            sheetname=_sheetname(na_de), na_de=na_de, na_en=konten[0].na_en,
            is_mixed=is_mixed, speist=_speist(is_mixed, konten),
            konten=sorted(konten, key=lambda m: m.konto), perioden=list(perioden),
        ))

    # Anzeigereihenfolge: ND-Aufrisse, dann WC, dann gemischte; je nach Namen
    reihenfolge = {"ND": 0, "WC": 1, "beide": 2}
    aufrisse.sort(key=lambda a: (reihenfolge[a.speist], a.sheetname))
    return Schedules(aufrisse=aufrisse, perioden=list(perioden), ohne_aufriss=ohne)
