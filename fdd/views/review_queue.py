"""Review-Queue als eigene Ausgabe: die ungelösten und geflaggten Konten.

Ab v2.3 zeigt sie zusätzlich die Regel-Marker als **Status** an
(Pflichtfrage: Aufriss / Pflichtfrage: Pension / Verhaltensprüfung offen). Das
ist reine Anzeige — hinter den Markern läuft in dieser Scheibe keine Logik.
Ziel: die Queue enthält danach vor allem Pflichtfragen (Dinge für den Menschen),
nicht mehr Fälle, die eine Regel hätte lösen können.
"""

from __future__ import annotations

from dataclasses import dataclass

from ..core.model import Klasse, MappedAccount


@dataclass
class ReviewEintrag:
    konto: str
    bezeichnung: str
    hgb_pfad: str
    klasse: str
    status: str          # abgeleitet aus den Markern (Anzeige)
    quelle: str
    regel_id: str
    grund: str
    salden: dict[str, float]


def _status(m: MappedAccount) -> str:
    """Status-Label aus den Markern (Anzeige, keine Logik)."""
    teile: list[str] = []
    if m.pflichtfrage == "aufriss":
        teile.append("Pflichtfrage: Aufriss")
    elif m.pflichtfrage == "pension":
        teile.append("Pflichtfrage: Pension")
    elif m.pflichtfrage:
        teile.append(f"Pflichtfrage: {m.pflichtfrage}")
    if m.verhaltenspruefung:
        teile.append("Verhaltensprüfung offen")
    if m.gekoppelt_mit:
        teile.append(f"gekoppelt mit {m.gekoppelt_mit}")
    if not teile:
        if m.klasse == Klasse.REVIEW:
            teile.append("ungelöst — Review")
        elif m.review:
            teile.append("Default gesetzt — bestätigen")
    return " · ".join(teile)


def _ist_relevant(m: MappedAccount) -> bool:
    return bool(m.review or m.klasse == Klasse.REVIEW
                or m.pflichtfrage or m.verhaltenspruefung)


def baue_review_queue(mapped: list[MappedAccount],
                      perioden: list[str]) -> list[ReviewEintrag]:
    eintraege: list[ReviewEintrag] = []
    for m in mapped:
        if not _ist_relevant(m):
            continue
        eintraege.append(ReviewEintrag(
            konto=m.konto, bezeichnung=m.bezeichnung, hgb_pfad=m.hgb_pfad,
            klasse=m.klasse.value, status=_status(m), quelle=m.quelle.value,
            regel_id=m.regel_id or "", grund=m.begruendung,
            salden={p: m.saldo(p) for p in perioden},
        ))
    # Pflichtfragen zuerst, dann übrige geflaggte Fälle
    eintraege.sort(key=lambda e: (0 if e.status.startswith("Pflichtfrage") else 1,
                                  e.konto))
    return eintraege
