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


#: Statuslabel je Berichtssprache. Die Queue geht mit dem Databook an
#: denselben Leser, deshalb folgt sie derselben Sprache.
_LABEL = {
    "de": {"aufriss": "Pflichtfrage: Aufriss",
           "pension": "Pflichtfrage: Pension",
           "pflicht": "Pflichtfrage: {}",
           "verhalten": "Verhaltensprüfung offen",
           "gekoppelt": "gekoppelt mit {}",
           "review": "ungelöst — Review",
           "default": "Default gesetzt — bestätigen"},
    "en": {"aufriss": "mandatory question: breakdown",
           "pension": "mandatory question: pensions",
           "pflicht": "mandatory question: {}",
           "verhalten": "behaviour check open",
           "gekoppelt": "coupled with {}",
           "review": "unresolved — review",
           "default": "default applied — please confirm"},
}


def _status(m: MappedAccount, sprache: str = "de") -> str:
    """Status-Label aus den Markern (Anzeige, keine Logik)."""
    t = _LABEL["de" if sprache.lower().startswith("d") else "en"]
    teile: list[str] = []
    if m.pflichtfrage == "aufriss":
        teile.append(t["aufriss"])
    elif m.pflichtfrage == "pension":
        teile.append(t["pension"])
    elif m.pflichtfrage:
        teile.append(t["pflicht"].format(m.pflichtfrage))
    if m.verhaltenspruefung:
        teile.append(t["verhalten"])
    if m.gekoppelt_mit:
        teile.append(t["gekoppelt"].format(m.gekoppelt_mit))
    if not teile:
        if m.klasse == Klasse.REVIEW:
            teile.append(t["review"])
        elif m.review:
            teile.append(t["default"])
    return " · ".join(teile)


def _ist_relevant(m: MappedAccount) -> bool:
    return bool(m.review or m.klasse == Klasse.REVIEW
                or m.pflichtfrage or m.verhaltenspruefung)


def baue_review_queue(mapped: list[MappedAccount], perioden: list[str],
                      sprache: str = "de") -> list[ReviewEintrag]:
    # Pflichtfragen zuerst, dann übrige geflaggte Fälle. Sortiert wird am
    # Marker, nicht am übersetzten Text — sonst hinge die Reihenfolge an der
    # Berichtssprache.
    relevant = [m for m in mapped if _ist_relevant(m)]
    relevant.sort(key=lambda m: (0 if m.pflichtfrage else 1, m.konto))
    return [ReviewEintrag(
        konto=m.konto, bezeichnung=m.bezeichnung, hgb_pfad=m.hgb_pfad,
        klasse=m.klasse.value, status=_status(m, sprache),
        quelle=m.quelle.value, regel_id=m.regel_id or "",
        grund=m.begruendung, salden={p: m.saldo(p) for p in perioden},
    ) for m in relevant]
