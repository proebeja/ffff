"""Review-Queue als eigene Ausgabe: die ungelösten und geflaggten Konten."""

from __future__ import annotations

from dataclasses import dataclass

from ..core.model import Klasse, MappedAccount


@dataclass
class ReviewEintrag:
    konto: str
    bezeichnung: str
    hgb_pfad: str
    klasse: str
    quelle: str
    regel_id: str
    grund: str
    salden: dict[str, float]


def baue_review_queue(mapped: list[MappedAccount],
                      perioden: list[str]) -> list[ReviewEintrag]:
    eintraege: list[ReviewEintrag] = []
    for m in mapped:
        if not m.review and m.klasse != Klasse.REVIEW:
            continue
        eintraege.append(ReviewEintrag(
            konto=m.konto, bezeichnung=m.bezeichnung, hgb_pfad=m.hgb_pfad,
            klasse=m.klasse.value, quelle=m.quelle.value,
            regel_id=m.regel_id or "", grund=m.begruendung,
            salden={p: m.saldo(p) for p in perioden},
        ))
    return eintraege
