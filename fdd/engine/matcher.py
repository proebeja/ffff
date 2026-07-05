"""Keyword-Matching für Typ-1 (HGB-Pfad) und Typ-2 (ND/OWC-Split).

Umsetzung exakt nach ``matching_logik`` der Hausconvention:
- ``alle``  : UND-verknüpft (alle Tokens müssen als Substring vorkommen)
- ``eines`` : ODER (mindestens eines)
- ``keines``: Ausschluss (keines darf vorkommen)
- höchste Priorität gewinnt.
"""

from __future__ import annotations

from typing import Optional

from ..core.hausconvention import Hausconvention, Typ1Regel, Typ2Regel, normalisiere


def _erfuellt(text_norm: str, alle: list[str], eines: list[str], keines: list[str]) -> bool:
    if any(k in text_norm for k in keines):
        return False
    if alle and not all(a in text_norm for a in alle):
        return False
    if eines and not any(e in text_norm for e in eines):
        return False
    # Eine Regel ohne jedes positive Kriterium greift nicht (Schutz vor Allmatch).
    if not alle and not eines:
        return False
    return True


def match_typ1(bezeichnung: str, kontotyp: Optional[str],
               regeln: list[Typ1Regel]) -> Optional[Typ1Regel]:
    """Beste Typ-1-Regel (höchste Priorität) oder None.

    ``kontotyp``-Gate: greift die Regel nur für einen bestimmten Kontotyp und
    passt der nicht, wird sie übersprungen (v2.1 Keyword-Riegel gegen GuV).
    Ist der Kontotyp unbekannt (None), lassen wir die Regel zu — der
    zusätzliche Keyword-Riegel ``keines`` fängt GuV-Fehlzündungen ab.
    """
    text = normalisiere(bezeichnung)
    treffer: list[Typ1Regel] = []
    for r in regeln:
        if r.kontotyp and kontotyp and r.kontotyp != kontotyp:
            continue
        if _erfuellt(text, r.alle, r.eines, r.keines):
            treffer.append(r)
    if not treffer:
        return None
    return max(treffer, key=lambda r: r.prioritaet)


def match_typ2(bezeichnung: str, regeln: list[Typ2Regel]) -> Optional[Typ2Regel]:
    """Beste Typ-2-Regel für ND/OWC-Split einer gemischten Position, oder None
    (dann Review-Queue)."""
    text = normalisiere(bezeichnung)
    treffer: list[Typ2Regel] = []
    for r in regeln:
        if _erfuellt(text, r.alle, r.eines, r.keines):
            treffer.append(r)
    if not treffer:
        return None
    return max(treffer, key=lambda r: r.prioritaet)
