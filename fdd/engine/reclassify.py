"""Reklassifizierung: HGB-Pfad -> analytische Klasse + NA-Zielzeile.

Longest-Prefix-Match gegen die Tabelle ``reklassifizierung`` (matching_logik:
"Tiefere Pfade gewinnen vor kürzeren"). Die Klasse ist damit eine Eigenschaft
der POSITION, nicht des einzelnen Kontos — zwei Konten auf demselben Pfad
können nicht auseinanderlaufen.
"""

from __future__ import annotations

from typing import Optional

from ..core.hausconvention import Hausconvention, ReklassRegel


def _segmente(pfad: str) -> list[str]:
    return [s for s in pfad.split("/") if s]


def _ist_prefix(kandidat: str, pfad: str) -> bool:
    """True, wenn ``kandidat`` ein Pfad-Präfix von ``pfad`` ist — segmentweise,
    damit '.../Sonstige Vermoegensgegenstaende' nicht fälschlich auf
    '.../Sonstige Vermoegensgegenstaende XY' als reiner Stringpräfix matcht."""
    ks, ps = _segmente(kandidat), _segmente(pfad)
    if len(ks) > len(ps):
        return False
    return ps[:len(ks)] == ks


def reklassifiziere(hgb_pfad: str, hc: Hausconvention) -> Optional[ReklassRegel]:
    """Beste (längste) passende Reklass-Regel oder None (unbekannter Pfad)."""
    kandidaten: list[ReklassRegel] = [
        r for r in hc.reklass_regeln if _ist_prefix(r.hgb_pfad, hgb_pfad)
    ]
    if not kandidaten:
        return None
    return max(kandidaten, key=lambda r: len(_segmente(r.hgb_pfad)))
