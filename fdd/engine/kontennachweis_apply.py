"""Kontennachweis auf eine SuSa anwenden (Kaskadenstufe 1 scharf schalten).

Die SuSa liefert die Salden, der Kontennachweis die **Struktur**. Zwei Wirkungen:

1. Jedes Konto, das der Kontennachweis kennt, bekommt seinen ``fs_pfad`` —
   damit greift Stufe 1 der Kaskade und der SKR-Default kommt gar nicht mehr
   zum Zug. Der SKR-Default bleibt nur für Konten, zu denen der
   Kontennachweis schweigt.

2. Konten, die der Kontennachweis führt, die SuSa aber nicht enthält, werden
   **ergänzt** (mit den Salden aus dem Kontennachweis). Genau hier schließt
   sich die Lücke, durch die bisher Positionen komplett fehlten.

Die Salden der SuSa werden dabei nie überschrieben — sie bleiben die
Werteseite; der Kontennachweis liefert nur Struktur bzw. fehlende Konten.
"""

from __future__ import annotations

from dataclasses import replace

from ..core.model import Account, NormalizedLedger, PeriodBalance
from ..readers.kontennachweis import Kontennachweis


def wende_kontennachweis_an(ledger: NormalizedLedger,
                            kn: Kontennachweis) -> NormalizedLedger:
    """Gibt ein neues Ledger zurück: Struktur aus dem Kontennachweis,
    Salden aus der SuSa, fehlende Konten ergänzt."""
    zuordnung = kn.zuordnung
    warnungen = list(ledger.warnungen)
    neue_accounts: list[Account] = []
    gesehen: set[str] = set()

    for a in ledger.accounts:
        pfad = zuordnung.get(a.konto)
        gesehen.add(a.konto)
        if pfad:
            neue_accounts.append(replace(a, fs_pfad=pfad,
                                         kontotyp=_kontotyp(pfad)))
        else:
            neue_accounts.append(a)

    fehlend = [k for k in kn.konten if k not in gesehen]
    for konto in sorted(fehlend):
        kk = kn.konten[konto]
        salden = tuple(PeriodBalance(p, kk.salden.get(p, 0.0))
                       for p in ledger.perioden)
        neue_accounts.append(Account(
            konto=kk.konto, bezeichnung=kk.bezeichnung, salden=salden,
            entity=ledger.entity, fs_pfad=kk.hgb_pfad,
            kontotyp=_kontotyp(kk.hgb_pfad)))
        warnungen.append(
            f"Konto {kk.konto} '{kk.bezeichnung}' fehlt in der SuSa und wurde "
            "aus dem Kontennachweis ergänzt.")

    return NormalizedLedger(
        accounts=neue_accounts, perioden=list(ledger.perioden),
        entity=ledger.entity, quelle_datei=ledger.quelle_datei,
        hat_kontennachweis=True,
        fingerprint=f"{ledger.fingerprint}+kn:{kn.fingerprint}",
        warnungen=warnungen,
    )


def _kontotyp(hgb_pfad: str) -> str | None:
    if hgb_pfad.startswith("/Aktiva"):
        return "bilanz_aktiv"
    if hgb_pfad.startswith("/Passiva"):
        return "bilanz_passiv"
    if hgb_pfad.startswith("/GuV"):
        return "guv"
    return None
