"""EINFRIEREN — ein Rerun leitet nur neu her, was sich geändert hat.

``ki_urteilsschicht.regeln`` verlangt: einmal getroffene Entscheidungen werden
im Entscheidungsprotokoll des Mandats festgeschrieben und beim nächsten Lauf
gelesen statt neu gefragt. Das Databook bleibt damit exakt reproduzierbar.

Dieses Modul zieht die Regel auf den ganzen Lauf hoch:

* Konten, die eine der benannten Änderungen **berührt**, werden neu hergeleitet.
* Alle übrigen behalten das eingefrorene Ergebnis.
* Weicht ein **unberührtes** Konto in der frischen Herleitung trotzdem ab, ist
  das ein **Reproduzierbarkeitsdefekt**: derselbe Input hätte dasselbe Ergebnis
  liefern müssen. Er wird protokolliert und das eingefrorene Ergebnis gilt.

Der letzte Punkt ist der eigentliche Zweck. Ein Rerun, der stillschweigend
andere Zahlen liefert, ist schlimmer als einer, der gar nicht läuft.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field, replace
from typing import Optional

from ..core.model import Klasse, MappedAccount


@dataclass
class DeltaEintrag:
    konto: str
    bezeichnung: str
    feld: str                     # "klasse" | "hgb_pfad"
    vorher: str
    nachher: str
    ausloeser: str                # welche Änderung, oder "UNERKLÄRT"

    @property
    def ist_defekt(self) -> bool:
        return self.ausloeser == UNERKLAERT


UNERKLAERT = "UNERKLÄRT — Reproduzierbarkeitsdefekt"


def _als_text(pfad_je_periode: dict) -> str:
    if not pfad_je_periode:
        return ""
    return "; ".join(f"{p}→{z.rsplit('/', 1)[-1]}"
                     for p, z in sorted(pfad_je_periode.items()))


@dataclass
class Einfrierergebnis:
    mapped: list[MappedAccount]
    delta: list[DeltaEintrag] = field(default_factory=list)
    eingefroren: int = 0
    neu_hergeleitet: int = 0
    neue_konten: list[str] = field(default_factory=list)

    @property
    def defekte(self) -> list[DeltaEintrag]:
        return [d for d in self.delta if d.ist_defekt]

    def je_ausloeser(self) -> dict[str, int]:
        out: dict[str, int] = {}
        for d in self.delta:
            out[d.ausloeser] = out.get(d.ausloeser, 0) + 1
        return out


def lade_snapshot(pfad: str) -> dict:
    with open(pfad, encoding="utf-8") as f:
        return json.load(f)


def schreibe_snapshot(pfad: str, mapped: list[MappedAccount], *,
                      lauf: int, hausconvention: str, fingerprint: str) -> None:
    daten = {
        "lauf": lauf, "hausconvention": hausconvention, "fingerprint": fingerprint,
        "konten": {m.konto: {"klasse": m.klasse.value, "hgb_pfad": m.hgb_pfad,
                             "pfad_je_periode": _als_text(m.pfad_je_periode),
                             "na_de": m.na_de, "quelle": m.quelle.value,
                             "bezeichnung": m.bezeichnung}
                   for m in mapped},
    }
    with open(pfad, "w", encoding="utf-8") as f:
        json.dump(daten, f, ensure_ascii=False, indent=1, sort_keys=True)


def bestimme_ausloeser(m: MappedAccount, hc) -> Optional[str]:
    """Welche der benannten Änderungen berührt dieses Konto?

    Bewusst am Ergebnis abgelesen und nicht an einer Kontoliste: eine
    Kontoliste wäre eine zweite Wahrheit, die mit der Regel auseinanderlaufen
    kann."""
    from .v28 import NICHT_ZUGEORDNET

    if m.pfad_je_periode:
        return "1 · Seitenwechsel (periodenabhängige Bilanzseite)"
    if hc.ist_saldenvortrag(m.konto):
        return "4 · Saldenvortragskonten nicht mehr TECH"
    if m.hgb_pfad == NICHT_ZUGEORDNET or "QA A6" in (m.begruendung or ""):
        return "3 · QA A6 (jedes Konto landet in einer Bilanzposition)"
    if m.regel_id and (m.regel_id.startswith("zdl-")
                       or m.regel_id.startswith("gesellschafter-gf")):
        return f"5 · Regelgruppe Zahlungsverkehr ({m.regel_id})"
    return None


def wende_einfrierung_an(mapped: list[MappedAccount], snapshot: dict, hc
                         ) -> Einfrierergebnis:
    """Frisch hergeleitet gegen eingefroren. Berührte Konten übernehmen die
    frische Herleitung, unberührte behalten die eingefrorene."""
    alt = snapshot.get("konten", {})
    erg = Einfrierergebnis(mapped=[])

    for m in mapped:
        vorher = alt.get(m.konto)
        ausloeser = bestimme_ausloeser(m, hc)

        if vorher is None:
            erg.neue_konten.append(m.konto)
            erg.mapped.append(m)
            erg.neu_hergeleitet += 1
            continue

        # Der Seitenwechsel ändert nicht den Basispfad, sondern ergänzt eine
        # periodenabhängige Abweichung. Ohne dieses Feld bliebe die wichtigste
        # Änderung des Laufs im Delta unsichtbar.
        felder = {"klasse": m.klasse.value, "hgb_pfad": m.hgb_pfad,
                  "pfad_je_periode": _als_text(m.pfad_je_periode)}
        geaendert = [(f, vorher.get(f, ""), neu) for f, neu in felder.items()
                     if vorher.get(f, "") != neu]

        if ausloeser:
            for feld, v, n in geaendert:
                erg.delta.append(DeltaEintrag(m.konto, m.bezeichnung, feld,
                                              v, n, ausloeser))
            erg.mapped.append(m)
            erg.neu_hergeleitet += 1
            continue

        if geaendert:
            # Unberührt, aber abweichend: Defekt protokollieren und den
            # eingefrorenen Stand wiederherstellen.
            for feld, v, n in geaendert:
                erg.delta.append(DeltaEintrag(m.konto, m.bezeichnung, feld,
                                              v, n, UNERKLAERT))
            m = replace(m, klasse=Klasse(vorher["klasse"]),
                        hgb_pfad=vorher["hgb_pfad"], na_de=vorher["na_de"],
                        begruendung=(m.begruendung + " [eingefroren aus Lauf 1: "
                                     "keine der benannten Änderungen berührt "
                                     "dieses Konto.]"))
        erg.mapped.append(m)
        erg.eingefroren += 1
    return erg
