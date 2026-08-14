"""Schicht 1 — die Mapping-Kaskade (Architektur-Spec v1.3).

Reihenfolge, erste greifende Quelle gewinnt:
  1. Kontennachweis / FS-Struktur des Abschlusses   (maßgeblich für Struktur)
  2. Hausconvention Typ-1                            (Verfeinerung; strukturell
                                                       nur ersatzweise ohne (1))
  3. Lernbibliothek                                  (v1: leer, Hook vorhanden)
  4. SKR-Default                                     (Sicherheitsnetz roh-SuSa)
  5. Review-Queue                                    (mit inaktivem KI-Hook davor)

Aus dem gesetzten HGB-Pfad wird die analytische Klasse ABGELEITET
(Reklassifizierung, Longest-Prefix). Nur die drei gemischten Positionen
bestimmen ND/OWC inhaltsabhängig über Typ-2-Regeln.
"""

from __future__ import annotations

from dataclasses import replace

from typing import Optional

from ..core.hausconvention import Hausconvention
from ..core.model import Account, Klasse, MappedAccount, NormalizedLedger, Quelle
from . import matcher, reclassify
from .ai_hook import KIUrteilsschicht
from .decision_log import Entscheidungsprotokoll

# HGB-Pfad-Suffixe der drei gemischten Positionen -> Typ-2-Regelset.
_MIXED_REGELSETS = {
    "Sonstige Vermoegensgegenstaende": "sonstige_vermoegensgegenstaende",
    "Sonstige Verbindlichkeiten": "sonstige_verbindlichkeiten",
    "Sonstige Rueckstellungen": "sonstige_rueckstellungen",
}


class Engine:
    def __init__(self, hc: Hausconvention,
                 protokoll: Optional[Entscheidungsprotokoll] = None,
                 ki: Optional[KIUrteilsschicht] = None,
                 lernbibliothek: Optional[dict[str, str]] = None):
        self.hc = hc
        self.protokoll = protokoll or Entscheidungsprotokoll()
        self.ki = ki or KIUrteilsschicht()          # v1: inaktiv
        self.lernbibliothek = lernbibliothek or {}   # konto -> hgb_pfad
        self._kanonische_pfade = [r.hgb_pfad for r in hc.reklass_regeln]

    # ---- öffentliche API --------------------------------------------------
    def map_ledger(self, ledger: NormalizedLedger) -> list[MappedAccount]:
        ergebnis = [self.map_account(a, ledger.hat_kontennachweis)
                    for a in ledger.accounts]
        for m in ergebnis:
            self.protokoll.protokolliere_map(m)
        return ergebnis

    def map_account(self, account: Account, hat_kontennachweis: bool) -> MappedAccount:
        # Technische DATEV-Konten (>=9000, EBK) fliegen früh raus — aber nur bei
        # SKR-nummerierten Konten OHNE gelieferte FS-Struktur. Liefert der
        # Abschluss den Pfad (z.B. SAP mit 10-stelligen Kontonummern), ist die
        # SKR-Schwelle bedeutungslos und der Abschluss hat bereits selektiert.
        # Der Reader kann ein Konto ausdrücklich als nicht abschlussrelevant
        # melden (SAP-Knoten "AUS — ausgesonderte Konten": IFRS-Brutto- und
        # Steuerbilanzkonten). Diese Aussage der Quelle sticht — sonst landeten
        # sie mangels HGB-Pfad in der Review-Queue und blähten sie auf.
        if account.kontotyp == "technisch":
            return self._tech(account)
        # Meldet der Reader ein Konto als ungeklärt (zwei Zeilen derselben
        # Kontonummer mit überschneidenden Perioden), darf die Kaskade keine
        # Zuordnung darüberlegen — das täuschte eine Klärung vor, die es nicht
        # gibt. Solche Konten gehen unmittelbar in die Review-Queue.
        if account.kontotyp == "strittig":
            m = self._review_ohne_pfad(account)
            return replace(m, begruendung=(
                "Kontonummer kommt mehrfach vor und die Werte überschneiden "
                "sich; welche Zeile gilt, ist ungeklärt."))
        if account.fs_pfad is None and self.hc.ist_technisch(account.konto):
            return self._tech(account)

        hgb_pfad, quelle, regel_id, begruendung = self._finde_pfad(
            account, hat_kontennachweis)

        if hgb_pfad is None:
            return self._review_ohne_pfad(account)

        return self._ableiten(account, hgb_pfad, quelle, regel_id, begruendung)

    # ---- Stufe 1..4: HGB-Pfad finden -------------------------------------
    def _finde_pfad(self, account: Account, hat_kontennachweis: bool
                    ) -> tuple[Optional[str], Quelle, Optional[str], str]:
        # Stufe 1: Kontennachweis / FS-Struktur (maßgeblich, vorrangig).
        # ``hat_kontennachweis`` markiert, dass die Quelle grundsätzlich eine
        # FS-Struktur liefert; die konkrete Zuordnung hängt am ``fs_pfad`` des
        # Einzelkontos. (v1.3-Feinheit "Hausconvention greift strukturell nur
        # ersatzweise ohne Kontennachweis" ist als spätere Verschärfung
        # vorgesehen; in v1 mappt ein Konto ohne fs_pfad regulär weiter über
        # Typ-1/SKR, statt sofort in Review zu gehen.)
        if account.fs_pfad:
            return (account.fs_pfad, Quelle.KONTENNACHWEIS, None,
                    "HGB-Pfad aus der FS-Struktur des Abschlusses (Abschlusstreue).")

        # Stufe 2: Hausconvention Typ-1
        r1 = matcher.match_typ1(account.bezeichnung, account.kontotyp,
                                self.hc.typ1_regeln)
        if r1 is not None:
            return (r1.hgb_pfad, Quelle.HAUSCONVENTION, r1.id,
                    f"Typ-1-Regel '{r1.id}' (Prio {r1.prioritaet}).")

        # Stufe 3: Lernbibliothek
        if account.konto in self.lernbibliothek:
            return (self.lernbibliothek[account.konto], Quelle.LERNBIBLIOTHEK, None,
                    "Bestätigte Zuordnung aus der Lernbibliothek.")

        # Stufe 4: SKR-Default (Sicherheitsnetz für rohe SuSa)
        skr = self.hc.skr_pfad(account.konto)
        if skr is not None:
            return (skr, Quelle.SKR_DEFAULT, None,
                    f"SKR03-Default über Kontonummer-Bereich ({account.konto}).")

        return (None, Quelle.REVIEW, None, "")

    # ---- Ableitung Klasse + NA-Zeile aus dem Pfad ------------------------
    def _ableiten(self, account: Account, hgb_pfad: str, quelle: Quelle,
                  regel_id: Optional[str], begruendung: str) -> MappedAccount:
        pfad_en = self.hc.uebersetze_pfad(hgb_pfad)

        # GuV-Konten: Klasse PL, keine WC/ND-Reklassifizierung (Hausconvention
        # _notes: "GuV-Konten ... Klasse = PL"). Für die Net-Debt-Sicht irrelevant.
        if hgb_pfad.startswith("/GuV"):
            na = hgb_pfad.split("/")[-1]
            return MappedAccount(
                account=account, hgb_pfad=hgb_pfad, hgb_pfad_en=pfad_en,
                klasse=Klasse.PL, na_de=na, na_en=self.hc.uebersetze_pfad(hgb_pfad).split("/")[-1],
                quelle=quelle, regel_id=regel_id, review=False,
                begruendung=begruendung + " GuV-Konto -> Klasse PL.",
            )

        reg = reclassify.reklassifiziere(hgb_pfad, self.hc)

        if reg is None:
            # Pfad gesetzt, aber keine Reklass-Regel -> geflaggt zur Review.
            m = MappedAccount(
                account=account, hgb_pfad=hgb_pfad, hgb_pfad_en=pfad_en,
                klasse=Klasse.REVIEW, na_de="(unbekannte Position)", na_en="(unknown)",
                quelle=quelle, regel_id=regel_id, review=True,
                begruendung=begruendung + " Kein Reklass-Eintrag zum Pfad — Review.",
            )
            return m

        if reg.klasse == "MIXED":
            return self._mixed(account, hgb_pfad, pfad_en, reg, quelle, regel_id,
                               begruendung)

        # Eindeutige Position: Klasse folgt direkt aus der Reklass-Tabelle.
        klasse = Klasse(reg.klasse)
        return MappedAccount(
            account=account, hgb_pfad=hgb_pfad, hgb_pfad_en=pfad_en,
            klasse=klasse, na_de=reg.na_de, na_en=reg.na_en,
            quelle=quelle, regel_id=regel_id, review=False,
            seite=_seite(hgb_pfad, klasse, reg),
            begruendung=begruendung + f" Klasse {klasse.value} via Reklassifizierung.",
        )

    def _mixed(self, account: Account, hgb_pfad: str, pfad_en: str, reg,
               quelle: Quelle, regel_id: Optional[str], begruendung: str
               ) -> MappedAccount:
        """Eine der drei gemischten Positionen: ND/OWC über Typ-2-Regel."""
        regelset = reg.split_regelset or _MIXED_REGELSETS.get(
            _letztes_bekanntes_segment(hgb_pfad))
        r2 = matcher.match_typ2(account.bezeichnung,
                                self.hc.typ2_regeln(regelset)) if regelset else None

        if r2 is None:
            return MappedAccount(
                account=account, hgb_pfad=hgb_pfad, hgb_pfad_en=pfad_en,
                klasse=Klasse.REVIEW, na_de=reg.na_de, na_en=reg.na_en,
                quelle=quelle, regel_id=regel_id, review=True,
                begruendung=begruendung + " Gemischte Position, keine Typ-2-Regel — Review.",
            )

        # REVIEW-Klasse (z.B. Direktversicherung) -> geflaggt, Default OWC.
        if r2.klasse == "REVIEW":
            klasse = Klasse.OWC
            review = True
        else:
            klasse = Klasse(r2.klasse)
            review = r2.review

        return MappedAccount(
            account=account, hgb_pfad=hgb_pfad, hgb_pfad_en=pfad_en,
            klasse=klasse, na_de=reg.na_de, na_en=reg.na_en,
            quelle=quelle, regel_id=r2.id, review=review, aus_mixed=True,
            # v2.3-Marker durchreichen — nur Status/Anzeige, keine Logik:
            pflichtfrage=r2.pflichtfrage, verhaltenspruefung=r2.verhaltenspruefung,
            gekoppelt_mit=r2.gekoppelt_mit, standardfrage=r2.standardfrage,
            seite=_seite(hgb_pfad, klasse, reg),
            begruendung=begruendung + f" Gemischt -> {klasse.value} via Typ-2 '{r2.id}'."
                        + (f" {r2.hinweis}" if r2.hinweis else ""),
        )

    # ---- Auffang-Zustände -------------------------------------------------
    def _review_ohne_pfad(self, account: Account) -> MappedAccount:
        # Hier säße im Live-Betrieb die KI-Urteilsschicht (v1: inaktiv).
        vorschlag = self.ki.vorschlagen(account, self._kanonische_pfade)
        if vorschlag is not None:
            m = self._ableiten(
                account, vorschlag.hgb_pfad, Quelle.KI_VORSCHLAG, None,
                f"KI-Vorschlag (Konfidenz {vorschlag.konfidenz:.0%}): {vorschlag.begruendung}",
            )
            m.review = True  # KI-Vorschlag immer geflaggt
            return m
        return MappedAccount(
            account=account, hgb_pfad="(offen)", hgb_pfad_en="(open)",
            klasse=Klasse.REVIEW, na_de="(unbestimmt)", na_en="(undetermined)",
            quelle=Quelle.REVIEW, regel_id=None, review=True,
            begruendung="Keine Kaskadenstufe griff — Review-Queue.",
        )

    def _tech(self, account: Account) -> MappedAccount:
        return MappedAccount(
            account=account, hgb_pfad="(technisch)", hgb_pfad_en="(technical)",
            klasse=Klasse.TECH, na_de="(technisch)", na_en="(technical)",
            quelle=Quelle.SKR_DEFAULT, regel_id=None, review=False,
            begruendung=f"Technisches Konto (>= {self.hc.tech_ab}) — nicht im Databook.",
        )


def _seite(hgb_pfad: str, klasse: Klasse, reg) -> Optional[str]:
    """WC-Seite (OA/OL) nach ``oa_ol_ableitung`` (Hausconvention v2.5).

    OA/OL sind keine eigene Klasse, sondern die Seite des Working Capital:
    /Aktiva + (TWC|OWC) -> OA, /Passiva + (TWC|OWC) -> OL. FA/ND/EQ/DT tragen
    keine Seite. Die Ableitung aus der Bilanzseite ist maßgeblich; das
    ``seite``-Feld der Reklass-Regel dient als Gegenprobe (Konfigfehler
    würden sonst stillschweigend durchlaufen).
    """
    if klasse not in (Klasse.TWC, Klasse.OWC):
        return None
    abgeleitet = "OA" if hgb_pfad.startswith("/Aktiva") else "OL"
    konfiguriert = getattr(reg, "seite", None)
    if konfiguriert and konfiguriert != abgeleitet:
        raise ValueError(
            f"Hausconvention-Widerspruch: Pfad '{reg.hgb_pfad}' ist laut "
            f"Bilanzseite {abgeleitet}, trägt aber seite={konfiguriert}."
        )
    return abgeleitet


def _letztes_bekanntes_segment(hgb_pfad: str) -> str:
    """Für MIXED-Pfade, die eine Ebene tiefer gehen: das gemischte Grund-
    segment zurückgeben, damit das richtige Regelset gewählt wird."""
    for seg in reversed([s for s in hgb_pfad.split("/") if s]):
        if seg in _MIXED_REGELSETS:
            return seg
    return hgb_pfad.split("/")[-1]
