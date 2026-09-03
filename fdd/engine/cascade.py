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

**Nicht-HGB-Mandate.** Bekommt die Engine einen ``kontenrahmen`` (AASB, später
IFRS/US GAAP/UK GAAP), läuft statt der HGB-Kette die Kaskade dieses Rahmens:
Kontennachweis, dann Stichwortregeln, Kontenbibliothek und Kontogruppe gegen
den Kontonamen. Beide Wege enden gleich — bei einer Klasse und einer
NA-Zeile — und sie berühren einander nicht: ohne ``kontenrahmen`` läuft
ausschließlich der HGB-Weg, mit ``kontenrahmen`` ausschließlich der andere.
Die einzige Stelle, an der sie sich treffen, sind die Typ-2-Regeln für
gemischte Positionen; sie sind inhaltlich (Gesellschafterdarlehen, Kreditkarte,
Kaution) und nicht rechtsformspezifisch formuliert.
"""

from __future__ import annotations

from dataclasses import replace

from typing import Optional

from ..core.hausconvention import Hausconvention
from ..core.kontenrahmen import Kontenrahmen
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
                 lernbibliothek: Optional[dict[str, str]] = None,
                 kontenrahmen: Optional[Kontenrahmen] = None,
                 konzernnamen: Optional[list[str]] = None):
        self.hc = hc
        self.protokoll = protokoll or Entscheidungsprotokoll()
        self.ki = ki or KIUrteilsschicht()          # v1: inaktiv
        self.lernbibliothek = lernbibliothek or {}   # konto -> hgb_pfad
        #: Nicht-HGB-Rahmen aus dem Setup-Dialog. ``None`` = HGB.
        self.kontenrahmen = kontenrahmen
        #: Namen der verbundenen Unternehmen und Gesellschafter (Setup-Frage 5).
        self.konzernnamen = konzernnamen or []
        self._kanonische_pfade = [r.hgb_pfad for r in hc.reklass_regeln]

    # ---- öffentliche API --------------------------------------------------
    def map_ledger(self, ledger: NormalizedLedger) -> list[MappedAccount]:
        ergebnis = [self.map_account(a, ledger.hat_kontennachweis)
                    for a in ledger.accounts]
        for m in ergebnis:
            self.protokoll.protokolliere_map(m)
        return ergebnis

    @property
    def ki_aufrufe(self) -> list[str]:
        """Wofür die KI-Schicht gerufen wurde. Ohne registrierten Provider
        bleibt die Liste leer — die Schicht ist eine Schnittstelle."""
        return getattr(self, "_ki_aufrufe", [])

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
        # Nicht-HGB-Mandat: eigener Weg. Er steht vor der SKR-Technikschwelle,
        # weil die Grenze 9000 eine Eigenheit des SKR ist — in einem
        # australischen Kontenplan ist "9-0100 Interest Income" ein
        # gewöhnliches Ertragskonto und kein Statistikkonto.
        if self.kontenrahmen is not None:
            return self._map_kontenrahmen(account)
        # v2.8: Saldenvortragskonten fallen nicht mehr pauschal in TECH. Sie
        # werden nachgelagert abgestimmt (siehe engine/saldenvortrag.py); der
        # nicht abstimmbare Rest gehört ins Eigenkapital.
        if (account.fs_pfad is None and self.hc.ist_technisch(account.konto)
                and not self.hc.ist_saldenvortrag(account.konto)):
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

    # ---- Nicht-HGB-Rahmen (AASB, später IFRS/US GAAP/UK GAAP) ------------
    #: Klasse ``MIXED`` des Rahmens -> Typ-2-Regelset der Hausconvention. Die
    #: Regelsets sind nach Bilanzposten benannt, ihre Regeln aber inhaltlich
    #: formuliert (Gesellschafterdarlehen, Kreditkarte, Kaution) und damit
    #: rechtsformunabhängig. Zugeordnet wird deshalb über Seite und Art der
    #: Position, nicht über den Namen des FS Line Items.
    _MIXED_AASB = {
        "AKTIVA": "sonstige_vermoegensgegenstaende",
        "PASSIVA": "sonstige_verbindlichkeiten",
        "RUECKSTELLUNG": "sonstige_rueckstellungen",
    }

    def _map_kontenrahmen(self, account: Account) -> MappedAccount:
        """Die Kaskade des Nicht-HGB-Rahmens für ein Konto.

        Stufe 1 bleibt der Kontennachweis: liefert die Quelle eine Position
        (``fs_pfad``), gilt sie. Alles Weitere macht der Rahmen selbst.
        """
        kr = self.kontenrahmen
        seite_quelle = _seite_aus_kontotyp(account.kontotyp)
        ist_guv = (account.kontotyp == "guv") if account.kontotyp else None

        if account.fs_pfad and account.fs_pfad in kr.fs_lines:
            z_fs, z_klasse = account.fs_pfad, kr.klasse_von(account.fs_pfad)
            quelle, regel_id = Quelle.KONTENNACHWEIS, None
            begruendung = ("FS Line Item aus dem Kontennachweis des "
                           "Abschlusses (Abschlusstreue).")
            seite, flags = seite_quelle or kr.seite_von(z_fs), []
        else:
            z = kr.zuordnen(account.bezeichnung, account.gruppe or "",
                            account.fristigkeit or "", seite=seite_quelle,
                            konzernnamen=self.konzernnamen, ist_guv=ist_guv)
            if z is None:
                # Ein GuV-Konto ohne Positionstreffer bleibt ein GuV-Konto.
                # Es in die Review-Queue OHNE Klasse zu schicken, hiesse es
                # aus der GuV zu nehmen und das Ergebnis zu verfaelschen; die
                # offene Frage ist die Position, nicht das Rechenwerk.
                #
                # Als Position tritt dann die Kontogruppe des Mandanten ein.
                # In der Bilanz waere das genau der Fehler, um den es in
                # diesem Rahmen geht — "R&D/Demo tools - Waterloo" ist keine
                # Bilanzposition. In der GuV ist die Kontogruppe dagegen eine
                # zulaessige Gliederung: sie ist die Gliederung, nach der der
                # Mandant selbst berichtet. Woher sie stammt, bleibt in der
                # Quellenspalte sichtbar.
                if ist_guv:
                    gruppe = (account.gruppe or "").strip()
                    return self._aasb_ergebnis(
                        account, kr, gruppe or "(Position offen)", Klasse.PL,
                        None, Quelle.AASB_GRUPPE if gruppe else Quelle.REVIEW,
                        "gruppe:mandant" if gruppe else None, [],
                        (f"GuV-Konto laut Quelle; der Rahmen kennt keine "
                         f"Position dafuer. Gegliedert wird nach der "
                         f"Kontogruppe des Mandanten ('{gruppe}')."
                         if gruppe else
                         f"GuV-Konto laut Quelle, aber weder Bibliothek noch "
                         f"Kontogruppe trafen '{account.bezeichnung}' — die "
                         f"Position ist offen, das Rechenwerk nicht."),
                        review=not gruppe)
                m = self._review_ohne_pfad(account)
                m.rahmen = kr.id
                return replace(m, begruendung=(
                    f"Weder Stichwort noch Bibliothek noch Kontogruppe "
                    f"trafen '{account.bezeichnung}' — Review-Queue."))
            z_fs, z_klasse = z.fs_line, z.klasse
            quelle = Quelle(z.quelle)
            regel_id, begruendung = z.regel_id, z.begruendung
            seite, flags = z.seite, list(z.flags)

        # GuV: der Rahmen führt die Ertrags- und Aufwandszeilen nur in der
        # Bibliothek, nicht in der Reklassifizierungstabelle. Sie tragen
        # Klasse PL und nehmen an der Net-Asset-Sicht nicht teil.
        if z_klasse is None or kr.ist_guv(z_fs):
            return self._aasb_ergebnis(account, kr, z_fs, Klasse.PL, None,
                                       quelle, regel_id, flags,
                                       begruendung + " GuV-Zeile -> Klasse PL.")

        if z_klasse == "MIXED":
            return self._mixed_kontenrahmen(account, kr, z_fs, seite, quelle,
                                            regel_id, flags, begruendung)

        klasse = Klasse(z_klasse)
        return self._aasb_ergebnis(
            account, kr, z_fs, klasse, seite, quelle, regel_id, flags,
            begruendung + f" Klasse {klasse.value} via Reklassifizierung.")

    def _mixed_kontenrahmen(self, account: Account, kr: Kontenrahmen,
                            fs_line: str, seite: Optional[str], quelle: Quelle,
                            regel_id: Optional[str], flags: list[str],
                            begruendung: str) -> MappedAccount:
        """Gemischte Position des Rahmens: die Typ-2-Regeln entscheiden.

        18 der 67 FS Line Items sind gemischt — 'Provisions' und 'Employee
        benefit obligations' sind je nach Inhalt Net Debt oder Working
        Capital. Diese Frage beantwortet kein Kontenrahmen, sondern nur der
        Kontoinhalt; zuständig sind dieselben Typ-2-Regeln wie bei HGB.
        """
        art = ("RUECKSTELLUNG" if "provision" in fs_line.lower()
               else seite or "PASSIVA")
        regelset = self._MIXED_AASB.get(art)
        r2 = matcher.match_typ2(account.bezeichnung, self.hc.typ2_regeln(regelset))
        if r2 is None:
            return self._aasb_ergebnis(
                account, kr, fs_line, Klasse.REVIEW, seite, quelle, regel_id,
                flags, begruendung + f" Gemischte Position '{fs_line}', keine "
                f"Typ-2-Regel griff — Review.", review=True)
        klasse = Klasse.OWC if r2.klasse == "REVIEW" else Klasse(r2.klasse)
        m = self._aasb_ergebnis(
            account, kr, fs_line, klasse, seite, quelle, r2.id, flags,
            begruendung + f" Gemischt -> {klasse.value} via Typ-2 '{r2.id}'."
            + (f" {r2.hinweis}" if r2.hinweis else ""),
            review=r2.review or r2.klasse == "REVIEW")
        m.aus_mixed = True
        m.pflichtfrage = r2.pflichtfrage
        m.verhaltenspruefung = r2.verhaltenspruefung
        m.gekoppelt_mit = r2.gekoppelt_mit
        m.standardfrage = r2.standardfrage
        return m

    def _aasb_ergebnis(self, account: Account, kr: Kontenrahmen, fs_line: str,
                       klasse: Klasse, seite: Optional[str], quelle: Quelle,
                       regel_id: Optional[str], flags: list[str],
                       begruendung: str, review: bool = False) -> MappedAccount:
        """Baut den Datensatz. Der Pfad trägt den Rahmen im ersten Segment.

        ``/AASB/Aktiva/Trade receivables, net`` statt ``/Aktiva/...``: ein
        FS Line Item ist kein HGB-Pfad und darf nicht so aussehen. Wer die
        Bilanzseite braucht, nimmt ``MappedAccount.bilanzseite``.

        DE- und EN-Pfad sind identisch. Der Rahmen ist englisch geführt und
        kennt keine deutsche Entsprechung; eine erfundene Übersetzung wäre
        schlechter als die Wiederholung des Originals.
        """
        knoten = ("GuV" if klasse == Klasse.PL
                  else "Aktiva" if seite == "AKTIVA"
                  else "Passiva" if seite == "PASSIVA" else "unbestimmt")
        pfad = f"/{kr.id.upper()}/{knoten}/{fs_line}"
        return MappedAccount(
            account=account, hgb_pfad=pfad, hgb_pfad_en=pfad, klasse=klasse,
            na_de=fs_line, na_en=fs_line, quelle=quelle, regel_id=regel_id,
            review=review or knoten == "unbestimmt",
            seite=(None if klasse not in (Klasse.TWC, Klasse.OWC)
                   else "OA" if seite == "AKTIVA" else "OL"),
            rahmen=kr.id, flags=flags,
            begruendung=begruendung + ("" if knoten != "unbestimmt" else
                                       " Bilanzseite nicht bestimmbar — Review."),
        )

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


def _seite_aus_kontotyp(kontotyp: Optional[str]) -> Optional[str]:
    """Bilanzseite, soweit der Reader sie gemeldet hat. Sie ist verlässlicher
    als jede Ableitung aus Kontonamen oder Vorzeichen und sticht deshalb."""
    if kontotyp == "bilanz_aktiv":
        return "AKTIVA"
    if kontotyp == "bilanz_passiv":
        return "PASSIVA"
    return None


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
