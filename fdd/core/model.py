"""Kern-Datensätze der FDD-Pipeline.

Bewusst schlanke, unveränderliche Datencontainer. Sie sind das gemeinsame
Vokabular zwischen Reader, Engine, View und Export — die Engine sieht nie ein
Excel- oder PDF-spezifisches Objekt, nur diese Typen (Formatagnostik).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


class Klasse(str, Enum):
    """Analytische Net-Asset-Klasse. Wird abgeleitet, nie roh eingegeben.

    Werte spiegeln die ``klassifizierungs_codes`` der Hausconvention.
    ``MIXED`` ist ein Zwischenzustand für die drei gemischten Positionen,
    bevor die Typ-2-Regel ND/OWC entscheidet; ``REVIEW`` und ``TECH`` sind
    Auffang-Zustände.
    """

    FA = "FA"      # Fixed Assets / Anlagevermögen
    TWC = "TWC"    # Trade Working Capital
    OWC = "OWC"    # Other Working Capital
    ND = "ND"      # Net Debt / Nettofinanzverbindlichkeiten
    EQ = "EQ"      # Equity / Eigenkapital
    DT = "DT"      # Deferred Tax / Latente Steuern
    PL = "PL"      # Profit & Loss (keine WC/ND-Klassifizierung)
    MIXED = "MIXED"    # Zwischenzustand der drei gemischten Positionen
    REVIEW = "REVIEW"  # nicht auflösbar -> Review-Queue
    TECH = "TECH"      # technische DATEV-Konten, nicht im Databook


class Quelle(str, Enum):
    """Welche Kaskadenstufe den HGB-Pfad gesetzt hat (Protokoll/Audit)."""

    KONTENNACHWEIS = "Kontennachweis/FS-Struktur"
    HAUSCONVENTION = "Hausconvention (Typ-1)"
    LERNBIBLIOTHEK = "Lernbibliothek"
    SKR_DEFAULT = "SKR-Default"
    REVIEW = "Review-Queue"
    KI_VORSCHLAG = "KI-Urteilsschicht (Vorschlag)"
    OVERRIDE = "Manueller Override"
    # Nicht-HGB-Kontenrahmen (erster: AASB). Drei getrennte Werte statt eines
    # gemeinsamen, weil sich die drei Stufen in ihrer Verlässlichkeit deutlich
    # unterscheiden: das Stichwort trifft das Konto, die Kontogruppe nur seine
    # Nachbarschaft. Ohne die Unterscheidung ist eine Fehlzuordnung im
    # Mastersheet später nicht auffindbar.
    AASB_STICHWORT = "AASB-Stichwort"
    AASB_BIBLIOTHEK = "AASB-Bibliothek"
    AASB_GRUPPE = "AASB-Gruppe"


@dataclass(frozen=True)
class PeriodBalance:
    """Ein Saldo je Periode. ``betrag`` ist bereits vorzeichenbehaftet
    (Soll positiv, Haben negativ), sodass eine Trial Balance zu 0 summiert."""

    periode: str          # z.B. "2024/12", "31.12.2024"
    betrag: float


@dataclass(frozen=True)
class Account:
    """Ein normalisiertes Konto aus dem Reader — noch ungemappt.

    ``fs_pfad`` trägt die vom Abschluss/Kontennachweis vorgegebene Position
    (falls die Quelle sie liefert, z.B. SAP-Hierarchie oder PDF-Überschrift).
    Ist er gesetzt, ist er die maßgebliche Strukturquelle in der Kaskade.
    """

    konto: str
    bezeichnung: str
    salden: tuple[PeriodBalance, ...]
    entity: str = "Single-Entity"
    fs_pfad: Optional[str] = None      # HGB-Pfad aus dem Abschluss, falls vorhanden
    kontotyp: Optional[str] = None     # "bilanz_aktiv"/"bilanz_passiv"/"guv"/None
    # Kontogruppe der Quelle, falls die Saldenliste eine führt (MYOB:
    # "Current Assets", SAP: Knotenname). Sie ist bei Nicht-HGB-Rahmen die
    # letzte Stufe vor der Review-Queue — nie die erste: eine Gruppe nennt die
    # Nachbarschaft eines Kontos, nicht seinen Inhalt.
    gruppe: Optional[str] = None
    fristigkeit: Optional[str] = None  # "Current"/"Non-current", falls geliefert

    def saldo(self, periode: str) -> float:
        for pb in self.salden:
            if pb.periode == periode:
                return pb.betrag
        return 0.0

    @property
    def perioden(self) -> tuple[str, ...]:
        return tuple(pb.periode for pb in self.salden)


@dataclass
class MappedAccount:
    """Ergebnis der Engine je Konto — der strukturierte Datensatz aus
    Abschnitt 4.2 des Handovers."""

    account: Account
    hgb_pfad: str
    hgb_pfad_en: str
    klasse: Klasse
    na_de: str
    na_en: str
    quelle: Quelle
    regel_id: Optional[str] = None
    review: bool = False
    begruendung: str = ""
    wesentlich: Optional[bool] = None
    # v2.8 Seitenwechsel: Periode -> abweichender HGB-Pfad. Der Basispfad
    # bleibt in ``hgb_pfad``; das Mastersheet weist die Abweichung als eigene
    # Spalte aus, damit die Ableitung sichtbar bleibt.
    pfad_je_periode: dict = field(default_factory=dict)
    #: Periode -> (na_de, na_en) der abweichenden Position (Seitenwechsel).
    na_je_periode: dict = field(default_factory=dict)
    # True, wenn die Klasse über den Typ-2-Split einer der drei gemischten
    # Positionen entstand (thereof-ND-Logik der Net-Debt-View).
    aus_mixed: bool = False
    # Marker aus der Regel (v2.3) — nur Status/Anzeige, KEINE Logik dahinter:
    pflichtfrage: Optional[str] = None      # "aufriss" | "pension"
    verhaltenspruefung: bool = False        # Wiederkehr-/Stabilitätsprüfung offen
    gekoppelt_mit: Optional[str] = None     # Regel-ID einer gekoppelten Position
    standardfrage: str = ""                 # v2.5: Vollständigkeits-/Werthaltigkeitsfrage
    # v2.5: WC-Seite (OA/OL), abgeleitet aus der Bilanzseite des HGB-Pfads.
    # Keine eigene Klasse; None für FA/ND/EQ/DT.
    seite: Optional[str] = None
    # Audit-Spur bei manuellem Override der abgeleiteten Klasse:
    override_von: Optional[Klasse] = None
    override_begruendung: str = ""
    #: Welcher Kontenrahmen die Position gesetzt hat: "HGB" oder die ID des
    #: Nicht-HGB-Rahmens ("aasb"). Steuert keine Logik, sondern sagt, wie
    #: ``hgb_pfad`` zu lesen ist.
    rahmen: str = "HGB"
    #: Hinweise aus dem Kontenrahmen (z.B. Fristigkeits-Widerspruch). Kein
    #: Review-Grund für sich, aber im Mastersheet sichtbar.
    flags: list[str] = field(default_factory=list)

    @property
    def konto(self) -> str:
        return self.account.konto

    @property
    def bilanzseite(self) -> Optional[str]:
        """AKTIVA/PASSIVA, rahmenunabhängig. ``None`` für GuV und Auffang-
        Zustände.

        HGB-Pfade beginnen mit ``/Aktiva`` bzw. ``/Passiva``, Pfade eines
        Nicht-HGB-Rahmens mit ``/<Rahmen>/Aktiva``. Wer die Seite braucht,
        fragt hier und nicht ``hgb_pfad.startswith``.
        """
        for seg in self.hgb_pfad.split("/"):
            if seg in ("Aktiva", "Passiva"):
                return "AKTIVA" if seg == "Aktiva" else "PASSIVA"
        return None

    @property
    def bezeichnung(self) -> str:
        return self.account.bezeichnung

    def saldo(self, periode: str) -> float:
        return self.account.saldo(periode)

    def pfad_in(self, periode: str) -> str:
        """Wirksamer HGB-Pfad in dieser Periode. Weicht nur bei einem
        Seitenwechsel vom Basispfad ab (v2.8)."""
        return self.pfad_je_periode.get(periode, self.hgb_pfad)

    def na_in(self, periode: str) -> tuple[str, str]:
        return self.na_je_periode.get(periode, (self.na_de, self.na_en))

    def saldo_fuer_pfad(self, periode: str, pfad: str) -> float:
        """Beitrag zu genau dieser Position in dieser Periode. Ein Konto mit
        Seitenwechsel trägt je Periode nur zu einer der beiden Seiten bei."""
        return self.saldo(periode) if self.pfad_in(periode) == pfad else 0.0

    @property
    def alle_pfade(self) -> list[str]:
        return list(dict.fromkeys([self.hgb_pfad, *self.pfad_je_periode.values()]))


@dataclass
class NormalizedLedger:
    """Was ein Reader zurückgibt: eine formatagnostische Kontenliste plus
    Metadaten. ``hat_kontennachweis`` steuert die Kaskade (v1.3)."""

    accounts: list[Account]
    perioden: list[str]
    entity: str
    quelle_datei: str
    hat_kontennachweis: bool = False   # True, wenn FS-Struktur eingebettet ist
    fingerprint: str = ""              # Prüfsumme der Eingabe (Reproduzierbarkeit)
    warnungen: list[str] = field(default_factory=list)
