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
    # True, wenn die Klasse über den Typ-2-Split einer der drei gemischten
    # Positionen entstand (thereof-ND-Logik der Net-Debt-View).
    aus_mixed: bool = False
    # Marker aus der Regel (v2.3) — nur Status/Anzeige, KEINE Logik dahinter:
    pflichtfrage: Optional[str] = None      # "aufriss" | "pension"
    verhaltenspruefung: bool = False        # Wiederkehr-/Stabilitätsprüfung offen
    gekoppelt_mit: Optional[str] = None     # Regel-ID einer gekoppelten Position
    # Audit-Spur bei manuellem Override der abgeleiteten Klasse:
    override_von: Optional[Klasse] = None
    override_begruendung: str = ""

    @property
    def konto(self) -> str:
        return self.account.konto

    @property
    def bezeichnung(self) -> str:
        return self.account.bezeichnung

    def saldo(self, periode: str) -> float:
        return self.account.saldo(periode)


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
