"""Formaterkennung — wählt den passenden Reader. Hält die Formatlogik an
einer Stelle gekapselt, damit CLI und Engine formatagnostisch bleiben."""

from __future__ import annotations

from .base import Reader
from .datev_kanzlei_susa import KanzleiSusaReader
from .datev_susa import DatevSusaReader
from .namur_databook import NamurDatabookReader
from .pdf_kontennachweis import PdfKontennachweisReader
from .sap_bw import SapBwReader

# Reihenfolge = Spezifität: Namur/SAP (spezifische Sheets) vor DATEV, PDF zuletzt.
_READER: list[type[Reader]] = [
    NamurDatabookReader, SapBwReader, DatevSusaReader, KanzleiSusaReader,
    PdfKontennachweisReader,
]


def waehle_reader(pfad: str) -> Reader:
    for cls in _READER:
        try:
            if cls.kann_lesen(pfad):
                return cls()
        except Exception:
            continue
    raise ValueError(f"Kein passender Reader für: {pfad}")
