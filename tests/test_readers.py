"""Reader-Regression gegen die vier realen Datensätze — inkl. der bekannten
Form-Fallstricke (DATEV-Spaltenversatz, SAP-Leerstrings, PDF-Wasserzeichen)."""

import pytest

from conftest import ECKART, SAP, HUCHTEMEIER, NAMUR, datei
from fdd.readers.base import parse_deutsche_zahl
from fdd.readers.datev_susa import DatevSusaReader
from fdd.readers.detect import waehle_reader
from fdd.readers.namur_databook import NamurDatabookReader
from fdd.readers.pdf_kontennachweis import PdfKontennachweisReader
from fdd.readers.sap_bw import SapBwReader


def test_detect_waehlt_richtigen_reader():
    assert isinstance(waehle_reader(datei(ECKART)), DatevSusaReader)
    assert isinstance(waehle_reader(datei(SAP)), SapBwReader)
    assert isinstance(waehle_reader(datei(HUCHTEMEIER)), PdfKontennachweisReader)
    assert isinstance(waehle_reader(datei(NAMUR)), NamurDatabookReader)


# ---- DATEV: Spaltenversatz korrekt aufgelöst ------------------------------
def test_datev_liest_konten_und_perioden():
    led = DatevSusaReader().lesen(datei(ECKART))
    assert len(led.accounts) > 150
    assert led.hat_kontennachweis is False
    assert set(led.perioden) == {"2024/12", "2023/12", "2022/S1", "2025/03"}
    assert led.fingerprint


def test_datev_saldo_und_vorzeichen_trotz_versatz():
    """Header 'Saldo' steht in Spalte T, Wert in S, Vorzeichen in W —
    der Reader muss den richtigen Wert mit richtigem Vorzeichen liefern."""
    led = DatevSusaReader().lesen(datei(ECKART))
    by = {a.konto: a for a in led.accounts}
    # 0956 Rückstellung Gewerbesteuer: Saldo 52.823 mit 'H' -> negativ (Passiv)
    assert by["0956"].saldo("2024/12") == pytest.approx(-52823, abs=1)
    # 1000 Kasse: Saldo 1.385,75 mit 'S' -> positiv (Aktiv)
    assert by["1000"].saldo("2024/12") == pytest.approx(1385.75, abs=0.01)


def test_datev_trial_balance_summiert_nahe_null():
    """Eine vorzeichenrichtige SuSa summiert je Periode ~0 (Soll=Haben)."""
    led = DatevSusaReader().lesen(datei(ECKART))
    for p in led.perioden:
        s = sum(a.saldo(p) for a in led.accounts)
        # Toleranz: SuSa kann kleine Rundungs-/Abgrenzungsreste tragen
        assert abs(s) < 5000, f"Periode {p}: Bilanzsumme {s:.2f} nicht ~0"


# ---- SAP: FS-Hierarchie + Leerstring-Toleranz -----------------------------
def test_sap_liest_blattkonten_mit_fs_pfad():
    led = SapBwReader().lesen(datei(SAP))
    assert led.hat_kontennachweis is True
    assert len(led.accounts) > 40
    # jedes Blattkonto trägt einen fs_pfad aus der SAP-Hierarchie (bis auf
    # nicht gecrosswalkte Knoten, die als Warnung vermerkt sind)
    mit_pfad = [a for a in led.accounts if a.fs_pfad]
    assert len(mit_pfad) > 40


def test_sap_intercompany_und_kasse_erkannt():
    led = SapBwReader().lesen(datei(SAP))
    by = {a.konto: a for a in led.accounts}
    # 1497000000 FORDG.L+L BAG-KONZERN -> verbundene Unternehmen
    assert "verbundene" in by["1497000000"].fs_pfad.lower()
    # 1000000000 KASSE -> Kassenbestand
    assert "Kassenbestand" in by["1000000000"].fs_pfad


# ---- PDF: absturzsicher trotz Wasserzeichen -------------------------------
def test_pdf_laeuft_ohne_absturz_und_liest_konten():
    led = PdfKontennachweisReader().lesen(datei(HUCHTEMEIER))
    assert len(led.accounts) > 50            # es kommen viele Konten durch
    by = {a.konto: a for a in led.accounts}
    # 1000 Kasse 1.159,47 sauber geparst
    assert by["1000"].saldo("31.12.2024") == pytest.approx(1159.47, abs=0.01)


# ---- Zahl-Parser ----------------------------------------------------------
@pytest.mark.parametrize("roh,erwartet", [
    ("1.999.512,89", 1999512.89),
    ("20.200,00-", -20200.00),      # DATEV nachgestelltes Minus
    ("-51.129,19", -51129.19),
    ("", 0.0),
    (None, 0.0),
    (1234.5, 1234.5),
    ("560457.80", 560457.80),       # Punkt als Dezimaltrenner
])
def test_parse_deutsche_zahl(roh, erwartet):
    assert parse_deutsche_zahl(roh) == pytest.approx(erwartet, abs=0.01)
