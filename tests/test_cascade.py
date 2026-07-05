"""Engine-Kaskade und die Klassifizierungslogik — inkl. des behobenen
'kst'-Kollisionsfehlers, der jede Rückstellung fälschlich zur
Steuerrückstellung machte."""

import pytest

from conftest import ECKART, datei
from fdd.core.model import Account, Klasse, PeriodBalance, Quelle
from fdd.engine.cascade import Engine
from fdd.readers.datev_susa import DatevSusaReader


def _konto(nr, bez, kontotyp=None, saldo=100.0):
    return Account(konto=nr, bezeichnung=bez,
                   salden=(PeriodBalance("P", saldo),), kontotyp=kontotyp)


@pytest.fixture
def eng(hc):
    return Engine(hc)


# ---- kst-Kollision behoben ------------------------------------------------
def test_kst_kollision_behoben_pensionen_bleiben_pensionen(eng):
    """'Rückstellungen f. Pensionen' darf NICHT als Steuerrückstellung
    zünden (früher via 'kst' ⊂ 'rüCKSTellung')."""
    m = eng.map_account(_konto("0950", "Rückstellungen f. Pensionen"), False)
    assert m.na_de == "Pensionsrueckstellungen"
    assert m.klasse == Klasse.ND
    assert m.regel_id == "hgb-pensions-rst"


def test_gewaehrleistung_ist_owc_nicht_steuer(eng):
    m = eng.map_account(_konto("0974", "Rückstell. f. Gewährleistungen"), False)
    assert m.na_de == "Sonstige Rueckstellungen"
    assert m.klasse == Klasse.OWC
    assert m.aus_mixed is True


def test_echte_steuerrueckstellung_bleibt_steuer(eng):
    for nr, bez in [("0956", "Rückstellung Gewerbesteuer"),
                    ("0963", "Körpersch.-St.-Rückstellung")]:
        m = eng.map_account(_konto(nr, bez), False)
        assert m.na_de == "Steuerrueckstellungen", bez
        assert m.klasse == Klasse.ND


# ---- Reklassifizierung eindeutiger Positionen -----------------------------
def test_kasse_ist_nd_via_skr_default(eng):
    m = eng.map_account(_konto("1000", "Kasse"), False)
    assert m.klasse == Klasse.ND
    assert m.na_de == "Liquide Mittel"
    assert m.quelle == Quelle.SKR_DEFAULT


def test_guv_konto_wird_pl(eng):
    m = eng.map_account(_konto("8400", "Erlöse 19% USt"), False)
    assert m.klasse == Klasse.PL
    assert m.hgb_pfad.startswith("/GuV")


def test_technisches_konto_ausgeklammert(eng):
    m = eng.map_account(_konto("9000", "Saldenvortrag Sachkonten"), False)
    assert m.klasse == Klasse.TECH


# ---- Typ-2-Split der gemischten Positionen --------------------------------
def test_mixed_split_darlehen_nd(eng):
    m = eng.map_account(_konto("1548", "Darlehen an Dritte"), False)
    assert m.klasse == Klasse.ND        # sva-darlehen
    assert m.aus_mixed is True


def test_mixed_split_kundenbonus_owc_vor_mitarbeiterbonus(eng):
    """rst-kundenboni (OWC, Vorrang) vor rst-boni (ND)."""
    m = eng.map_account(_konto("0979", "Rückstellung für Bonuszahlungen an Kunden"), False)
    assert m.klasse == Klasse.OWC


def test_mixed_ohne_regel_geht_in_review(eng):
    m = eng.map_account(_konto("1590", "Abstimmkonto ungeklärt"), False)
    # Sonstige VG ohne treffende Typ-2-Regel (nur sva-generic 'forderung'?) ->
    # hier kein Schlüsselwort -> Review
    assert m.review is True


# ---- Keyword-Riegel gegen GuV-Fehlzündung ---------------------------------
def test_guv_ertrag_aus_aufloesung_nicht_als_rueckstellung(eng):
    """Konto 'Erträge aus Auflösung von Rückstellungen' darf nicht als
    Rückstellung zünden (Keyword-Riegel 'ertraege'/'aufloesung')."""
    m = eng.map_account(_konto("2735", "Erträge aus Auflösung von Rückstellungen",
                               kontotyp="guv"), False)
    assert m.klasse != Klasse.ND
    assert "Rueckstellungen" not in m.na_de


# ---- Determinismus --------------------------------------------------------
def test_lauf_ist_reproduzierbar(hc):
    led = DatevSusaReader().lesen(datei(ECKART))
    a = [(m.konto, m.hgb_pfad, m.klasse.value) for m in Engine(hc).map_ledger(led)]
    b = [(m.konto, m.hgb_pfad, m.klasse.value) for m in Engine(hc).map_ledger(led)]
    assert a == b
