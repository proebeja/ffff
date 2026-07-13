"""Longest-Prefix-Reklassifizierung, Override und Matcher-Boundary-Verhalten."""

import pytest

from fdd.core.model import Account, Klasse, PeriodBalance
from fdd.engine.cascade import Engine
from fdd.engine.decision_log import Entscheidungsprotokoll
from fdd.engine.matcher import match_typ1
from fdd.engine.reclassify import reklassifiziere


def _konto(nr, bez):
    return Account(konto=nr, bezeichnung=bez, salden=(PeriodBalance("P", 1.0),))


# ---- Longest-Prefix -------------------------------------------------------
def test_longest_prefix_tiefer_gewinnt(hc):
    # Wertpapiere des AV (tief, ND) schlägt Finanzanlagen-Ebene
    r = reklassifiziere(
        "/Aktiva/A Anlagevermoegen/III Finanzanlagen/Wertpapiere des Anlagevermoegens", hc)
    assert r.klasse == "ND"


def test_prefix_ist_segmentweise(hc):
    # exakte gemischte Position -> MIXED
    r = reklassifiziere(
        "/Aktiva/B Umlaufvermoegen/II Forderungen und sonstige Vermoegensgegenstaende/"
        "Sonstige Vermoegensgegenstaende", hc)
    assert r.klasse == "MIXED"


def test_mixed_unterzeile_erbt_mixed(hc):
    # eine Ebene tiefer (Vorsteuer) muss noch die MIXED-Position treffen
    r = reklassifiziere(
        "/Aktiva/B Umlaufvermoegen/II Forderungen und sonstige Vermoegensgegenstaende/"
        "Sonstige Vermoegensgegenstaende/Vorsteuer", hc)
    assert r.klasse == "MIXED"


# ---- Matcher: 'kst' zündet nicht mehr auf 'rueckstellung' -----------------
def test_matcher_kst_nicht_auf_rueckstellung(hc):
    r = match_typ1("Rückstellungen Personalkosten", "bilanz_passiv", hc.typ1_regeln)
    assert r is None or r.id != "hgb-steuer-rst"


def test_matcher_kst_variante_zuendet_auf_echte_kst(hc):
    r = match_typ1("KSt-Rückstellung 2024", "bilanz_passiv", hc.typ1_regeln)
    assert r is not None and r.id == "hgb-steuer-rst"


def test_matcher_compound_gewerbesteuerrueckstellung(hc):
    """Deutsche Komposita (ein Wort) müssen weiter über Substring greifen."""
    r = match_typ1("Gewerbesteuerrückstellung", "bilanz_passiv", hc.typ1_regeln)
    assert r is not None and r.id == "hgb-steuer-rst"


# ---- Override (begründungspflichtig, protokolliert) -----------------------
def test_override_verlangt_begruendung(hc):
    eng = Engine(hc, protokoll=Entscheidungsprotokoll())
    # v2.3: 'Rückstellungen Personalkosten' -> OWC (rst-personalkosten-generisch,
    # nutzerbestätigt). Der Override demonstriert die begründungspflichtige
    # manuelle Korrektur einer abgeleiteten Klasse.
    m = eng.map_account(_konto("0965", "Rückstellungen Personalkosten"), False)
    assert m.klasse == Klasse.OWC
    with pytest.raises(ValueError):
        eng.protokoll.override_klasse(m, Klasse.ND, "")
    eng.protokoll.override_klasse(m, Klasse.ND, "Aufriss zeigt Boni-Anteil (debt-like)")
    assert m.klasse == Klasse.ND
    assert m.override_von == Klasse.OWC
    assert any(e.aktion == "override" for e in eng.protokoll.eintraege)
