"""Absicherungen aus dem Code-Review: SUMIFS-Escaping, Doppelzähl-Riegel,
SAP-Bezeichnung, Namur-Periodenerkennung."""

import pytest

from fdd.export.excel import _krit, schreibe_databook
from fdd.readers.sap_bw import SapBwReader
from fdd.views.net_debt import NetDebtView, NetDebtZeile


def test_krit_escaped_quote_und_wildcards():
    assert _krit('Anleihen "AT1"') == 'Anleihen ""AT1""'
    assert _krit("Sonstige* Verb?") == "Sonstige~* Verb~?"
    assert _krit("~x") == "~~x"


def test_sap_bezeichnung_behaelt_nicht_konto_token():
    # Kunstfall: führendes Token ist NICHT die Kontonummer -> darf nicht wegfallen
    assert SapBwReader._bezeichnung("3M Klebeband", "H099999999") == "3M Klebeband"
    # echter Fall: führende (0-normalisierte) Kontonummer wird entfernt
    assert SapBwReader._bezeichnung("0992300000 Abgrenzung", "992300000") == "Abgrenzung"


def test_export_riegel_gegen_doppelzaehlung(tmp_path):
    """Erscheint dieselbe NA-Zeile in direkter UND Umgliederungs-Gruppe, muss
    der Export abbrechen statt doppelt zu summieren."""
    z1 = NetDebtZeile("Sonstige Verbindlichkeiten", "Other liab.", {"P": -10.0})
    z2 = NetDebtZeile("Sonstige Verbindlichkeiten", "Other liab.", {"P": -5.0},
                      aus_mixed=True)
    nd = NetDebtView(perioden=["P"], direkt=[z1], umgliederung=[z2], entity="X")
    with pytest.raises(ValueError, match="doppelt"):
        schreibe_databook(str(tmp_path / "x.xlsx"), [], nd, [], ["P"], "X")
