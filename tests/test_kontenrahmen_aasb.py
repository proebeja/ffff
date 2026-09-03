"""Der AASB-Kontenrahmen: Reihenfolge, Vorrang, Wortgrenze — und die Zusage,
dass die HGB-Kette davon nichts merkt.

Die Fälle stammen aus der Übergabe zum Kontenrahmen. Sie prüfen nicht, ob eine
Zuordnung gefällt, sondern ob die Kaskade in der festgelegten Reihenfolge
läuft: Stichwortregeln vor Kontenbibliothek vor Kontogruppe, und die
Konzernregel vor allem anderen.
"""

from __future__ import annotations

import pytest

from fdd.core.hausconvention import Hausconvention
from fdd.core.kontenrahmen import Kontenrahmen, lade_rahmen
from fdd.core.model import Account, Klasse, PeriodBalance, Quelle
from fdd.engine.cascade import Engine
from fdd.engine.setup import setup


@pytest.fixture(scope="module")
def aasb() -> Kontenrahmen:
    return lade_rahmen("aasb")


def konto(bezeichnung, konto="1-00000", typ=None, gruppe=None, frist=None):
    return Account(konto=konto, bezeichnung=bezeichnung,
                   salden=(PeriodBalance("FY2024", 1000.0),),
                   kontotyp=typ, gruppe=gruppe, fristigkeit=frist)


# ---- die Datei selbst ----------------------------------------------------

def test_rahmen_laedt_vollstaendig(aasb):
    assert aasb.id == "aasb"
    assert aasb.version == "1.0"
    assert len(aasb.fs_lines) == 67
    assert len(aasb.bibliothek) == 209
    assert len(aasb.stichwortregeln) == 32
    assert sum(len(r.stichworte) for r in aasb.stichwortregeln) == 248


def test_jede_stichwortregel_kennt_ihr_fs_line_item(aasb):
    """Sonst zeigte eine Regel auf eine Position, die es nicht gibt — die
    Klasse fiele lautlos weg. ``Kontenrahmen._pruefe`` wirft beim Laden."""
    for r in aasb.stichwortregeln:
        assert r.fs_line in aasb.fs_lines
        assert aasb.fs_lines[r.fs_line].klasse == r.klasse


def test_working_capital_positionen_haben_eine_bilanzseite(aasb):
    """OA/OL ist die Seite des Working Capital. Ohne Seite keine Zuordnung."""
    for fs, reg in aasb.fs_lines.items():
        if reg.klasse in ("TWC", "OWC"):
            assert aasb.seite_von(fs) or fs in aasb.beidseitige_positionen


# ---- Reihenfolge der Kaskade ---------------------------------------------

def test_stichwort_schlaegt_bibliothek(aasb):
    """Der Kern der Reihenfolge. "CBA Main Cheque Acct" enthält nirgends
    "Bank current accounts" — die Bibliothek allein träfe das Konto nicht."""
    z = aasb.zuordnen("CBA Main Cheque Acct", "Current Assets")
    assert z.quelle == Kontenrahmen.QUELLE_STICHWORT
    assert z.fs_line == "Cash and cash equivalents"
    assert aasb._per_bibliothek(" cba main cheque acct ") is None


def test_laengster_treffer_gewinnt(aasb):
    """'petty cash' (10) schlägt 'cash' (4) — beide zeigen hier auf dieselbe
    Position, aber der Mechanismus muss stimmen, sonst entscheidet die
    Reihenfolge in der Datei."""
    z = aasb.zuordnen("Petty Cash Float")
    assert z.regel_id == "name:petty cash"


def test_kontogruppe_erst_wenn_der_name_nichts_liefert(aasb):
    """Die Gruppe nennt die Nachbarschaft eines Kontos, nicht seinen Inhalt.
    Sie darf einen Namenstreffer nie überschreiben."""
    z = aasb.zuordnen("Trade Debtors Control", "Property, plant & equipment")
    assert z.quelle == Kontenrahmen.QUELLE_STICHWORT
    assert z.fs_line == "Trade receivables, net"

    z = aasb.zuordnen("Waterloo Site 2", "Property, plant & equipment")
    assert z.quelle == Kontenrahmen.QUELLE_GRUPPE
    assert z.fs_line == "Property, plant and equipment"


def test_kein_treffer_ist_kein_notbehelf(aasb):
    """Ein unbrauchbarer Kontoname liefert None und damit die Review-Queue —
    nicht die nächstbeste Position."""
    assert aasb.zuordnen("efine", "") is None


# ---- Vorrang der Konzernregel --------------------------------------------

def test_konzernname_schlaegt_zinsregel(aasb):
    """Der dokumentierte Fehlfall: '1-90151 Accrued Interest - Aurora' landet
    über 'accrued interest' bei den Zinsabgrenzungen (ND), gehört aber zu den
    Konzernforderungen. Ohne erfassten Namen ist das nicht entscheidbar."""
    ohne = aasb.zuordnen("Accrued Interest - Aurora", "Current Assets")
    assert ohne.fs_line == "Accrued interest / borrowings"
    assert ohne.klasse == "ND"

    mit = aasb.zuordnen("Accrued Interest - Aurora", "Current Assets",
                        konzernnamen=["Aurora"])
    assert mit.fs_line == "Related-party receivables"
    assert mit.regel_id == "konzern:aurora"


def test_konzernseite_folgt_der_bilanzseite(aasb):
    z = aasb.zuordnen("Loan from Aurora Pty Ltd", "Current Liabilities",
                      konzernnamen=["Aurora"])
    assert z.fs_line == "Related-party payables"


# ---- Wortgrenze ----------------------------------------------------------

@pytest.mark.parametrize("bezeichnung", ["Encashment Fees", "Overcash Ltd"])
def test_stichwort_trifft_nur_am_wortanfang(aasb, bezeichnung):
    """Reiner Substring-Match ist die Fehlerquelle Nummer eins solcher
    Regelwerke ('kst' in 'Rückstellung', 'cash' in 'encashment')."""
    z = aasb.zuordnen(bezeichnung, "")
    assert z is None or z.fs_line != "Cash and cash equivalents"


def test_mehrzahl_bleibt_erhalten(aasb):
    """Rechts wird nicht geschnitten — sonst verlöre 'trade receivable' die
    Mehrzahl, und davon lebt eine Stichwortliste."""
    assert aasb.zuordnen("Trade Receivables").fs_line == "Trade receivables, net"


def test_sonderzeichen_stoeren_nicht(aasb):
    """'BT Logo (Capitalised)' und 'BT Logo Capitalised' sind derselbe Name."""
    assert (aasb.zuordnen("BT Logo (Capitalised)").fs_line
            == aasb.zuordnen("BT Logo Capitalised").fs_line == "Intangible assets")


# ---- Fristigkeit ---------------------------------------------------------

def test_fristigkeit_widerspruch_flaggt_statt_abzulehnen(aasb):
    z = aasb.zuordnen("Trade Debtors Control", "Current Assets",
                      fristigkeit="Non-current")
    assert z.fs_line == "Trade receivables, net"
    assert z.flags and "widerspricht" in z.flags[0]


def test_beidseitige_fristigkeit_erzeugt_keinen_fehlalarm(aasb):
    """Die Bibliothek führt 'Current / Non-current' in EINEM Feld. Ungetrennt
    gelesen widerspräche der Eintrag jeder Angabe der Quelle."""
    for frist in ("Current", "Non-current"):
        z = aasb.zuordnen("Provision for Annual Leave", "", fristigkeit=frist)
        assert z.flags == []


# ---- Setup-Dialog --------------------------------------------------------

def test_setup_waehlt_den_rahmen():
    assert setup(None, kontenrahmen="skr03").kontenrahmen is None
    assert setup(None, kontenrahmen="aasb").kontenrahmen.id == "aasb"


def test_unbekannter_kontenrahmen_ist_ein_fehler():
    """Kein stiller Rückfall auf HGB — sonst liefe ein IFRS-Mandat unbemerkt
    gegen die SKR-Bereichstabelle."""
    with pytest.raises(ValueError, match="Unbekannter Kontenrahmen"):
        setup(None, kontenrahmen="ifrs")


def test_setup_meldet_fehlende_konzernnamen():
    s = setup(None, kontenrahmen="aasb")
    assert "keine verbundenen Unternehmen erfasst" in s.rahmen_meldung


# ---- Engine: beide Welten berühren einander nicht ------------------------

@pytest.fixture(scope="module")
def hc() -> Hausconvention:
    return Hausconvention.laden()


def test_hgb_mandat_bleibt_unveraendert(hc):
    """Ohne Kontenrahmen läuft ausschließlich die HGB-Kette — SKR-Default
    inklusive."""
    m = Engine(hc).map_account(konto("Betriebsausstattung", "410"), False)
    assert m.rahmen == "HGB"
    assert m.hgb_pfad.startswith("/Aktiva")
    assert m.klasse == Klasse.FA


def test_aasb_mandat_umgeht_die_skr_bereichstabelle(hc, aasb):
    """'1000 BT Imaging Product' fiel bisher in den SKR03-Bereich 1000–1299
    und damit auf 'Kassenbestand'. Ein australisches Konto hat mit SKR-Nummern
    nichts zu tun."""
    a = konto("BT Imaging Product", "1000", typ="bilanz_aktiv")
    hgb = Engine(hc).map_account(a, False)
    assert hgb.klasse == Klasse.ND and hgb.na_de == "Liquide Mittel"

    m = Engine(hc, kontenrahmen=aasb).map_account(a, False)
    assert m.rahmen == "aasb"
    assert m.klasse == Klasse.TWC and m.na_de == "Inventories"


def test_pfad_traegt_den_rahmen_im_ersten_segment(hc, aasb):
    """Ein FS Line Item ist kein HGB-Pfad und darf nicht so aussehen. Die
    Bilanzseite fragt man über ``bilanzseite``."""
    m = Engine(hc, kontenrahmen=aasb).map_account(
        konto("Trade Debtors Control", "1-12000", typ="bilanz_aktiv"), False)
    assert m.hgb_pfad == "/AASB/Aktiva/Trade receivables, net"
    assert m.bilanzseite == "AKTIVA"
    assert m.seite == "OA"


def test_quellenspalte_unterscheidet_die_drei_stufen(hc, aasb):
    """Ohne die Unterscheidung ist eine Fehlzuordnung im Mastersheet später
    nicht auffindbar."""
    eng = Engine(hc, kontenrahmen=aasb)
    stichwort = eng.map_account(konto("CBA Main Cheque Acct", typ="bilanz_aktiv"), False)
    gruppe = eng.map_account(
        konto("Waterloo Site 2", typ="bilanz_aktiv",
              gruppe="Property, plant & equipment"), False)
    assert stichwort.quelle == Quelle.AASB_STICHWORT
    assert gruppe.quelle == Quelle.AASB_GRUPPE
    assert Quelle.AASB_BIBLIOTHEK.value == "AASB-Bibliothek"


def test_kontennachweis_bleibt_massgeblich(hc, aasb):
    """Stufe 1 der Kaskade gilt in beiden Welten: liefert der Abschluss die
    Position, entscheidet nicht das Stichwort."""
    a = Account(konto="1-10100", bezeichnung="CBA Main Cheque Acct",
                salden=(PeriodBalance("FY2024", 10.0),),
                kontotyp="bilanz_aktiv", fs_pfad="Restricted cash / other current assets")
    m = Engine(hc, kontenrahmen=aasb).map_account(a, True)
    assert m.quelle == Quelle.KONTENNACHWEIS
    assert m.na_de == "Restricted cash / other current assets"


def test_gemischte_position_ohne_typ2_regel_geht_in_review(hc, aasb):
    """18 der 67 FS Line Items sind gemischt. Ob eine Rückstellung Net Debt
    oder Working Capital ist, sagt kein Kontenrahmen — nur der Inhalt. Greift
    keine Typ-2-Regel, wird das nicht geraten."""
    m = Engine(hc, kontenrahmen=aasb).map_account(
        konto("Provision for Annual Leave", "2-80400", typ="bilanz_passiv"), False)
    assert m.na_de == "Provisions"
    assert m.review and m.klasse == Klasse.REVIEW
