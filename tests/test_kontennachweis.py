"""Teil 3: Kontennachweis als vorrangige Strukturquelle, Ergänzung fehlender
Konten, Reconciliation und Setup-Modus."""

import pytest

from conftest import HUCHTEMEIER, ECKART, datei
from fdd.core.hausconvention import Hausconvention
from fdd.core.model import NormalizedLedger, Quelle
from fdd.engine.cascade import Engine
from fdd.engine.kontennachweis_apply import wende_kontennachweis_an
from fdd.engine.reconciliation import reconcile
from fdd.engine.setup import setup
from fdd.readers.kontennachweis import lies_kontennachweis
from fdd.readers.pdf_kontennachweis import PdfKontennachweisReader


@pytest.fixture(scope="module")
def kn():
    return lies_kontennachweis(datei(HUCHTEMEIER))


def test_kn_liest_konten_und_positionen(kn):
    assert len(kn.konten) > 150
    assert len(kn.positionen()) > 15


@pytest.mark.parametrize("konto,erwartet_teil", [
    ("0027", "Immaterielle Vermoegensgegenstaende"),
    ("0210", "Technische Anlagen und Maschinen"),
    ("0530", "Wertpapiere des Anlagevermoegens"),   # trotz Wasserzeichen
    ("3980", "Fertige Erzeugnisse und Waren"),      # trotz 'WJaren'
    ("1400", "Forderungen aus Lieferungen und Leistungen"),
    ("1600", "Verbindlichkeiten aus Lieferungen und Leistungen"),
])
def test_kn_ordnet_konten_der_richtigen_position_zu(kn, konto, erwartet_teil):
    assert konto in kn.konten, f"{konto} nicht im Kontennachweis"
    assert erwartet_teil in kn.konten[konto].hgb_pfad


def test_kontennachweis_schlaegt_skr_default(kn):
    """Mit Kontennachweis kommt der HGB-Pfad aus Stufe 1 — nicht aus dem
    SKR-Default."""
    led = PdfKontennachweisReader().lesen(datei(HUCHTEMEIER))
    led = wende_kontennachweis_an(led, kn)
    mapped = Engine(Hausconvention.laden()).map_ledger(led)
    quellen = {m.quelle for m in mapped}
    assert Quelle.KONTENNACHWEIS in quellen
    assert Quelle.SKR_DEFAULT not in quellen, (
        "SKR-Default darf bei vorhandenem Kontennachweis nicht greifen")


def test_fehlende_konten_werden_ergaenzt_und_recon_zeigt_luecke(kn):
    """Konten, die die SuSa nicht führt, kommen über den Kontennachweis ins
    Databook — und die Reconciliation macht die Lücke sichtbar."""
    voll = PdfKontennachweisReader().lesen(datei(HUCHTEMEIER))
    fehlen = {k for k in ("0950", "1400", "1600")
              if k in {a.konto for a in voll.accounts} and k in kn.konten}
    assert len(fehlen) >= 2

    beschnitten = NormalizedLedger(
        accounts=[a for a in voll.accounts if a.konto not in fehlen],
        perioden=voll.perioden, entity=voll.entity, quelle_datei=voll.quelle_datei,
        hat_kontennachweis=False, fingerprint=voll.fingerprint, warnungen=[])
    susa_konten = {a.konto for a in beschnitten.accounts}

    ergaenzt = wende_kontennachweis_an(beschnitten, kn)
    assert len(ergaenzt.accounts) == len(beschnitten.accounts) + len(fehlen)

    mapped = Engine(Hausconvention.laden()).map_ledger(ergaenzt)
    by = {m.konto: m for m in mapped}
    for k in fehlen:
        assert by[k].quelle == Quelle.KONTENNACHWEIS
        assert by[k].hgb_pfad.startswith(("/Aktiva", "/Passiva"))

    rec = reconcile(mapped, kn, voll.perioden, susa_konten)
    assert set(rec.nur_im_kn) == fehlen
    # Die betroffenen Positionen tragen eine Differenz in Höhe der Lücke
    p = voll.perioden[0]
    betroffen = {by[k].hgb_pfad for k in fehlen}
    diff_pfade = {z.hgb_pfad for z in rec.zeilen_mit_differenz}
    assert betroffen <= diff_pfade
    for k in fehlen:
        z = next(z for z in rec.zeilen if z.hgb_pfad == by[k].hgb_pfad)
        assert abs(z.differenz(p)) >= abs(by[k].saldo(p)) - 0.01


def test_setup_modus_ohne_kontennachweis_ist_vorlaeufig():
    erg = setup(None, 215, 0)
    assert erg.modus == "vorlaeufig"
    assert not erg.ist_abschlusstreu
    assert "NICHT ABSCHLUSSTREU" in erg.databook_kennzeichen
    assert erg.anforderung, "Der Kontennachweis muss aktiv angefordert werden"


def test_setup_modus_mit_kontennachweis_ist_abschlusstreu():
    erg = setup("kn.pdf", 100, 100)
    assert erg.ist_abschlusstreu
    assert erg.abdeckung == pytest.approx(1.0)
    assert "NICHT ABSCHLUSSTREU" not in erg.databook_kennzeichen


def test_eckart_laeuft_ohne_kn_im_default_modus():
    """Ohne Kontennachweis bleibt Eckart im ausdrücklichen Default-Modus."""
    from fdd.cli import run
    res = run(datei(ECKART), "/tmp/eckart_probe.xlsx", verbose=False)
    assert res["setup"].modus == "vorlaeufig"
    assert res["recon"] is None
