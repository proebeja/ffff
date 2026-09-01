"""Die Mechaniken, die Hausconvention v2.8 neu verlangt.

Vier Bausteine, die alle nach dem Mapping ansetzen und es **nicht** ersetzen:

* ``seitenwechsel``      — periodenabhängige Bilanzseite für definierte
  Pfadpaare. Die Klasse bleibt fest, nur der Pfad folgt dem Vorzeichen.
* ``verhaltenspruefung`` — Querschnitt über alle Konten mit Salden in mehreren
  Perioden; erzeugt Flags und Fragen, nie eine Umklassifizierung.
* ``saldenvortrag``      — Auflaufweg der DATEV-Vortragskonten: abstimmen,
  Rest ins Eigenkapital, nie stillschweigend in TECH.
* ``vorlaeufige_pfade``  — QA A6: jedes Konto landet in einer Bilanzposition;
  was wirklich keinen Pfad hat, wird eine sichtbare Zeile *innerhalb* der
  Bilanz statt eines Ausgleichspostens daneben.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import Optional

from ..core.model import Klasse, MappedAccount

NICHT_ZUGEORDNET = "/Aktiva/Z Noch nicht zugeordnet"


# ---- 1) Seitenwechsel ----------------------------------------------------
@dataclass
class SeitenwechselFall:
    konto: str
    bezeichnung: str
    basispfad: str
    klasse: str
    #: Periode -> abweichender Pfad
    abweichend: dict[str, str] = field(default_factory=dict)
    salden: dict[str, float] = field(default_factory=dict)

    @property
    def pflichtfrage(self) -> str:
        return (f"Konto {self.konto} wechselt über die Perioden die Bilanzseite "
                f"({', '.join(f'{p}: {self.salden[p]:,.2f}' for p in self.salden)}). "
                "Hintergrund des Wechsels? Die Umgliederung ist mechanisch "
                "korrekt, der Sachverhalt dahinter ist es nicht automatisch.")


def _na_fuer_pfad(hc, pfad: str) -> Optional[tuple[str, str]]:
    """Net-Asset-Zeile der Gegenposition — aus der Reklassifizierung, damit die
    umgegliederte Periode in der richtigen Lead-Zeile landet."""
    for r in hc.reklass_regeln:
        if r.hgb_pfad == pfad:
            return (r.na_de, r.na_en)
    return None


def _pfadpaare(hc) -> list[tuple[str, str]]:
    return [(p["aktiv"], p["passiv"])
            for p in hc._d.get("seitenwechsel", {}).get("pfadpaare", [])]


def wende_seitenwechsel_an(mapped: list[MappedAccount], perioden: list[str],
                           hc, nachgewiesene_seite: Optional[dict] = None
                           ) -> tuple[list[MappedAccount], list[SeitenwechselFall]]:
    """Setzt je Periode die Bilanzseite nach dem Vorzeichen — nur innerhalb der
    definierten Pfadpaare. Ausserhalb bleibt der Pfad fest; ein Vorzeichen-
    wechsel dort ist ein Review-Fall der Verhaltensprüfung, keine Umgliederung.

    Die Klasse bleibt unangetastet: ein Cash-Pool bleibt ND, ob er nun
    Forderung oder Verbindlichkeit ist.

    ``nachgewiesene_seite`` ist die Seite, die der **Kontennachweis** je Konto
    und Periode ausweist. Wo er spricht, gilt er: die Vorzeichenableitung ist
    ein Ersatz für fehlende Evidenz, kein Ersatz für den Abschluss. Konto
    1789 0 wechselt zwar das Vorzeichen, der Abschluss führt es aber in beiden
    Jahren im Umsatzsteuerblock der sonstigen Verbindlichkeiten — es wird
    deshalb nicht umgegliedert."""
    paare = _pfadpaare(hc)
    if not paare:
        return mapped, []
    aktiv_zu_passiv = {a: p for a, p in paare}
    passiv_zu_aktiv = {p: a for a, p in paare}

    faelle: list[SeitenwechselFall] = []
    neu: list[MappedAccount] = []
    for m in mapped:
        gegen = aktiv_zu_passiv.get(m.hgb_pfad) or passiv_zu_aktiv.get(m.hgb_pfad)
        if gegen is None:
            neu.append(m)
            continue
        # Voraussetzung ist ein echter Wechsel ÜBER die Perioden — dasselbe
        # Kriterium wie verhaltenspruefung.kriterien.vorzeichenwechsel, auf das
        # der Block ausdrücklich verweist.
        #
        # Ein Konto, das in ALLEN Perioden dasselbe Vorzeichen trägt, wechselt
        # nichts: es steht dauerhaft dort, wo der Abschluss es ausweist. Die
        # Vorsteuerkonten 1571–1577 etwa haben durchgehend Sollcharakter und
        # gehören im Abschluss trotzdem in den Umsatzsteuerblock der sonstigen
        # Verbindlichkeiten. Würde man auch sie umgliedern, liefe die
        # Abstimmung gegen den Abschluss um genau diesen Betrag auseinander.
        belegt = {p: m.saldo(p) for p in perioden if abs(m.saldo(p)) > 0.005}
        if len({v > 0 for v in belegt.values()}) < 2:
            neu.append(m)
            continue
        ist_aktiv_basis = m.hgb_pfad in aktiv_zu_passiv
        nachgewiesen = (nachgewiesene_seite or {}).get(m.konto, {})
        abweichend: dict[str, str] = {}
        for p, saldo in belegt.items():
            seite = nachgewiesen.get(p)
            gehoert_aktiv = (seite == "AKTIVA") if seite else (saldo > 0)
            if gehoert_aktiv != ist_aktiv_basis:
                abweichend[p] = gegen
        if not abweichend:
            neu.append(m)
            continue
        fall = SeitenwechselFall(
            konto=m.konto, bezeichnung=m.bezeichnung, basispfad=m.hgb_pfad,
            klasse=m.klasse.value, abweichend=abweichend,
            salden={p: m.saldo(p) for p in perioden if abs(m.saldo(p)) > 0.005})
        faelle.append(fall)
        na_abw = {p: _na_fuer_pfad(hc, z) for p, z in abweichend.items()}
        neu.append(replace(m, pfad_je_periode=dict(abweichend),
                           na_je_periode={p: n for p, n in na_abw.items() if n},
                           begruendung=(
            m.begruendung + " Seitenwechsel v2.8: die Bilanzseite folgt je "
            "Periode dem Vorzeichen, die Klasse bleibt fest.")))
    return neu, faelle


# ---- 2) Verhaltensprüfung (Querschnitt) ----------------------------------
@dataclass
class VerhaltensBefund:
    konto: str
    bezeichnung: str
    klasse: str
    kriterium: str            # "vorzeichenwechsel" | "nicht wiederkehrend" | "schwankung"
    wesentlich: bool
    salden: dict[str, float]
    hinweis: str

    @property
    def ist_frage(self) -> bool:
        return self.wesentlich


def pruefe_verhalten(mapped: list[MappedAccount], perioden: list[str], hc,
                     schwelle: float) -> list[VerhaltensBefund]:
    """Querschnittsprüfung über ALLE Konten mit Salden in mehreren Perioden.

    Sie entscheidet nie über die Klasse — sie markiert Auffälligkeiten und
    erzeugt oberhalb der Wesentlichkeitsschwelle eine Rückfrage. Der
    Vorzeichenwechsel gilt unabhängig von der Bandbreite als auffällig, weil
    dahinter regelmäßig ein Sachverhalt steckt."""
    vp = hc._d.get("verhaltenspruefung", {})
    bandbreite = vp.get("kriterien", {}).get("bandbreite_prozent_default", 50) / 100.0
    befunde: list[VerhaltensBefund] = []

    for m in mapped:
        if m.klasse in (Klasse.PL, Klasse.TECH):
            continue
        werte = {p: m.saldo(p) for p in perioden}
        belegt = {p: v for p, v in werte.items() if abs(v) > 0.005}
        if len(belegt) < 2:
            if len(belegt) == 1 and abs(next(iter(belegt.values()))) > 0.005:
                befunde.append(VerhaltensBefund(
                    m.konto, m.bezeichnung, m.klasse.value, "nicht wiederkehrend",
                    max(abs(v) for v in belegt.values()) >= schwelle, belegt,
                    "Saldo nur in einer Periode — Kandidat für die "
                    "WC-Normalisierung, keine Umklassifizierung."))
            continue

        vorzeichen = {p: (v > 0) for p, v in belegt.items()}
        if len(set(vorzeichen.values())) > 1:
            befunde.append(VerhaltensBefund(
                m.konto, m.bezeichnung, m.klasse.value, "vorzeichenwechsel",
                max(abs(v) for v in belegt.values()) >= schwelle, belegt,
                "Bestandskonto wechselt über die Perioden das Vorzeichen "
                "(aktivisch/passivisch). Gilt unabhängig von der Bandbreite "
                "als auffällig."))
            continue

        mittel = sum(abs(v) for v in belegt.values()) / len(belegt)
        if mittel > 0 and any(abs(abs(v) - mittel) > bandbreite * mittel
                              for v in belegt.values()):
            befunde.append(VerhaltensBefund(
                m.konto, m.bezeichnung, m.klasse.value, "schwankung",
                max(abs(v) for v in belegt.values()) >= schwelle, belegt,
                f"Schwankung ausserhalb der Bandbreite von "
                f"{bandbreite:.0%} um den Mittelwert."))
    return befunde


# ---- 3) Saldenvorträge ---------------------------------------------------
@dataclass
class SaldenvortragsWeg:
    konto: str
    bezeichnung: str
    salden: dict[str, float]
    abgestimmt_gegen: str = ""
    rest: dict[str, float] = field(default_factory=dict)
    ziel: str = ""
    begruendung: str = ""


#: Saldenvortragskonto -> Sammelkonto derselben Periode, gegen das abgestimmt wird.
_ABSTIMMPAARE = {"9008": "1400", "9009": "1600"}
_EK_VORTRAG = "/Passiva/A Eigenkapital/IV Gewinnvortrag Verlustvortrag"


def loese_saldenvortraege(mapped: list[MappedAccount], perioden: list[str], hc
                          ) -> tuple[list[MappedAccount], list[SaldenvortragsWeg]]:
    """Auflaufweg nach v2.8: Debitoren-/Kreditorenvortrag gegen die
    Sammelkonten abstimmen; der nicht abstimmbare Rest ist regelmäßig der
    Vortrag des Vorjahresergebnisses und geht ins Eigenkapital. Was danach
    bleibt, wird eine sichtbare Zeile — nie TECH."""
    wege: list[SaldenvortragsWeg] = []
    neu: list[MappedAccount] = []
    for m in mapped:
        if not hc.ist_saldenvortrag(m.konto):
            neu.append(m)
            continue
        nummer = m.konto.split()[0]
        salden = {p: m.saldo(p) for p in perioden if abs(m.saldo(p)) > 0.005}
        weg = SaldenvortragsWeg(konto=m.konto, bezeichnung=m.bezeichnung,
                                salden=salden)
        gegen = _ABSTIMMPAARE.get(nummer)
        if gegen:
            weg.abgestimmt_gegen = f"Sammelkonto {gegen}"
            weg.ziel = _EK_VORTRAG
            weg.begruendung = (
                f"Vortrag der {'Debitoren' if gegen == '1400' else 'Kreditoren'}; "
                f"gegen das Sammelkonto {gegen} derselben Periode zu stellen. "
                "Der Saldo bleibt bis zur Klärung im Eigenkapital sichtbar "
                "statt in TECH zu verschwinden.")
        else:
            weg.ziel = _EK_VORTRAG
            weg.begruendung = ("Vortrag der Sachkonten — regelmäßig das noch "
                               "nicht auf die Eigenkapitalkonten gebuchte "
                               "Vorjahresergebnis.")
        weg.rest = dict(salden)
        wege.append(weg)
        neu.append(replace(
            m, hgb_pfad=_EK_VORTRAG, klasse=Klasse.EQ,
            na_de="Eigenkapital", na_en="Equity", review=True,
            begruendung=(weg.begruendung + " (v2.8: Saldenvortragskonten sind "
                         "keine Statistikkonten.)")))
    return neu, wege


# ---- 4) QA A6: vorläufige Pfade -----------------------------------------
@dataclass
class UngeloestesKonto:
    konto: str
    bezeichnung: str
    salden: dict[str, float]
    grund: str


#: Vorläufige Klasse, wenn die Position sie nicht eindeutig hergibt. Der
#: operative Teil ist die konservative Annahme — sie hält den Betrag im
#: Working Capital statt ihn ins Net Debt zu heben.
_VORLAEUFIG_MIXED = Klasse.OWC


def _klasse_aus_pfad(pfad: str, hc) -> Optional[Klasse]:
    for r in hc.reklass_regeln:
        if r.hgb_pfad == pfad:
            return _VORLAEUFIG_MIXED if r.klasse == "MIXED" else Klasse(r.klasse)
    return None


def _vorlaeufige_seite(pfad: str, klasse: Klasse) -> Optional[str]:
    """OA/OL auch für die vorläufig gesetzte Klasse ableiten.

    ``oa_ol_ableitung`` verlangt die Seite deterministisch aus der Bilanzseite
    des Pfads. Sie hier zu vergessen, fällt lange nicht auf: die Klasse stimmt,
    der Lead im alten Hausformat gruppiert nach Klasse, und erst die
    Dealtool-Vorlage fragt nach ``OA``/``OL`` als Ticker. Im Luma-Lauf standen
    dadurch 34 Konten im Mastersheet, ohne in einer Position des Lead NA zu
    landen.
    """
    if klasse not in (Klasse.TWC, Klasse.OWC):
        return None
    return "OA" if pfad.startswith("/Aktiva") else "OL"


def setze_vorlaeufige_pfade(mapped: list[MappedAccount], perioden: list[str],
                            hc=None
                            ) -> tuple[list[MappedAccount], list[UngeloestesKonto]]:
    """QA A6: kein Konto hängt ausserhalb der Bilanz.

    Konten ohne bestimmbaren Pfad bekommen die sichtbare Position "noch nicht
    zugeordnet" INNERHALB der Bilanz. Der Review-Status bleibt zusätzlich
    bestehen — die Zuordnung ersetzt die Klärung nicht, sie verhindert nur,
    dass der Betrag als Ausgleichsposten neben der Bilanz hängt."""
    offen: list[UngeloestesKonto] = []
    neu: list[MappedAccount] = []
    for m in mapped:
        if m.klasse == Klasse.TECH:
            neu.append(m)
            continue

        hat_pfad = bool(m.hgb_pfad) and not m.hgb_pfad.startswith("(")
        if hat_pfad and m.klasse == Klasse.REVIEW and hc is not None:
            # Pfad da, Klasse offen: das Konto bekommt die vorläufige Klasse
            # seiner Position und läuft damit in den Lead. Der Review-Status
            # bleibt daneben bestehen — die Zuordnung ersetzt die Klärung
            # nicht, sie verhindert nur, dass der Betrag ausserhalb der Bilanz
            # hängt (QA A6).
            vorlaeufig = _klasse_aus_pfad(m.hgb_pfad, hc)
            if vorlaeufig is not None:
                neu.append(replace(m, klasse=vorlaeufig, review=True,
                                   seite=_vorlaeufige_seite(m.hgb_pfad,
                                                            vorlaeufig),
                                   begruendung=(m.begruendung + " QA A6: "
                                                f"vorläufig als {vorlaeufig.value} "
                                                "in der Position geführt, Klärung "
                                                "offen.")))
                continue
        if hat_pfad:
            neu.append(m)
            continue
        salden = {p: m.saldo(p) for p in perioden if abs(m.saldo(p)) > 0.005}
        offen.append(UngeloestesKonto(m.konto, m.bezeichnung, salden,
                                      m.begruendung or "kein Pfad bestimmbar"))
        neu.append(replace(
            m, hgb_pfad=NICHT_ZUGEORDNET, klasse=Klasse.OWC,
            na_de="Noch nicht zugeordnet", na_en="Not yet allocated",
            review=True,
            begruendung=(m.begruendung + " QA A6: vorläufig in einer sichtbaren "
                         "Bilanzzeile geführt, damit der Betrag nicht ausserhalb "
                         "der Bilanz hängt. Klärung offen.")))
    return neu, offen
