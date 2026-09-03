"""Kontenrahmen für Nicht-HGB-Abschlüsse — erster Rahmen AASB.

Für HGB läuft die Kette ``Konto -> HGB-Pfad -> Klasse``. Den HGB-Pfad liefern
Kontennachweis, Typ-1-Regeln oder die SKR-Bereichstabelle; die Klasse fällt
danach aus der Reklassifizierungstabelle. Für AASB, IFRS, US GAAP und UK GAAP
fehlte dieses Fundament. Ohne es vergibt die Software keinen Bilanzposten und
behilft sich mit der Kontogruppe der Saldenliste — so stand in einem Lauf
"R&D/Demo tools - Waterloo" als Position im Lead NA.

Dieses Modul schließt das. Der Aufbau ist derselbe wie bei HGB, mit zwei
Unterschieden:

**Der Treffer läuft über den Kontonamen, nicht über die Kontonummer.** AASB
kennt keine verbindliche Nummerierung; eine Bereichstabelle wie
``skr03_default_bereiche`` wäre wertlos, weil jeder Mandant eigene Nummern
vergibt.

**Statt HGB-Pfaden gibt es FS Line Items.** Also ``Trade and other payables``
statt ``/Passiva/Verbindlichkeiten/Verbindlichkeiten aus Lieferungen und
Leistungen``.

Reihenfolge der Kaskade
-----------------------

1. Kontennachweis des Abschlusses (in der Engine, nicht hier).
2. Konzernregel: die zu Mandatsbeginn erfassten Namen der verbundenen
   Unternehmen. Sie geht den übrigen Stichworten **vor** — sonst landet
   ``Accrued Interest - Aurora`` über ``accrued interest`` bei den
   Zinsabgrenzungen statt bei den Konzernforderungen.
3. Stichwortregeln gegen den Kontonamen, längster Treffer gewinnt.
4. Kontenbibliothek gegen den Kontonamen (Feinabgleich).
5. Stichwortregeln gegen die Kontogruppe.
6. Kategorie/Unterkategorie der Bibliothek gegen die Kontogruppe.
7. Kein Treffer -> Review-Queue.

Die Reihenfolge ist der Kern und nicht verhandelbar: die Bibliothek allein
traf im Test gegen 220 australische Konten 9 Prozent, weil Bankkonten nach der
Bank heißen und nicht nach ihrer Funktion. "CBA Main Cheque Acct" enthält
nirgends "Bank current accounts". Die Bibliothek bleibt trotzdem, weil sie
definiert, welche FS Line Items überhaupt zulässig sind.

**Widerspruch in der gelieferten Datei.** Ihr Block ``_kaskade`` führt die
Bibliothek auf Stufe 3 und kennt die Stichwortregeln erst danach für MIXED.
Ihr eigener Block ``_matching`` sagt das Gegenteil ("REIHENFOLGE: erst die
Stichwortregeln, dann die Kontenbibliothek als Feinabgleich"), ebenso die
Übergabe und der Auftrag. Umgesetzt ist ``_matching``; ``_kaskade`` ist in der
Datei zu korrigieren.
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from typing import Iterable, Optional

_CONFIG = os.path.join(os.path.dirname(__file__), "..", "config")

#: Nicht-Alphanumerisches wird zu Leerzeichen — "BT Logo (Capitalised)" und
#: "BT Logo Capitalised" sind derselbe Kontoname.
_UNWORT = re.compile(r"[^a-z0-9]+")


def normalisiere(text: object) -> str:
    """Klein schreiben, Umlaute auflösen, Sonderzeichen zu Leerzeichen.

    Das Ergebnis ist mit je einem Leerzeichen umschlossen, damit ein Stichwort
    am Wortanfang geprüft werden kann (siehe :func:`_trifft`).
    """
    if text is None:
        return " "
    t = str(text).lower()
    t = (t.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue")
           .replace("ß", "ss"))
    return " " + _UNWORT.sub(" ", t).strip() + " "


def _trifft(stichwort_norm: str, text_norm: str) -> bool:
    """Substring-Treffer **am Wortanfang**.

    Reiner Substring-Match ist die Fehlerquelle Nummer eins dieser Regelwerke:
    ``kst`` trifft in "Rückstellung", ``cash`` in "encashment". Verlangt wird
    deshalb, dass das Stichwort links an einer Wortgrenze beginnt. Rechts wird
    nicht geschnitten — sonst verlöre "receivable" die Mehrzahl
    "receivables", und genau davon lebt eine Stichwortliste.
    """
    stichwort = stichwort_norm.strip()
    if not stichwort:
        return False
    return (" " + stichwort) in text_norm


@dataclass(frozen=True)
class FSLine:
    """Ein FS Line Item der Reklassifizierungstabelle."""

    fs_line: str
    klasse: str                 # FA/TWC/OWC/ND/EQ/DT/MIXED
    deal_hinweis: str = ""


@dataclass(frozen=True)
class Musterkonto:
    """Ein Eintrag der Kontenbibliothek."""

    name: str
    fs_line: str
    statement: str = ""
    current: str = ""
    kategorie: str = ""
    unterkategorie: str = ""
    normal: str = ""
    name_norm: str = ""
    kategorie_norm: str = ""
    unterkategorie_norm: str = ""


@dataclass(frozen=True)
class Stichwortregel:
    """Eine Stichwortregel: n Stichworte, ein FS Line Item als Ziel."""

    fs_line: str
    klasse: str
    stichworte: tuple[str, ...]          # normalisiert


@dataclass
class Zuordnung:
    """Was der Rahmen zu einem Konto sagt — inklusive Herkunft."""

    fs_line: str
    klasse: str
    quelle: str                  # "AASB-Stichwort" | "AASB-Bibliothek" | "AASB-Gruppe"
    regel_id: str
    begruendung: str
    seite: Optional[str] = None  # "AKTIVA" | "PASSIVA" | None (GuV)
    current: str = ""
    flags: list[str] = field(default_factory=list)


class Kontenrahmen:
    """Objektschnittstelle auf eine ``kontenrahmen_*.json``.

    Die Datei ist Konfiguration, kein Code. Dieses Modul hartkodiert kein
    Stichwort und kein FS Line Item.
    """

    QUELLE_STICHWORT = "AASB-Stichwort"
    QUELLE_BIBLIOTHEK = "AASB-Bibliothek"
    QUELLE_GRUPPE = "AASB-Gruppe"

    def __init__(self, data: dict, rahmen_id: str = "aasb"):
        self.id = rahmen_id
        self.name: str = data.get("_name", rahmen_id)
        self.version: str = data.get("_version", "?")
        self.quelle: str = data.get("_quelle", "")

        self.fs_lines: dict[str, FSLine] = {
            r["fs_line"]: FSLine(r["fs_line"], r["klasse"],
                                 r.get("deal_hinweis", ""))
            for r in data.get("reklassifizierung", [])
        }
        self.bibliothek: list[Musterkonto] = [
            Musterkonto(
                name=e.get("name", ""), fs_line=e.get("fs_line", ""),
                statement=e.get("statement", ""), current=e.get("current", ""),
                kategorie=e.get("kategorie", ""),
                unterkategorie=e.get("unterkategorie", ""),
                normal=e.get("normal", ""),
                name_norm=normalisiere(e.get("name", "")),
                kategorie_norm=normalisiere(e.get("kategorie", "")),
                unterkategorie_norm=normalisiere(e.get("unterkategorie", "")),
            )
            for e in data.get("kontenbibliothek", [])
        ]
        self.stichwortregeln: list[Stichwortregel] = [
            Stichwortregel(
                fs_line=r["fs_line"], klasse=r["klasse"],
                stichworte=tuple(normalisiere(s).strip()
                                 for s in r.get("wenn_eines", [])),
            )
            for r in data.get("stichwortregeln", [])
        ]
        #: FS Line Items, die je nach Saldo auf beiden Bilanzseiten stehen.
        #: Bei ihnen muss die Seite aus der Quelle kommen.
        self._beidseitig: set[str] = set()
        self._seite_je_fs_line = self._seiten_ableiten()
        self._statement_je_fs_line = {
            m.fs_line: m.statement for m in self.bibliothek if m.statement
        }
        self._pruefe()

    # ---- Fabrik ----------------------------------------------------------
    @classmethod
    def laden(cls, dateiname: str, rahmen_id: str = "aasb") -> "Kontenrahmen":
        pfad = (dateiname if os.path.isabs(dateiname)
                else os.path.join(_CONFIG, dateiname))
        with open(pfad, encoding="utf-8") as f:
            return cls(json.load(f), rahmen_id)

    # ---- Ableitungen beim Laden ------------------------------------------
    def _seiten_ableiten(self) -> dict[str, str]:
        """Bilanzseite je FS Line Item aus der Spalte ``normal`` der Bibliothek.

        Die Reklassifizierungstabelle nennt keine Seite. Sie steckt in der
        Bibliothek: ``Debit`` ist Aktiva, ``Credit`` Passiva. Gegenkonten
        (kumulierte Abschreibung, Wertberichtigung) stehen auf derselben
        Position mit umgekehrtem ``normal`` — deshalb entscheidet die
        Mehrheit, nicht der erste Eintrag.
        """
        stimmen: dict[str, dict[str, int]] = {}
        for m in self.bibliothek:
            if m.statement and "Balance Sheet" not in m.statement \
                    and m.statement != "Equity":
                continue
            seite = ("AKTIVA" if m.normal == "Debit"
                     else "PASSIVA" if m.normal == "Credit" else None)
            if seite is None:
                # ``Debit / Credit``: die Position steht je nach Saldo auf
                # beiden Seiten (GST receivable/payable, Verrechnungskonten).
                # Ihre Seite kann kein Kontenrahmen sagen, nur die Quelle.
                self._beidseitig.add(m.fs_line)
                continue
            stimmen.setdefault(m.fs_line, {}).setdefault(seite, 0)
            stimmen[m.fs_line][seite] += 1
        return {fs: max(s, key=lambda k: s[k]) for fs, s in stimmen.items()}

    def _pruefe(self) -> None:
        """Grobvalidierung. Ein Konfigfehler soll beim Laden auffallen und
        nicht erst als stille Fehlzuordnung im Mandat."""
        fehler = []
        for r in self.stichwortregeln:
            if r.fs_line not in self.fs_lines:
                fehler.append(f"Stichwortregel zeigt auf unbekanntes FS Line "
                              f"Item '{r.fs_line}'.")
            elif self.fs_lines[r.fs_line].klasse != r.klasse:
                fehler.append(
                    f"Stichwortregel '{r.fs_line}' trägt Klasse {r.klasse}, "
                    f"die Reklassifizierung {self.fs_lines[r.fs_line].klasse}.")
        for fs, reg in self.fs_lines.items():
            if (reg.klasse in ("TWC", "OWC")
                    and fs not in self._seite_je_fs_line
                    and fs not in self._beidseitig):
                fehler.append(f"Working-Capital-Position '{fs}' ohne "
                              f"ableitbare Bilanzseite.")
        if fehler:
            raise ValueError(f"{self.name} {self.version}: "
                             + " ".join(fehler))

    # ---- Zugriffe --------------------------------------------------------
    def klasse_von(self, fs_line: str) -> Optional[str]:
        reg = self.fs_lines.get(fs_line)
        return reg.klasse if reg else None

    def seite_von(self, fs_line: str) -> Optional[str]:
        return self._seite_je_fs_line.get(fs_line)

    @property
    def beidseitige_positionen(self) -> set[str]:
        """FS Line Items, deren Bilanzseite nur die Quelle sagen kann."""
        return set(self._beidseitig)

    def ist_guv(self, fs_line: str) -> bool:
        """FS Line Items der Bibliothek ohne Reklass-Eintrag sind GuV-Zeilen —
        die Reklassifizierungstabelle deckt nur die Bilanz ab."""
        if fs_line in self.fs_lines:
            return False
        stmt = self._statement_je_fs_line.get(fs_line, "")
        return "Income Statement" in stmt or "OCI" in stmt

    # ---- die Kaskade -----------------------------------------------------
    def zuordnen(self, bezeichnung: str, gruppe: str = "",
                 fristigkeit: str = "", seite: Optional[str] = None,
                 konzernnamen: Iterable[str] = ()) -> Optional[Zuordnung]:
        """Kontoname (und ersatzweise Kontogruppe) -> FS Line Item + Klasse.

        ``seite`` ist die Bilanzseite aus der Quelle, falls sie eine liefert
        (MYOB: "Current Assets"). Sie entscheidet bei der Konzernregel, ob
        Forderung oder Verbindlichkeit, und sticht die aus der Bibliothek
        abgeleitete Seite. ``fristigkeit`` ist das Zusatzsignal
        Current/Non-current; ein Widerspruch zur Bibliothek erzeugt einen
        Flag, keine Ablehnung.
        """
        name = normalisiere(bezeichnung)
        gruppe_norm = normalisiere(gruppe)
        seite = seite or self._seite_aus_gruppe(gruppe_norm)

        z = (self._konzern(name, seite, konzernnamen)
             or self._per_stichwort(name, self.QUELLE_STICHWORT, "name")
             or self._per_bibliothek(name)
             or self._per_stichwort(gruppe_norm, self.QUELLE_GRUPPE, "gruppe")
             or self._per_kategorie(gruppe_norm))
        if z is None:
            return None
        if seite and self.klasse_von(z.fs_line) not in (None,):
            z.seite = seite
        z.flags += self._fristigkeit_pruefen(z.fs_line, fristigkeit)
        return z

    def _seite_aus_gruppe(self, gruppe_norm: str) -> Optional[str]:
        if " asset" in gruppe_norm:
            return "AKTIVA"
        if " liabilit" in gruppe_norm or " equity" in gruppe_norm:
            return "PASSIVA"
        return None

    def _konzern(self, name: str, seite: Optional[str],
                 konzernnamen: Iterable[str]) -> Optional[Zuordnung]:
        """Verbundene Unternehmen und Gesellschafter gehen allem anderen vor.

        Ihre Namen stehen in keinem Regelwerk — sie werden zu Mandatsbeginn
        erfasst (Setup-Frage 5) und hier angehängt. Ob Forderung oder
        Verbindlichkeit, sagt die Bilanzseite; ohne Seite bleibt es die
        gemischte Konzernposition, die die Typ-2-Regeln auflösen.
        """
        for roh in konzernnamen:
            treffer = normalisiere(roh).strip()
            if not treffer or not _trifft(treffer, name):
                continue
            fs = ("Related-party receivables" if seite == "AKTIVA"
                  else "Related-party payables" if seite == "PASSIVA"
                  else "Related-party financial assets")
            reg = self.fs_lines.get(fs)
            if reg is None:
                continue
            return Zuordnung(
                fs_line=fs, klasse=reg.klasse, quelle=self.QUELLE_STICHWORT,
                regel_id=f"konzern:{treffer}",
                begruendung=(f"Name des verbundenen Unternehmens '{roh}' im "
                             f"Kontonamen — die Konzernregel geht den übrigen "
                             f"Stichworten vor."),
                seite=seite)
        return None

    def _per_stichwort(self, text: str, quelle: str,
                       gegen: str) -> Optional[Zuordnung]:
        bestes = ""
        treffer: Optional[Stichwortregel] = None
        for r in self.stichwortregeln:
            for s in r.stichworte:
                if len(s) > len(bestes) and _trifft(s, text):
                    bestes, treffer = s, r
        if treffer is None:
            return None
        return Zuordnung(
            fs_line=treffer.fs_line, klasse=treffer.klasse, quelle=quelle,
            regel_id=f"{gegen}:{bestes}",
            begruendung=(f"Stichwort '{bestes}' im "
                         f"{'Kontonamen' if gegen == 'name' else 'Kontogruppe'}"
                         f" -> {treffer.fs_line} (längster Treffer)."),
            seite=self.seite_von(treffer.fs_line))

    def _per_bibliothek(self, name: str) -> Optional[Zuordnung]:
        bestes: Optional[Musterkonto] = None
        for m in self.bibliothek:
            kern = m.name_norm.strip()
            if not kern or not _trifft(kern, name):
                continue
            if bestes is None or len(kern) > len(bestes.name_norm.strip()):
                bestes = m
        if bestes is None:
            return None
        klasse = self.klasse_von(bestes.fs_line) or "PL"
        return Zuordnung(
            fs_line=bestes.fs_line, klasse=klasse,
            quelle=self.QUELLE_BIBLIOTHEK,
            regel_id=f"bibliothek:{bestes.name}",
            begruendung=(f"Musterkonto '{bestes.name}' der Kontenbibliothek "
                         f"im Kontonamen -> {bestes.fs_line}."),
            seite=self.seite_von(bestes.fs_line), current=bestes.current)

    def _per_kategorie(self, gruppe_norm: str) -> Optional[Zuordnung]:
        """Kontogruppe der SuSa gegen Kategorie/Unterkategorie der Bibliothek.

        Letzte Stufe vor der Review-Queue. Sie trifft grob — eine Gruppe wie
        "Trade & other receivables" sagt die Position, nicht das Konto.
        """
        bestes: Optional[tuple[str, Musterkonto]] = None
        for m in self.bibliothek:
            for feld in (m.kategorie_norm, m.unterkategorie_norm):
                kern = feld.strip()
                if not kern or not _trifft(kern, gruppe_norm):
                    continue
                if bestes is None or len(kern) > len(bestes[0]):
                    bestes = (kern, m)
        if bestes is None:
            return None
        kern, m = bestes
        klasse = self.klasse_von(m.fs_line) or "PL"
        return Zuordnung(
            fs_line=m.fs_line, klasse=klasse, quelle=self.QUELLE_GRUPPE,
            regel_id=f"kategorie:{kern.strip()}",
            begruendung=(f"Kein Treffer über den Kontonamen. Kontogruppe "
                         f"trifft die Kategorie '{kern.strip()}' der "
                         f"Bibliothek -> {m.fs_line}."),
            seite=self.seite_von(m.fs_line), current=m.current)

    def _fristigkeit_pruefen(self, fs_line: str, fristigkeit: str) -> list[str]:
        if not fristigkeit:
            return []
        # ``current`` fuehrt Positionen, die beides sein koennen, als
        # "Current / Non-current" in EINEM Feld. Ungetrennt gelesen widerspricht
        # dieser Eintrag jeder Angabe der Quelle — und die Warnung waere immer
        # falsch, gerade bei den Positionen, wo sie zaehlt (Provisions).
        soll = {teil.strip() for m in self.bibliothek
                if m.fs_line == fs_line and m.current
                for teil in m.current.split("/") if teil.strip()}
        if not soll:
            return []
        # Hier wird verglichen, nicht gesucht: "Non-current" enthält
        # "current" und träfe jede Substring-Prüfung.
        ist = normalisiere(fristigkeit)
        if any(normalisiere(s) == ist for s in soll):
            return []
        return [f"Fristigkeit der Quelle ('{fristigkeit}') widerspricht der "
                f"Bibliothek ({', '.join(sorted(soll))}) für '{fs_line}'."]


#: Antwort des Setup-Dialogs -> Konfigurationsdatei. ``None`` heißt: die
#: bestehende HGB-Kette (Typ-1-Regeln + SKR-Bereichstabelle) greift unverändert.
RAHMEN: dict[str, tuple[Optional[str], str]] = {
    "skr03": (None, "SKR03 (DATEV) — HGB-Kette der Hausconvention"),
    "skr04": (None, "SKR04 (DATEV) — HGB-Kette der Hausconvention"),
    "hgb": (None, "eigener ERP-Plan nach HGB — HGB-Kette der Hausconvention"),
    "aasb": ("kontenrahmen_aasb_v1.json",
             "AASB (Australien) — Zuordnung über den Kontonamen"),
}


def lade_rahmen(antwort: str) -> Optional[Kontenrahmen]:
    """Antwort auf die Kontenrahmen-Frage des Setup-Dialogs -> Rahmen.

    Gibt ``None`` zurück, wenn die HGB-Kette gilt. Eine unbekannte Antwort ist
    ein Fehler und kein stiller Rückfall auf HGB — sonst liefe ein
    IFRS-Mandat unbemerkt gegen die SKR-Bereichstabelle.
    """
    schluessel = (antwort or "").strip().lower()
    if schluessel not in RAHMEN:
        raise ValueError(
            f"Unbekannter Kontenrahmen '{antwort}'. Zulässig: "
            + ", ".join(sorted(RAHMEN)) + ".")
    datei, _ = RAHMEN[schluessel]
    return Kontenrahmen.laden(datei, schluessel) if datei else None
