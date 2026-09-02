"""Fuehrt die Entscheidungsdatei ``klassifizierung_v1.json`` aus.

Die Datei ist die Entscheidung, dieses Modul nur der Motor. Es steht hier
bewusst keine einzige Kontobezeichnung und keine einzige Kategorie im Code —
wer eine Regel aendern will, aendert die JSON und laest ``selbsttest.py``
laufen. Genau darum ist die Klassifizierung im Databook abgeleitet und nicht
eingegeben.

Das Verfahren steht in der Datei unter ``_verfahren`` und wird hier Schritt
fuer Schritt umgesetzt:

1. HGB-Gliederung, falls vorhanden: Longest-Prefix-Match auf ``hgb_positionen``
   liefert eine Grobklasse.
2. Ist die Grobklasse ``MIXED`` oder gibt es keine Gliederung, entscheiden die
   Stichwortregeln.
3. Kontoname normalisieren und gegen ``regeln`` pruefen.
4. Nur wenn der Name nichts liefert, die Kontogruppe gegen
   ``regeln_kontogruppe``.
5. Innerhalb einer Quelle gewinnt die hoechste Prioritaet, bei Gleichstand
   geht das Konto in die Review-Queue.
6. Greift nichts, entscheidet der Fallback nach Bilanzseite — immer mit Review.

Zwei Punkte, an denen eine naive Umsetzung falsch wird:

**Normalisierung muss beidseitig sein.** Die Stichworte tragen selbst
Sonderzeichen (``r&d finance``, ``tools & lab``, ``kst.``, ``ausst.rg``).
Wer nur den Kontonamen normalisiert und die Stichworte roh laesst, verliert
diese Regeln lautlos — sie treffen dann nie. Deshalb laeuft dieselbe Funktion
ueber beide Seiten.

**Leerzeichen bleiben erhalten.** Das Stichwort ``moss `` traegt sein
Leerzeichen als Wortgrenze, damit es nicht in ``mossy`` faellt. Ein
``strip()`` oder ein Zusammenziehen mehrfacher Leerzeichen wuerde diese
Absicht wegwerfen. Sonderzeichen werden deshalb durch ein Leerzeichen
**ersetzt** und nicht entfernt: auf beiden Seiten entsteht dieselbe Form.
"""

from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import dataclass, field
from typing import Optional

#: Umlaute und Eszett werden aufgeloest, nicht entfernt — ``rueckstellung``
#: und ``Rückstellung`` muessen dieselbe Form ergeben.
_UMLAUTE = {"ä": "ae", "ö": "oe", "ü": "ue", "ß": "ss",
            "Ä": "ae", "Ö": "oe", "Ü": "ue"}

_NICHT_ALPHANUM = re.compile(r"[^a-z0-9 ]")


def normalisiere(text: object) -> str:
    """Klein, Umlaute aufgeloest, Sonderzeichen zu Leerzeichen.

    Wird auf Kontonamen, Kontogruppen UND auf die Stichworte der Regeln
    angewandt. Nur so trifft ``r&d finance`` das Konto ``R&D finance``.
    """
    if text is None:
        return ""
    s = str(text)
    for zeichen, ersatz in _UMLAUTE.items():
        s = s.replace(zeichen, ersatz)
    # Akzente aufloesen (é -> e), damit auch fremdsprachige Namen greifen.
    s = unicodedata.normalize("NFKD", s)
    s = "".join(z for z in s if not unicodedata.combining(z))
    return _NICHT_ALPHANUM.sub(" ", s.lower())


@dataclass
class Treffer:
    """Eine Regel, die auf ein Konto passt."""

    regel_id: str
    kategorie: str
    prioritaet: int
    regeltyp: str               # "inhaltsregel" | "kategorieregel"
    quelle: str                 # "kontoname" | "kontogruppe" | "hgb" | "fallback"
    review: bool = False
    pflichtfrage: str = ""
    hinweis: str = ""


@dataclass
class Regel:
    """Eine Stichwortregel aus der Entscheidungsdatei, vorbereitet."""

    id: str
    kategorie: str
    prioritaet: int
    regeltyp: str
    quelle: str
    wenn_eines: list[str]
    wenn_keines: list[str] = field(default_factory=list)
    review: bool = False
    pflichtfrage: str = ""
    hinweis: str = ""

    def trifft(self, text: str) -> bool:
        if not any(w in text for w in self.wenn_eines):
            return False
        return not any(w in text for w in self.wenn_keines)


class Classifier:
    """Ordnet einem Sachkonto genau eine Kategorie zu.

    ``spec`` ist der Pfad zur Entscheidungsdatei oder die bereits geladene
    Struktur. ``mandantennamen`` ergaenzt die Setup-Frage aus der Datei: die
    Namen der verbundenen Unternehmen und Gesellschafter, die kein
    generisches Stichwort erkennt. Sie werden der Regel ``rp-group``
    zugeschlagen, statt die Datei zu veraendern.
    """

    def __init__(self, spec, mandantennamen: Optional[list[str]] = None,
                 aktivkonto_erste_ziffer: Optional[str] = None) -> None:
        if isinstance(spec, (str, bytes)):
            with open(spec, encoding="utf-8") as f:
                self.spec = json.load(f)
        else:
            self.spec = spec
        self.version = self.spec.get("_version", "?")
        self.aktivkonto_erste_ziffer = aktivkonto_erste_ziffer

        self._name_regeln = self._lade("regeln", "kontoname")
        self._gruppen_regeln = self._lade("regeln_kontogruppe", "kontogruppe")
        if mandantennamen:
            self._ergaenze_konzernnamen(mandantennamen)

        # Longest-Prefix-Match braucht die laengsten Pfade zuerst.
        self._positionen = sorted(self.spec.get("hgb_positionen", []),
                                  key=lambda p: len(p["hgb_pfad"]), reverse=True)
        self._bereiche = self.spec.get("hgb_kontonummern_default", {}) \
                                  .get("bereiche", [])
        self._fallback = self.spec.get("fallback", {})

    # -- Aufbau ------------------------------------------------------------
    def _lade(self, schluessel: str, standardquelle: str) -> list[Regel]:
        regeln = []
        for r in self.spec.get(schluessel, []):
            regeln.append(Regel(
                id=r["id"], kategorie=r["kategorie"],
                prioritaet=int(r.get("prioritaet", 0)),
                regeltyp=r.get("regeltyp", "inhaltsregel"),
                quelle=r.get("quelle", standardquelle),
                wenn_eines=[normalisiere(w) for w in r.get("wenn_eines", [])],
                wenn_keines=[normalisiere(w) for w in r.get("wenn_keines", [])],
                review=bool(r.get("review", False)),
                pflichtfrage=r.get("pflichtfrage", ""),
                hinweis=r.get("hinweis", "")))
        return regeln

    def _ergaenze_konzernnamen(self, namen: list[str]) -> None:
        """Antwort auf die Setup-Frage: Namen der Konzerngesellschaften.

        Konzernkonten heissen oft nur nach der Gesellschaft. Ohne diese Namen
        faellt ein Konzerndarlehen in den Fallback und steht im Working
        Capital statt im Net Debt.
        """
        for regel in self._name_regeln + self._gruppen_regeln:
            if regel.id in ("rp-group", "grp-relatedparty"):
                regel.wenn_eines += [normalisiere(n) for n in namen if n]

    # -- Klassifizierung ---------------------------------------------------
    def classify(self, account_no: object, account_name: object,
                 account_group: object = "", hgb_pfad: Optional[str] = None,
                 seite: Optional[str] = None) -> dict:
        """Kategorie, Herkunft und Protokollspur fuer ein Konto.

        ``hgb_pfad`` und ``seite`` sind optional: eine MYOB- oder NetSuite-
        Saldenliste liefert keinen HGB-Pfad, und dann entscheiden allein die
        Stichwortregeln. ``seite`` ("AKTIVA"/"PASSIVA") steuert nur den
        Fallback und wird sonst aus der Kontonummer abgeleitet.
        """
        name = normalisiere(account_name)
        gruppe = normalisiere(account_group)

        grobklasse = self._grobklasse(hgb_pfad)
        name_treffer, name_gleichstand = self._beste(self._name_regeln, name)
        gruppen_treffer, gruppen_gleichstand = (None, [])
        if name_treffer is None:
            # SCHRITT 4: die Kontogruppe erst, wenn der Name nichts liefert.
            # Der Name hat Vorrang, auch bei niedrigerer Prioritaet.
            gruppen_treffer, gruppen_gleichstand = self._beste(
                self._gruppen_regeln, gruppe)

        treffer = name_treffer or gruppen_treffer
        gleichstand = name_gleichstand if name_treffer else gruppen_gleichstand

        konflikt = ""
        if grobklasse and grobklasse != "MIXED":
            # Die HGB-Gliederung weiss mehr als ein Kontoname. Eine
            # Stichwortregel darf sie verfeinern, aber nicht umstossen.
            if treffer is None or treffer.kategorie.split("/")[0] != grobklasse:
                if treffer is not None:
                    konflikt = (f"Stichwortregel {treffer.regel_id} wollte "
                                f"{treffer.kategorie}, die HGB-Gliederung sagt "
                                f"{grobklasse}. Die Gliederung gewinnt.")
                treffer = self._aus_hgb(hgb_pfad, grobklasse, konflikt)

        if treffer is None:
            treffer = self._fallback_treffer(account_no, hgb_pfad, seite)

        klasse, _, position = treffer.kategorie.partition("/")
        review = (treffer.review or treffer.regeltyp == "kategorieregel"
                  or treffer.quelle == "fallback" or bool(gleichstand))
        return {
            "account_no": str(account_no or ""),
            "account_name": str(account_name or ""),
            "account_group": str(account_group or ""),
            "category": treffer.kategorie,
            "klasse": klasse,
            "position": position,
            "source": treffer.quelle,
            "rule_id": treffer.regel_id,
            "prioritaet": treffer.prioritaet,
            "regeltyp": treffer.regeltyp,
            "review": review,
            "pflichtfrage": treffer.pflichtfrage,
            "hinweis": (treffer.hinweis + (" " + konflikt if konflikt else "")).strip(),
            "wc_seite": self._wc_seite(treffer.kategorie, account_no, hgb_pfad,
                                       seite),
            "gleichstand": [t.regel_id for t in gleichstand],
        }

    # -- Bausteine ---------------------------------------------------------
    def _beste(self, regeln: list[Regel], text: str
               ) -> tuple[Optional[Treffer], list[Treffer]]:
        """Hoechste Prioritaet gewinnt. Bei Gleichstand: Review-Queue.

        Der Gleichstand wird nicht stillschweigend nach Reihenfolge
        aufgeloest. Zwei Regeln gleicher Prioritaet mit verschiedenen
        Kategorien sind ein Befund fuer die Entscheidungsdatei, kein
        Zufallsergebnis.
        """
        if not text:
            return None, []
        passend = [r for r in regeln if r.trifft(text)]
        if not passend:
            return None, []
        hoechste = max(r.prioritaet for r in passend)
        spitze = [r for r in passend if r.prioritaet == hoechste]
        gewinner = spitze[0]
        treffer = Treffer(gewinner.id, gewinner.kategorie, gewinner.prioritaet,
                          gewinner.regeltyp, gewinner.quelle, gewinner.review,
                          gewinner.pflichtfrage, gewinner.hinweis)
        uneinig = [r for r in spitze if r.kategorie != gewinner.kategorie]
        gleich = [Treffer(r.id, r.kategorie, r.prioritaet, r.regeltyp,
                          r.quelle, r.review, r.pflichtfrage, r.hinweis)
                  for r in uneinig]
        return treffer, gleich

    def _grobklasse(self, hgb_pfad: Optional[str]) -> Optional[str]:
        """Longest-Prefix-Match auf die HGB-Positionstabelle."""
        if not hgb_pfad:
            return None
        for pos in self._positionen:
            if hgb_pfad.startswith(pos["hgb_pfad"]):
                return pos["klasse"]
        return None

    def _aus_hgb(self, hgb_pfad: str, grobklasse: str, hinweis: str) -> Treffer:
        pos = next((p for p in self._positionen
                    if hgb_pfad.startswith(p["hgb_pfad"])), None)
        kategorie = {"FA": "FA/PPE", "DT": "DT/Deferred tax",
                     "EQ": "EQ/Equity", "TWC": "TWC/Trade receivables",
                     "OWC": "OWC/Other receivables",
                     "ND": "ND/Other debt like"}.get(grobklasse,
                                                     "OWC/Other payables")
        return Treffer("hgb-" + (pos["na_de"] if pos else grobklasse),
                       kategorie, 200, "inhaltsregel", "hgb",
                       review=bool(hinweis), hinweis=hinweis)

    def _seite(self, account_no: object, hgb_pfad: Optional[str],
               seite: Optional[str]) -> Optional[str]:
        """Bilanzseite: gegebene Angabe, sonst HGB-Pfad, sonst Kontonummer."""
        if seite:
            return seite.upper()
        if hgb_pfad:
            return "AKTIVA" if hgb_pfad.startswith("/Aktiva") else "PASSIVA"
        nummer = str(account_no or "").strip()
        bereich = self._bereich(nummer)
        if bereich:
            return "AKTIVA" if bereich.startswith("/Aktiva") else "PASSIVA"
        # Kontenplaene der Form ``1-10100``: die erste Ziffer traegt die Seite.
        erste = self.aktivkonto_erste_ziffer or "1"
        if nummer[:1].isdigit():
            return "AKTIVA" if nummer[:1] == erste else "PASSIVA"
        return None

    def _bereich(self, nummer: str) -> Optional[str]:
        """HGB-Pfad aus dem Kontonummernbereich (nur reine Ziffernkonten)."""
        if not nummer.isdigit():
            return None
        n = int(nummer)
        for von, bis, pfad in self._bereiche:
            if von <= n <= bis:
                return pfad
        return None

    def _fallback_treffer(self, account_no, hgb_pfad, seite) -> Treffer:
        """SCHRITT 6: kein Treffer. Der Fallback ist ein Platzhalter.

        Er entscheidet nichts — er haelt die Stelle frei und stellt die
        Pflichtfrage. Deshalb traegt er immer ``review``.
        """
        pfad = hgb_pfad or self._bereich(str(account_no or "").strip()) or ""
        if pfad.startswith("/Aktiva/A Anlagevermoegen"):
            return Treffer("fallback-anlagevermoegen", "FA/PPE", 0,
                           "kategorieregel", "fallback", review=True,
                           pflichtfrage="aufriss",
                           hinweis=self._fallback.get("anlagevermoegen", ""))
        if pfad.startswith("/Passiva/A Eigenkapital"):
            return Treffer("fallback-eigenkapital", "EQ/Equity", 0,
                           "kategorieregel", "fallback", review=True,
                           pflichtfrage="aufriss",
                           hinweis=self._fallback.get("eigenkapital", ""))
        aktiv = self._seite(account_no, hgb_pfad, seite) != "PASSIVA"
        kategorie = self._fallback["aktivseite_umlaufvermoegen"] if aktiv \
            else self._fallback["passivseite_kurzfristig"]
        return Treffer("fallback-" + ("aktiv" if aktiv else "passiv"),
                       kategorie, 0, "kategorieregel", "fallback", review=True,
                       pflichtfrage="aufriss",
                       hinweis=self._fallback.get("regel", ""))

    def _wc_seite(self, kategorie: str, account_no, hgb_pfad,
                  seite) -> str:
        """OA/OL nach ``oa_ol_ableitung``: die Seite, nicht die Klasse.

        Aktive Rechnungsabgrenzung bleibt OWC und liegt auf der OA-Seite,
        passive bleibt OWC auf der OL-Seite. Klassen ausserhalb des Working
        Capital tragen keine Seite.
        """
        if kategorie.split("/")[0] not in ("TWC", "OWC"):
            return ""
        return "OA" if self._seite(account_no, hgb_pfad, seite) != "PASSIVA" \
            else "OL"
