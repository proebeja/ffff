"""Reader für PDF-Kontennachweis (SKR03), z.B. Huchtemeier.

Bewusst best-effort und absturzsicher (Handover: PDF "kann folgen"). Der reale
Datensatz trägt ein diagonales "Testversion"-Wasserzeichen, das einzelne
Zahlen zerhackt — solche Zeilen werden übersprungen und als Warnung vermerkt,
nie als Absturz.

Extrahiert Kontozeilen (SKR03-Nummer + Bezeichnung + 1–2 Beträge). Der HGB-Pfad
wird in dieser Scheibe über die Kaskade (Typ-1 -> SKR-Default) bestimmt, nicht
aus den PDF-Überschriften — die fettgedruckten §266-Überschriften als
FS-Struktur zu nutzen ist ein späterer Ausbau.
"""

from __future__ import annotations

import re
from typing import Optional

from ..core.model import Account, NormalizedLedger, PeriodBalance
from .base import Reader, fingerprint, parse_deutsche_zahl

# Kontozeile: 3–4-stellige Kontonummer, Bezeichnung, dann Beträge (dt. Format,
# optional nachgestelltes Minus). Am Zeilenende ein oder zwei Beträge.
_BETRAG = r"-?\d{1,3}(?:\.\d{3})*,\d{2}-?"
_LINE_RE = re.compile(
    rf"^(?P<konto>\d{{3,4}})\s+(?P<bez>.+?)\s+(?P<b1>{_BETRAG})(?:\s+(?P<b2>{_BETRAG}))?\s*$"
)


class PdfKontennachweisReader(Reader):
    name = "pdf_kontennachweis"

    @classmethod
    def kann_lesen(cls, pfad: str) -> bool:
        return pfad.lower().endswith(".pdf")

    def lesen(self, pfad: str) -> NormalizedLedger:
        warnungen: list[str] = []
        p_cur, p_prev = self._perioden_aus_name(pfad)
        entity = "Unbekannt"
        konten: dict[str, tuple[str, float, float]] = {}

        try:
            import pdfplumber
        except Exception as e:  # pragma: no cover - Abhängigkeit fehlt
            return NormalizedLedger(
                accounts=[], perioden=[p_cur, p_prev], entity=entity,
                quelle_datei=pfad, hat_kontennachweis=False,
                fingerprint=fingerprint(pfad),
                warnungen=[f"pdfplumber nicht verfügbar: {e}"],
            )

        try:
            with pdfplumber.open(pfad) as pdf:
                for page in pdf.pages:
                    text = page.extract_text() or ""
                    for raw in text.splitlines():
                        line = raw.strip()
                        if line.startswith("Huchtemeier"):
                            entity = "Huchtemeier Papier GmbH"
                        m = _LINE_RE.match(line)
                        if not m:
                            continue
                        konto = m.group("konto")
                        bez = re.sub(r"\s+", " ", m.group("bez")).strip()
                        # Wasserzeichen-Schutz: Bezeichnung mit vielen Einzelbuchstaben
                        # oder Ziffernbrei überspringen.
                        if self._sieht_korrupt_aus(bez):
                            warnungen.append(f"Zeile zu Konto {konto} übersprungen (Wasserzeichen?).")
                            continue
                        b1 = parse_deutsche_zahl(m.group("b1"))
                        b2 = parse_deutsche_zahl(m.group("b2")) if m.group("b2") else 0.0
                        # Kontonummern kommen je Seite höchstens einmal vor; erste gewinnt.
                        konten.setdefault(konto, (bez, b1, b2))
        except Exception as e:
            warnungen.append(f"PDF-Parsing abgebrochen: {e}")

        accounts = [
            Account(
                konto=k, bezeichnung=bez,
                salden=(PeriodBalance(p_cur, b1), PeriodBalance(p_prev, b2)),
                entity=entity, fs_pfad=None, kontotyp=None,
            )
            for k, (bez, b1, b2) in sorted(konten.items())
        ]
        return NormalizedLedger(
            accounts=accounts, perioden=[p_cur, p_prev], entity=entity,
            quelle_datei=pfad, hat_kontennachweis=False,
            fingerprint=fingerprint(pfad), warnungen=warnungen,
        )

    @staticmethod
    def _sieht_korrupt_aus(bez: str) -> bool:
        if len(bez) < 2:
            return True
        # Zerhacktes Wasserzeichen erzeugt viele isolierte Einzelzeichen. Bewusst
        # konservativ (hohe Schwelle), damit legitime Abkürzungsnamen wie
        # 'u. a. Kosten' nicht fälschlich verworfen werden — im Zweifel behalten.
        tokens = bez.split()
        einzeln = sum(1 for t in tokens if len(t) == 1)
        return len(tokens) >= 6 and einzeln / len(tokens) > 0.6

    @staticmethod
    def _perioden_aus_name(pfad: str) -> tuple[str, str]:
        m = re.search(r"(20\d{2})", pfad)
        if m:
            jahr = int(m.group(1))
            return f"31.12.{jahr}", f"31.12.{jahr - 1}"
        return "Berichtsjahr", "Vorjahr"
