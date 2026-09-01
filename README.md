# FDD-Databook-Software — erste Scheibe

KI-gestützte Financial-Due-Diligence-Software. Diese erste, bewusst schmale
Scheibe führt **Reader → Engine (Mapping-Kaskade) → Net-Debt-View → Excel-Export**
end-to-end durch und ist gegen die vier realen Testdatensätze validiert.

Gebaut gegen: `FDD_Tool_Architektur_Spezifikation.md` (v1.3),
`FDD_Tool_Fragen_Modul_Spezifikation.md` (v1.0), `hausconvention.json` (v2.2).

## Schnellstart

```bash
pip install openpyxl pandas pdfplumber xlsxwriter pytest
python -m fdd.cli testdata/Testdaten_Eckart_SuSa_2022-2025-03.xlsx -o out/Eckart.xlsx
pytest -q
```

Die CLI erkennt das Format selbst, mappt jedes Konto, baut den Net-Debt-Tab und
schreibt ein Databook im Hausformat (Mastersheet, Net Debt, Review-Queue, Info).

## Architektur (modular, erweiterbar)

```
fdd/
  config/hausconvention.json     Konfiguration (gelesen, nicht hartkodiert)
  core/       model.py           unveränderliche Datensätze (Account, MappedAccount …)
              hausconvention.py  Loader + Normalisierung + Pfad-Wörterbuch
  readers/    base.py            Reader-Interface + dt. Zahl-Parser + Fingerprint
              datev_susa.py      rohe DATEV-SuSa (Eckart) — Spaltenversatz-robust
              sap_bw.py          SAP-BW HGB-Export (BK4756) — FS-Hierarchie-Crosswalk
              pdf_kontennachweis.py  PDF-Kontennachweis (Huchtemeier) — best-effort
              namur_databook.py  Namur-Databook (Regressions-Referenz)
              detect.py          Formaterkennung (Engine bleibt formatagnostisch)
  engine/     cascade.py         Schicht-1-Kaskade (v1.3-Reihenfolge)
              matcher.py         Typ-1/Typ-2 Keyword-Matching
              reclassify.py      Longest-Prefix Pfad → Klasse/NA-Zeile
              ai_hook.py         KI-Urteilsschicht (Interface, in v1 inaktiv)
              decision_log.py    Entscheidungsprotokoll + begründungspflichtiger Override
  views/      schedules.py       Aufriss-Schicht: je NA-Zeile ein Aufriss; die drei
                                 gemischten (NA_OA/OL/OP) mit operating + thereof-ND
              net_debt.py        Net-Debt-View (Klasse=ND, nach NA-Zeile gruppiert)
              working_capital.py Working-Capital-View (OA/OL × TWC/OWC, Ist je Periode)
              review_queue.py    ungelöste/geflaggte Konten (mit v2.3-Marker-Status)
  export/     vorlage.py         befüllt die Dealtool-Vorlage (v3.0, siehe unten)
              vorlage_layout.py  liest den Aufbau eines Lead-Tabs aus der Vorlage
              vorlage_zuordnung.py  unser Positionsvokabular -> das der Vorlage
              pruefung.py        Formatvergleich gegen die Vorlage + Recalc
              excel.py           altes Hausformat — nur noch für die Mandate,
                                 die noch nicht auf die Vorlage umgestellt sind
  vorlagen/   Dealtool_Template_v4_1.xlsx   die Hausvorlage (39 Tabs)
  cli.py                         end-to-end-Verdrahtung
```

### Ausgabe: die Vorlage wird befüllt, nicht gebaut (v3.0)

Ein Databook entsteht als `shutil.copy` der Hausvorlage; der Code schreibt nur
Werte in vorhandene Zellen. Es gibt keinen Grund mehr, Font, Fill, Rahmen oder
Spaltenbreite zu erzeugen — Theme-Farben und Zeilenhöhen lassen sich aus einer
Beschreibung ohnehin nicht rekonstruieren.

Drei Regeln bestimmen den Aufbau, und alle drei sind in `tests/test_vorlage.py`
festgenagelt:

- **Keine Zeileneinschübe in Lead NA oder Lead PL.** Über 2.000 Verweise aus
  den übrigen Tabs zeigen auf feste Zeilennummern, und openpyxl zieht sie beim
  Einfügen nicht mit. Eine neue Position entsteht durch Umbenennen einer
  Dummy-Zeile; reichen die acht Kontoslots nicht, ist das ein Befund für die
  Review-Queue.
- **Ticker werden nie getippt.** Ticker 1 (NA-Zeile bzw. Kontonummer) und
  Ticker 2 (Klasse) stammen aus dem Mastersheet. Eine abweichende Schreibweise
  liefert stumm null, während die Kontozeile darunter einen Wert zeigt.
- **Summen immer als `SUMIF(<Zeilentyp>;"<>KTO";…)`.** Die Kontoslots liegen
  innerhalb der Summenbereiche; ein blankes `SUM` zählt sie doppelt.

Perioden und Beschriftungen stehen ausschließlich im Cockpit. Bilanz und GuV
haben unterschiedlich viele Periodenspalten: in der Bilanz folgt auf die
Jahre unmittelbar der CYT-Stichtag, in der GuV liegt davor noch die
LTM-Spalte. `vorlage.Zeitachse` rechnet diese Zuordnung aus, statt sie
abzuzählen — genau dort saß der Off-by-one im Equity Roll Forward.

### Tragende Prinzipien (aus der Architektur-Spec)

- **Ein HGB-Pfad je Konto** als String; die analytische Klasse (ND/TWC/OWC/FA/EQ/DT)
  wird per Longest-Prefix-Match daraus **abgeleitet**, nie separat eingegeben.
- **Mapping-Kaskade (v1.3):** Kontennachweis/FS-Struktur (maßgeblich) → Hausconvention
  (Typ-1) → Lernbibliothek → SKR-Default → Review-Queue.
- **Drei gemischte Positionen** (sonstige VG / Verbindlichkeiten / Rückstellungen)
  bestimmen ND/OWC inhaltsabhängig über Typ-2-Keyword-Regeln; sonst Review.
- **Single Source of Truth über die Aufriss-Schicht:** Mastersheet → Aufrisse →
  Leads. Jede Kontozahl wohnt im Mastersheet; jeder Aufriss (eine NA-Zeile)
  referenziert die Einzelkonten per sichtbarer Formel und summiert; jede
  Lead-Zeile zieht aus **genau einem** Aufriss — kein Lead greift mehr direkt
  aufs Mastersheet. Die drei gemischten Aufrisse (NA_OA/OL/OP) führen zwei
  Spalten, `operating` (→ WC-Lead) und `thereof ND` (→ ND-Lead), gesteuert allein
  durch die Klasse-Zelle des Mastersheets — die Aufteilung ist an genau einer
  Stelle definiert und kann nicht auseinanderlaufen. Die Kontrollzeile jedes
  Leads prüft die Aufriss-Kette unabhängig gegen die Mastersheet-Summe (→ 0).
- **WC-Definition periodenkonsistent:** der Working-Capital-Tab zeigt das Ist-WC
  je Periode (OA/OL × TWC/OWC); jede Periode läuft durch dieselbe Klassifizierung.
  Die normalisierte Referenz / das Target WC kommt später (braucht die
  `verhaltenspruefung`, in dieser Scheibe noch nicht implementiert).
- **Deterministisch**, mit **gebundenem KI-Hook** an der Review-Stelle (v1 inaktiv).

## Die vier Testdatensätze und ihre Fallstricke

| Datensatz | Reader | Fallstrick (behandelt) |
|---|---|---|
| Eckart (DATEV-SuSa, 4 Perioden) | `datev_susa` | Spaltenversatz: Header „Saldo" in T, Wert in S, Vorzeichen in W |
| SAP BK4756 (HGB-Export) | `sap_bw` | eingebettete FS-Hierarchie; **leere Strings** statt None |
| Huchtemeier (PDF, 3 Jahre) | `pdf_kontennachweis` | „Testversion"-Wasserzeichen zerhackt einzelne Zahlen → Zeile übersprungen + Warnung |
| Namur (fertiges Databook) | `namur_databook` | dient als Regressions-Referenz gegen die vorhandene Kategorisierung |

## Gefundener und behobener Config-Bug (`hausconvention.json`)

Das Keyword `"kst"` (Abkürzung für Körperschaftsteuer) in der Regel
`hgb-steuer-rst` ist selbst ein Substring von „rü**ckst**ellung". Da die Regel
Priorität 100 hat, **feuerte sie auf jeder Rückstellung** — Pensions-, Urlaubs-
und Gewährleistungsrückstellungen wurden fälschlich zu Steuerrückstellungen (ND).
Das korrumpiert den kaufpreisrelevantesten Teil des Net-Debt-Tabs.

Fix (surgical, dokumentiert im `hinweis` der Regel): `"kst"` →
`["kst.", "kst-", " kst", "kst "]` (grenzsichere Varianten). „KSt" bleibt über
`"koerpersch"`/`"koerperschaftsteuer"` und die Varianten abgedeckt. Substring-
Matching bleibt sonst unverändert — nötig für deutsche Komposita wie
„Gewerbesteuerrückstellung". Regressionstest: `tests/test_reclassify_matcher.py`.

## Offene Nahtstellen (spätere Module)

Working-Capital-/GuV-Leads, Aufrisse, Reconciliation (Schicht 3), Konsolidierung
(Schicht 4), Cashflow (Schicht 7), die **live rufende** KI-Schicht und der
Berichts-Schreiber. Deren Anschlusspunkte bleiben offen: der KI-Hook existiert als
Interface, das Datenmodell trägt Entity/Perioden/Fingerprint, und die HGB-Pfade
sind so geschnitten, dass jeder spätere Aufriss seine Konten eindeutig zieht.
