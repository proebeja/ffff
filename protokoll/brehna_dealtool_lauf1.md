# Umstellung auf die Dealtool-Vorlage — erster Lauf (Projekt Brehna)

Die Vorlage ist eingebaut, die Hausconvention steht auf v3.0, und der
Brehna-Lauf schreibt kein selbstgebautes Databook mehr, sondern befüllt
`Dealtool_Template_v4_1.xlsx`.

Dieses Protokoll hält fest, was gebaut wurde, was die Abnahme ergeben hat und
was offen bleibt.

## Was gebaut ist

| Änderungspunkt der Übergabe | Stand |
| --- | --- |
| 1 Vorlage-Mechanik, keine Formatkonstruktion | umgesetzt für Lead NA, Lead PL, Cockpit, Mastersheet |
| 2 Setup-Dialog schreibt ins Cockpit | umgesetzt (`vorlage.Mandat` → `_schreibe_cockpit`) |
| 3 Berichtssprache blendet ein und aus | umgesetzt, 26 Tabs |
| 4 Positionen über Dummy-Zeilen | umgesetzt, drei Dummy-Zeilen belegt |
| 5 Kontoslots befüllen | umgesetzt |
| 6 Architektur-Option B (Aufriss speist die Positionssumme) | **offen** |
| 7 Recalc mit `--force` | umgesetzt (`export/pruefung.py`) |

Option A ist damit vollständig: die Lead-Tabs ziehen ihre Positionssummen per
`SUMIFS` direkt aus dem Mastersheet, die Kontoslots ebenso. Für Option B
fehlen die befüllten Aufriss-Tabs und die Kontrollzeile Aufriss gegen Konten.

## Abnahme

| Kriterium | Ergebnis |
| --- | --- |
| Zellweise formatgleich mit der Vorlage | 0 Abweichungen über 38 Vorlagenblätter |
| Zeilennummern unverändert | Lead NA 237, Lead PL 265 — wie die Vorlage |
| Gliederungssymbol neben der Positionszeile | `summaryBelow = False` gesetzt |
| Kontoslots eingeklappt und grau | unverändert aus der Vorlage übernommen |
| Kein blankes `SUM` über Kontoslots | keins |
| Check-Zeile Lead NA geht auf null | ja, in allen neun Spalten |
| Check-Zeile Lead PL geht auf null | ja, in allen zehn Spalten |
| Fehlerzahl nach Recalc | Vorlage 142, Ausgabe 139 — **0 neue Fehler** |

Die Fehlerzahl ist gegen die **leere Vorlage** gemessen, nicht gegen die 117
aus der Übergabe. Der Grund: die Zahl 117 stammt aus Excel, hier rechnet der
`formulas`-Kern, und der ordnet dieselben Add-in-Aufrufe anderen Fehlerarten
zu (128 `#VALUE!` in `NA_monthly` statt 112 `#NAME?` über `NA_monthly` und
`PL_monthly`). Vergleichbar ist deshalb nur die Differenz, und die ist null.
Drei Fehler sind sogar verschwunden: die `#DIV/0!` in `PL_Top customer` haben
jetzt einen Nenner.

Das `--force` der Kommandozeile gibt es in dieser `formulas`-Fassung nicht
mehr; `export/pruefung.py` baut es nach, indem es Zellen, deren Text mit `=`
beginnt ohne eine Formel zu sein, als Text behandelt. Betroffen ist genau eine
Zelle: `Cockpit!G5` mit dem Text `=technische hilfszelle --> nicht löschen`.

## Kontrollzeilen

Alle im Datenmodell gerechnet und danach gegen die Arbeitsmappe gehalten.

| Kontrolle | FY2023 | FY2024 | FY2025 | YTD 07/2026 |
| --- | ---: | ---: | ---: | ---: |
| Bilanzidentität (Summe aller Konten) | 0,00 | 0,00 | 0,00 | 0,00 |
| Nettovermögen + Eigenkapital + Periodenergebnis | 0,00 | 0,00 | 0,00 | 0,00 |
| Equity Roll Forward gegen Nettovermögen | 0,00 | 0,00 | 0,00 | 0,00 |
| Ergebnis lt. Quelle gegen Summe der GuV-Positionen | 0,00 | 0,00 | 0,00 | 0,00 |
| Betrag in Positionen ohne Kontoslots | 0,00 | 1.800,00 | 1.486.008,89 | 3.717.033,73 |

Die letzte Zeile ist kein Rechenfehler, sondern der Befund aus Punkt 1 unten:
diese Beträge stehen in Positionen, für die die Vorlage keine Kontoslots
vorhält.

Das Periodenergebnis gehört in die zweite Zeile hinein. Der Abschluss weist
den Bilanzverlust als Position **ohne eigenes Konto** aus; das Ergebnis steht
also noch auf den GuV-Konten und nicht im Eigenkapital. Ohne den Summanden
ginge die Zeile Jahr für Jahr genau um das Periodenergebnis daneben.

## Laufprotokoll

| Phase | Sekunden | Anteil |
| --- | ---: | ---: |
| Setup | 0,00 | 0 % |
| Einlesen Jahresabschlüsse (PDF) | 5,18 | 71 % |
| Mapping (Kaskade) | 0,00 | 0 % |
| v2.8-Nachlauf | 0,00 | 0 % |
| Views | 0,00 | 0 % |
| Aufrisse | 0,00 | 0 % |
| QA-Diagnose | 0,00 | 0 % |
| Vorlage befüllen | 2,11 | 29 % |
| **gesamt** | **7,30** | |

1.038 Zellen geschrieben, 32 Mastersheet-Zeilen, 44 Blätter in der Ausgabe.
Null Aufrufe an die KI-Schicht — es ist weiterhin kein Provider registriert.

Der Recalc dauert mit rund 140 Sekunden das Zwanzigfache des Laufs selbst. Er
gehört deshalb in die Abnahme, nicht in die Testsuite.

## Befunde

Die ersten beiden sind zentral an der Vorlage zu entscheiden, nicht je
Mandat.

**1. Der Net-Debt-Block hat zu wenige Positionen.** Brehna braucht dort drei
Zeilen, die die Vorlage nicht führt: den thereof-ND-Anteil der sonstigen
Verbindlichkeiten und der sonstigen Vermögensgegenstände sowie die
Verbindlichkeiten gegenüber verbundenen Unternehmen. Sie sind über die drei
Dummy-Zeilen angelegt — damit sind alle drei verbraucht, und Dummy-Zeilen
haben keine Kontoslots. 3.717 T€ stehen in der YTD-Spalte ohne Kontodetail.

**2. `NA_Cons` Zeile 45.** Der zerstörte Verweis lässt sich fachlich
eingrenzen, aber nicht eindeutig setzen. Der Spiegel des Net-Debt-Blocks
(Zeilen 39 bis 43, summiert in 44) führt Lead NA 180, 189, 207, 208, 209 —
**Zeile 198, Pensionsrückstellungen, fehlt.** Der Spiegel des Blocks für
latente Steuern (45 bis 49, summiert in 50) hat dagegen eine Zeile zu viel:
46 bis 49 decken Lead NA 211, 220, 221, 222 bereits ab. Zeile 45 ist also die
überzählige Zeile im falschen Block. Ein Verweis auf 198 an dieser Stelle
würde die Pensionsrückstellungen in die Summe der latenten Steuern ziehen,
also wäre zusätzlich `E44` um `E45` zu erweitern und `E50` auf `SUM(E46:E49)`
zu verkürzen. Das sind zwei Formeländerungen ohne Zeileneinschub — aber eine
Entscheidung an der Vorlage, deshalb hier nur der Vorschlag.

**3. Die Sprachblöcke stehen nicht überall in denselben Zeilen.** In den
Lead-Tabs sind sie zwei Zeilen lang (deutsch 4/5, englisch 6/7), in `NA_Cons`,
`PL_Cons`, `Delta_JA_SuSa` und `Plan_IST_Vergleich` drei (4 bis 6 und 7 bis 9)
und in `PL_YTD` drei Zeilen tiefer angesetzt (5 bis 7 und 8 bis 10). Eine
Erkennung auf den Zeilen 4 bis 7 hätte fünf Tabs stillschweigend übersprungen
— das war der erste Stand und ist korrigiert. Die Blockgrenzen werden jetzt
aus dem Blatt gelesen: der deutsche Block beginnt, wo Spalte C den
Projektnamen zieht, der englische, wo Spalte D dasselbe tut, und ihr Abstand
ist die Blocklänge. Damit schalten alle 26 Tabs um, die beide Fassungen
führen. Das ist kein Befund an der Vorlage, sondern eine Notiz für den
nächsten, der die Erkennung anfasst.

## Zwei Vokabulare

Die Vorlage benennt ihre Positionen anders als unsere Hausconvention. Auf der
Bilanzseite stimmen beide Zeichen für Zeichen überein; die GuV der Vorlage ist
dagegen keine § 275-Gliederung, sondern eine Kostenartenrechnung. Sieben
Positionen heißen anders, zwei fehlen ganz. Die Zuordnung steht vollständig in
`export/vorlage_zuordnung.py` und je Lauf im Tab `Zuordnung`.

Der Fall, vor dem die Übergabe warnt — ein Ticker heißt `verbundene` statt
`verbundenen` und die Position zeigt stumm null —, ist damit nicht mehr eine
Frage der Sorgfalt: `tests/test_vorlage.py` prüft in beide Richtungen, dass
jedes Ticker-Paar im Mastersheet vorkommt und jedes Mastersheet-Paar eine
Positionszeile hat.

## Vorzeichen der GuV

Unsere Konvention ist Soll positiv, damit eine Summen- und Saldenliste zu null
aufgeht. Die Vorlage rechnet `Rohergebnis = Gesamtleistung + Materialaufwand`
und braucht deshalb Erträge positiv und Aufwendungen negativ. Beim Schreiben
ins Mastersheet wird die GuV gedreht, die Bilanz nicht. Ohne diese Drehung
stimmt jede Zwischensumme der GuV im Vorzeichen nicht.

## Was noch nicht umgestellt ist

`kitchenstories.py` und `cli.py` schreiben weiterhin über `export/excel.py`
und damit im alten Hausformat. Sie brauchen Tabs, für die die Vorlage keinen
Platz vorsieht (Benchmark, Reconciliation, Delta Lauf 1/Lauf 2) — das ist eine
fachliche Entscheidung, keine Umbauarbeit, und deshalb hier nicht vorweggenommen.
