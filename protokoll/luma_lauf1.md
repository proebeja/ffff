# Projekt Luma — Databook nach Option A

Eine Quelle, vier Jahrestabs, 198 Konten. Kein Jahresabschluss, kein
Kontennachweis, kein Prüfbericht. Das Databook ist eine Kopie der
Dealtool-Vorlage, befüllt nach **Option A**: die Lead-Tabs ziehen ihre
Positionssummen per `SUMIFS` direkt aus dem Mastersheet, die Kontoslots ebenso.

## Was die Quelle ist — und was sie nicht ist

Der Export stammt aus MYOB, nicht aus DATEV. Vier Eigenheiten entscheiden über
das Databook, und jede davon führt bei falscher Lesart zu einem Buch, das
plausibel aussieht und falsch ist.

**Die Wertespalte ist nicht die erste Zahlenspalte.** Spalte L führt die
Bewegung der Periode, Spalte O den Schlussbestand — und die Kopfzeile nennt
Spalte O irreführend `Year`. Beide summieren je Periode auf null, eine
Saldenliste tut das und eine Bewegungsspalte auch. Wer L nimmt, baut ein
Databook aus Veränderungen: die Aktivseite läge in FY2023 bei 260 TAUD statt
bei 10,4 Mio.

**Das Geschäftsjahr endet am 31. März**, und FY2023 ist ein Rumpfjahr über
neun Monate (01.07.2022 bis 31.03.2023). Aus dem Tabellennamen `FY2023` würde
der 31.12.2023 abgeleitet; die Zeitachse kommt deshalb aus `PeriodFrom` und
`PeriodTo`. Auch die sind nicht wörtlich zu nehmen: `PeriodTo = 2023-03-01`
meint den ersten Tag des letzten Periodenmonats, der Stichtag ist der
31.03.2023.

**Der Kontenplan ist kein SKR.** `1-12300 Trade Debtors - USD` trifft weder
die SKR03-Bereichstabelle noch die deutschen Stichworte der Typ-1-Regeln. Der
Export bringt aber seine eigene Gliederung mit, und die ist im Reader auf
HGB-Pfade übersetzt — 65 Kontogruppen plus 8 Konten, deren Gruppe zu grob
ist, jede Zuordnung mit Begründung. Damit greift
Stufe 1 der Kaskade, und Klasse und NA-Zeile leitet die Reklassifizierung wie
gewohnt ab. Eine Gruppe blieb offen: `9-99999 Suspense Account`. Sie geht in
die Review-Queue statt in eine Bilanzposition, die jemand später für belegt
hält; der Saldo ist in allen vier Perioden null.

**Es gibt keine GuV.** Der Export führt nur die Klassen 10 bis 50. Das
Periodenergebnis steht als ein Betrag auf `3-90000 Current Earnings`.

## Status je Spalte

Alle vier Spalten: **abgeleitet, nicht abschlusstreu**. Es liegt kein
Abschluss vor, gegen den sich das Databook überleiten ließe. Die Gliederung
stammt aus der Systemgliederung des Exports — das trägt die Bilanz, ersetzt
aber keinen Kontennachweis. Für die GuV gibt es keine Daten.

| Periode | Stichtag | Monate | Bilanz | GuV |
| --- | --- | ---: | --- | --- |
| FY2023 | 31.03.2023 | 9 | abgeleitet | keine Daten |
| FY2024 | 31.03.2024 | 12 | abgeleitet | keine Daten |
| FY2025 | 31.03.2025 | 12 | abgeleitet | keine Daten |
| FY2026 | 31.03.2026 | 12 | abgeleitet | keine Daten |

## Kontrollzeilen

| Kontrolle | FY2023 | FY2024 | FY2025 | FY2026 |
| --- | ---: | ---: | ---: | ---: |
| Bilanzidentität (Summe aller Konten) | −0,01 | −0,01 | −0,04 | −0,04 |
| Nettovermögen + Eigenkapital | −0,01 | −0,01 | −0,04 | −0,04 |
| Equity Roll Forward gegen Nettovermögen | 0,01 | 0,00 | 0,03 | −0,00 |
| Nicht erklärte Eigenkapitalbewegung | 0,00 | **268.375,20** | −0,00 | −0,00 |
| Konten ohne Position im Lead | 0,00 | −0,01 | 0,00 | 0,00 |
| Konten ohne Kontoslot (Wert ohne Detail) | 1.430.614 | 1.995.711 | 1.965.742 | 1.866.123 |

Die Rundungsdifferenzen von bis zu 4 Cent stammen aus dem Export selbst; die
Spaltensumme des Rohblatts trägt sie ebenso.

## Ein Fehler, den dieser Lauf aufgedeckt hat

Die drei gemischten Positionen — sonstige Vermögensgegenstände, sonstige
Verbindlichkeiten, sonstige Rückstellungen — lösen die Typ-2-Regeln über
deutsche Stichworte auf, und die greifen auf `Accrued Expenses` oder
`Sundry Debtors` nicht. Die Konten bekommen deshalb über QA A6 eine
vorläufige Klasse. Dabei wurde bislang die **WC-Seite nicht mitgesetzt**:
`m.seite` blieb leer, der Ticker der Vorlage lautete `OWC` statt `OA`/`OL`,
und die Vorlage kennt kein `OWC`.

Folge: **34 Konten standen im Mastersheet und in keiner Position des Lead NA**,
während die Bilanzidentität im Datenmodell tadellos aufging. Im alten
Hausformat wäre das nie aufgefallen — dort gruppiert der Lead nach Klasse, und
die Klasse war richtig. Erst der Ticker der Vorlage fragt nach der Seite.

Der Fehler ist in `engine/v28.py` behoben (`_vorlaeufige_seite`, abgeleitet aus
der Bilanzseite des Pfads, wie `oa_ol_ableitung` es vorschreibt). Zusätzlich
gibt es jetzt die Kontrollzeile **„Konten ohne Position im Lead“**: eine
Position, die in keinem Lead-Tab landet, war vorher nur eine Zeile im
Zuordnungsblatt und ist jetzt eine Kontrolle. Übrig bleibt dort einzig das
Suspense-Konto mit einem Cent.

## Zwei Befunde

**1. FY2024: 268.375,20 AUD Eigenkapitalbewegung ohne Beleg.** Sie setzt sich
so zusammen: das FY2023-Ergebnis von 287.823,17 verlässt `Current Earnings`,
`Retained Earnings` nimmt davon aber nur 15.144,98 auf, und das
`Historical Balancing Account` bewegt sich um 4.302,99. Der Rest von
272.678,19 ist als Vorjahresberichtigung direkt gegen den Gewinnvortrag
gebucht worden, statt als Ergebnis vorgetragen zu werden. Das gehört in die
manuelle Zeile des Equity Roll Forward, sobald der Mandant es erklärt — bis
dahin steht es in der Kontrollzeile und nirgends sonst. Verteilt wird es
nicht: eine Zeile namens „Gewinnausschüttung“ mit einer Vorjahresberichtigung
zu füllen wäre eine falsche Aussage.

Sichtbar wird das in der **Check-Zeile 237 des Lead NA**, die ab FY2024
−268,38 T AUD zeigt und den Betrag über die Folgejahre mitträgt. Anders als im
Brehna-Lauf geht sie hier also nicht auf null — und das ist kein Fehler der
Befüllung, sondern die Kontrolle, die ihre Arbeit tut.

**2. Acht Kontoslots reichen für diesen Kontenplan nicht.** Die Vorlage hält
je Position acht eingeklappte Slots vor; das passt zu einem deutschen
SKR-Kontenplan, nicht zu 198 MYOB-Konten:

| Position | Konten | Slots | ohne Detail | Wert ohne Detail FY2026 |
| --- | ---: | ---: | ---: | ---: |
| Sachanlagen | 51 | 8 | 43 | 10.699 |
| Immaterielle Vermögensgegenstände | 18 | 8 | 10 | 1.387 |
| Sonstige Verbindlichkeiten | 15 | 8 | 7 | 0 |
| Vorräte | 14 | 8 | 6 | 0 |
| Liquide Mittel | 14 | 8 | 6 | 0 |
| Sonstige Rückstellungen | 12 | 8 | 4 | 0 |
| Verbindlichkeiten ggü. Kreditinstituten | 12 | 8 | 4 | 0 |
| Verbindlichkeiten aus L+L | 10 | 8 | 2 | 0 |
| Forderungen ggü. verbundenen Unternehmen (ND) | 3 | 0 | 3 | 1.847.486 |
| Steuerrückstellungen (ND) | 2 | 0 | 2 | −273.076 |
| Aktive latente Steuern (DT) | 1 | 0 | 1 | 279.627 |
| Anleihen (ND) | 1 | 0 | 1 | 0 |

Die Positionszeilen selbst tragen den vollen Betrag — es fehlt das sichtbare
Detail, nicht der Wert. Wo die Slots knapp sind, bekommen sie die **größten**
Konten der Position; die Reihenfolge des Mastersheets wäre die Kontonummer,
und dann verschwände das größte Konto womöglich hinter acht kleinen. Genau
deshalb bleibt bei den Positionen der Vorlage kaum etwas unsichtbar: 43 nicht
dargestellte Sachanlagenkonten tragen zusammen 10.699 AUD.

Die vier letzten Zeilen sind etwas anderes. Es sind über Dummy-Zeilen
angelegte Positionen, die die Vorlage gar nicht führt, und Dummy-Zeilen haben
konstruktionsbedingt keine Slots. Dort steht das Geld: 1,85 Mio auf den
Konzernforderungen. Beides — der Zuschnitt des Net-Debt-Blocks und die Zahl
der Slots — ist zentral an der Vorlage zu entscheiden, nicht je Mandat.

## Entscheidungen, die im Reader stehen

Drei Zuordnungen sind fachliche Urteile und keine Ablesungen:

* **Finanzierung ist Net Debt, auch wenn der Export sie unter den
  kurzfristigen Verbindlichkeiten führt.** Betroffen sind `Trade Facility`,
  `R&D finance`, `Finance Leases` (kurz- und langfristig), die
  Firmenkreditkarte, das Darlehen von Partners for Growth und die Wandelanleihe.
  Der Inhalt entscheidet, nicht die Rubrik.
* **`1-12500 Unearned Revenue - Revenue Recognition`** ist ein Aktivkonto mit
  Habensaldo, das abgegrenzte Umsatzerlöse trägt. Der Inhalt ist eine
  erhaltene Anzahlung, deshalb steht es auf der Passivseite. Dadurch liegt die
  Aktivsumme des Databooks um 176.548 unter der des Exports.
* **Kostenstellen werden je Konto summiert.** Drei Konten führt der Export je
  Kostenstelle mehrfach (bis zu neun Zeilen). Das Mastersheet führt ein Konto
  genau einmal; die Kostenstelle wäre ein zweiter Schlüssel neben der
  Kontonummer und liefe der Single-Source-Regel zuwider.

Was nicht eindeutig war, ist bewusst auf eine der drei gemischten Positionen
gelegt worden und steht in der **Review-Queue: 34 Konten**, überwiegend
`Accrued Expenses`, `Sundry Debtors` und Verrechnungskonten. Die Typ-2-Regeln
arbeiten mit deutschen Stichworten und greifen auf englischen Bezeichnungen
nicht; die Konten haben über QA A6 einen vorläufigen Pfad und sind als
bestätigungspflichtig markiert.

## Abnahme gegen die Vorlage

| Kriterium | Ergebnis |
| --- | --- |
| Zellweise formatgleich mit der Vorlage | 0 Abweichungen über 38 Blätter |
| Neu angelegte Zellen | 8 (Ticker der vier Dummy-Zeilen) |
| Zeilennummern | Lead NA 237, Lead PL 265 — unverändert |
| Kein blankes `SUM` über Kontoslots | keins |
| Fehlerzahl nach Recalc | 142 — genau der Stand der leeren Vorlage, **0 neue** |
| Check-Zeile Lead NA | ab FY2024 −268,38 T AUD — siehe Befund 1 |

Die drei `#DIV/0!` in `PL_Top customer` bleiben hier stehen, anders als im
Brehna-Lauf: sie brauchen einen Umsatz als Nenner, und den gibt es in dieser
Quelle nicht.

## Laufprotokoll

| Phase | Sekunden |
| --- | ---: |
| Setup | 0,00 |
| Einlesen Saldenliste (MYOB) | 0,10 |
| Mapping (Kaskade) | 0,01 |
| v2.8-Nachlauf | 0,00 |
| Views | 0,01 |
| QA-Diagnose | 0,00 |
| Vorlage befüllen | 1,86 |
| **gesamt** | **1,97** |

198 Mastersheet-Zeilen, 44 Blätter, null Aufrufe an die KI-Schicht.
