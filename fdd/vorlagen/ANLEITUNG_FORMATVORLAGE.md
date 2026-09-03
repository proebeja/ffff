# Formatvorlage für Lead NA, Lead PL und die Aufriss-Tabs

`Formatvorlage_Lead.xlsx` legt fest, wie die Ausgabeblätter aussehen und woher sie ihre Zahlen holen. Vier Blätter: **Lead NA**, **Lead PL**, **Mastersheet** und **A_Muster**.

Die Vorlage enthält keine Mandantendaten. Sie zeigt je Zeilentyp eine Musterzeile mit dem fertigen Format und im Blatt A_Muster die Formellogik.

## Was sich gegenüber dem jetzigen Stand ändert

**1. Aufrisse ziehen per SUMIFS statt per Direktverweis.**

Bisher verlinken die Aufriss-Tabs direkt. Künftig zieht jede Kontozeile aus dem Mastersheet über zwei Kriterien, Kontonummer und Klasse:

```
=SUMIFS(Mastersheet!$H$2:$H$436; Mastersheet!$A$2:$A$436;$E9; Mastersheet!$D$2:$D$436;$F9)/1000
```

`$E9` ist die Kontonummer aus der Ticker-Spalte, `$F9` die Klasse. Beide Kriterien kommen aus der Zeile selbst, nie als Text in die Formel geschrieben.

Damit ist das Mastersheet die einzige Datenquelle. Kein Blatt verweist mehr direkt auf ein anderes, außer den Lead-Tabs, die ihre Positionssummen aus den Aufriss-Tabs holen.

**2. Zwei Sprachspalten, umschaltbar.**

Spalte C trägt Deutsch, Spalte D trägt Englisch. Beide bleiben immer befüllt. Die Umschaltung läuft über die Spaltengruppierung: Spalte C hat `outlineLevel = 2` und ist eingeklappt, deshalb ist Englisch sichtbar. Über das Gliederungssymbol schaltet der Leser um.

In openpyxl:

```python
ws.column_dimensions['C'].outlineLevel = 2
ws.column_dimensions['C'].hidden = True
```

Breiten: C = 45.33, D = 37.11. Beide gleich formatiert, damit die Optik unabhängig von der Sprache gleich bleibt.

**3. Ticker in E und F, eingeklappt.**

`outlineLevel = 1`, `hidden = True`, Breite 13, Schrift blau.

In den **Lead-Tabs**: Ticker 1 = Aufriss-Tab, Ticker 2 = Klasse.
In den **Aufriss-Tabs**: Ticker 1 = Kontonummer, Ticker 2 = Klasse.

**4. Kopfzeile mit Projektname und Einheit.**

Zeile 6 trägt `Projekt <Name> — Lead NA`, Zeile 7 die Einheit, deutsch `in TEUR`, englisch `in kEUR`. Beide Zeilen über die volle Breite dunkel hinterlegt, weiße fette Schrift.

**5. Alle Zahlen in Tausend.**

Das Mastersheet führt weiterhin EUR. **Die Division durch 1.000 steht in der Formel**, nicht im Zahlenformat. Also `SUMIFS(...)/1000`.

Nur die Zeilen teilen, die direkt aus dem Mastersheet ziehen. Summenzeilen, die bereits geteilte Zeilen addieren, dürfen nicht noch einmal geteilt werden. Das ist die häufigste Falle bei dieser Umstellung.

## Formatwerte

| Element | Wert |
|---|---|
| Schrift | Arial 8, Kontrollzeilen Arial 9 |
| Kopfzeilen | Füllung `FF17191F`, Schrift weiß fett |
| Zeilenhöhe | 11,4 · Abstandszeilen 4,5 |
| Spalte A | 11,44 (Rand) |
| Spalte B | 0,89 (schmaler Rand) |
| Spalte C / D | 45,33 / 37,11 (Position DE / EN) |
| Spalte E / F | 13,0 (Ticker, eingeklappt) |
| Periodenspalten | 10,44 |
| danach | 0,89 schmal, dann 48 für Kommentare |
| Zahlenformat | `_(* #,##0_);_(* \(#,##0\);_(* "-"?_);_(@_)` |
| Kontrollzeilen | dasselbe mit zwei Nachkommastellen |
| Gitternetz | aus |
| Konten | Einzug 1, grau `FF808080`, `outlineLevel 1`, eingeklappt |
| Summe | fett, dünne Linie oben |
| Blocksumme | fett, Linie oben und unten |
| Fußnote | kursiv, grau |
| summaryBelow | `False`, Position steht über den Konten |

## Umsetzung

Stile aus der Vorlage **kopieren**, nicht aus dieser Beschreibung nachbauen:

```python
import shutil, copy
shutil.copy('Formatvorlage_Lead.xlsx', ziel)
# innerhalb derselben Mappe:
ziel_zelle._style = copy.copy(muster_zelle._style)
```

Theme-Farben und Zeilenhöhen lassen sich aus einer Beschreibung nicht zuverlässig rekonstruieren. Deshalb ist die Vorlage die Quelle, nicht der Text.

## Prompt für Claude Code

```
Hier ist Formatvorlage_Lead.xlsx mit der Beschreibung.

Stell die Ausgabe auf dieses Format um. Drei Änderungen:

1. Die Aufriss-Tabs ziehen ihre Kontozeilen per SUMIFS aus dem Mastersheet
   über Kontonummer und Klasse, nicht mehr per Direktverweis. Beide Kriterien
   kommen aus den Ticker-Spalten der jeweiligen Zeile.

2. Lead NA, Lead PL und alle Aufriss-Tabs bekommen zwei Sprachspalten,
   C deutsch und D englisch, umschaltbar über die Spaltengruppierung.
   Die Ticker liegen in E und F und sind eingeklappt.

3. Kopfzeile mit "Projekt <Name>" und darunter der Einheit, alle Zahlen
   in Tausend. Die Division durch 1.000 steht in der Formel, nicht im
   Zahlenformat, und nur bei Zeilen, die direkt aus dem Mastersheet ziehen.

Kopiere die Stile aus der Vorlage, bau sie nicht aus der Beschreibung nach.

Fang mit EINEM Aufriss-Tab an und zeig ihn mir, bevor du die anderen
umstellst. Danach Lead NA, dann Lead PL, dann der Rest.

Prüfe nach jedem Schritt, dass die Kontrollzeilen auf null stehen.
```

## Worauf zu achten ist

**Der Projektname gehört an eine Stelle.** Er steht in der Kopfzeile jedes Blatts, sollte aber aus einer einzigen Quelle kommen, etwa dem Info-Tab. Sonst steht er nach der dritten Änderung auf zwanzig Blättern unterschiedlich.

**Die Kontrollzeilen müssen nach der Umstellung neu geprüft werden.** Ein Direktverweis und ein SUMIFS liefern nur dann dasselbe, wenn Kontonummer und Klasse im Mastersheet eindeutig sind. Gibt es ein Konto zweimal, etwa in zwei Gesellschaften, summiert SUMIFS beide und der Direktverweis tat es nicht.

**Konten mit abweichender Schreibweise im Ticker liefern still null.** Die Ticker müssen aus dem Mastersheet übernommen werden, nie getippt.
