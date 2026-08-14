# Übergabe an Claude Code — Kitchenstories, zweiter Lauf

Der erste Lauf war strukturell gut. 1.724 Formeln ohne Fehler, alle Kontrollzeilen auf null, saubere Kette Lead zu Aufriss zu Mastersheet, Reconciliation FY2023 findet die offene Differenz von 172 und trennt sie von den erklärten Abstimmposten, Benchmark-Tab gebaut. Das bleibt alles so.

Es gibt fünf Befunde. Vier davon sind keine Fehler in deiner Umsetzung, sondern Lücken in der Hausconvention, die jetzt in Version 2.8 geschlossen sind. Lies zuerst `_aenderungen_v2_8`.

## Was zu ändern ist

**1. Seitenwechsel, neuer Block `seitenwechsel`.** Konto 701 Cash-Pool steht FY2022 mit (3.604.535) passivisch und FY2023 mit 208.710 aktivisch. Du hast den Pfad aus FY2023 genommen, damit steht FY2022 mit minus 3,6 Mio unter den Forderungen. Genau dieser Betrag erscheint in der Recon FY2022 auf beiden Seiten. Setze die neue Regel um: bei den definierten Pfadpaaren folgt die Bilanzseite je Periode dem Vorzeichen, die Klasse bleibt fest. Betroffen sind hier mindestens 701, 1600 (FY2024 aktivisch mit 382) und 1731 2.

Setze im selben Zug `verhaltenspruefung.kriterien.vorzeichenwechsel` um. Das steht seit v2.6 in der Konvention und war nicht implementiert, sonst wäre ein Wechsel über 3,8 Mio in der Review-Queue aufgetaucht.

**2. Jahresergebnis ins Eigenkapital, neuer Block `lead_na_eigenkapital`.** Das Eigenkapital im Lead NA steht FY2023 bei (3.097.578), im Abschluss bei 6.079. Der Jahresfehlbetrag hängt als nachrichtliche Zeile daneben. Die Spalte trägt aber den Status abschlusstreu. Das Ergebnis wird eine echte Eigenkapitalzeile, gespeist aus der Summe der Klasse PL, und die Kontrollzeile ändert sich entsprechend.

**3. Ungelöste Konten, neue QA-Regel A6.** In FY2024 hingen 349.821 in der Review-Zeile außerhalb jeder Bilanzposition, bei einer Bilanzsumme von 1,2 Mio. Die Kontrollzeile ging nur auf, weil sie diese Zeile mitgerechnet hat. Ab sofort bekommt jedes Konto einen vorläufigen Pfad und landet in einer Position. Kontrollzeilen prüfen, sie gleichen nicht aus.

**4. Saldenvorträge, geänderter Block `technische_konten`.** In YTD Jul25 lagen 2.720.068 auf den DATEV-Saldenvortragskonten 9000 bis 9009 und fielen als TECH heraus. Dadurch ist das Eigenkapital dieser Spalte identisch zur Vorperiode, also nicht fortgeschrieben. Saldenvortragskonten sind keine Statistikkonten. Die Regel ab 9000 gleich TECH ist aufgeteilt.

Richtig gemacht hast du 9960 Bewertungskorrektur, das ist korrekt bei den Forderungen aL gelandet und nicht als TECH gestorben. Der Vorrang der Typ-1-Regeln steht jetzt ausdrücklich in der Konvention.

**5. Regelgruppe Zahlungsverkehr, neuer Block `regelgruppe_zahlungsverkehr`.** Acht der zehn Benchmark-Abweichungen waren derselbe Typ. Kreditkarten, Spendesk, Shopify, Finway, Forderungen gegen Geschäftsführer landeten in Review statt in Net Debt, deshalb steht die Zeile „davon ND" in allen Perioden auf null. Drei Regelgruppen sind jetzt kodiert. Drei weitere Fälle sind bewusst nicht kodiert und bleiben Review mit Pflichtfrage, das ist kein Versehen.

## Was zusätzlich zu liefern ist

Ein eigener Tab mit dem Namen **QA**. Er hat im ersten Lauf gefehlt, die Befunde lagen verstreut in „Info" und „Status je Spalte". Beide Tabs bleiben, ersetzen den QA-Tab aber nicht. Der QA-Tab weist jede Einzelprüfung A1 bis A6, B1 bis B5, C1 bis C5 und D1 mit bestanden oder nicht bestanden aus. Von diesen Prüfungen wurden im ersten Lauf mehrere gar nicht berichtet, insbesondere B1, also ob die Monatsspalten kumuliert oder periodisch sind.

Alle Differenz- und Restspalten auf zwei Nachkommastellen runden. In der Recon standen Werte wie minus 2,9 mal zehn hoch minus elf in einer Spalte, die man auf null liest.

## Prüfungen, die dieser Lauf bestehen muss

1. Recon FY2022 gegen den Prüfbericht: Forderungen und Verbindlichkeiten gehen auf, die Differenz von 3.605.431 auf beiden Seiten ist verschwunden
2. Eigenkapital FY2023 im Lead NA entspricht dem Abschluss mit 6.079, Restdifferenz nur die bekannten 172
3. Eigenkapital YTD Jul25 ist gegenüber FY2024 fortgeschrieben und nicht identisch
4. Keine Kontrollzeile enthält eine Review-, Ergebnis- oder Technikzeile als Ausgleichsposten
5. Net Debt enthält die Zahlungsdienstleister- und Kreditkartenkonten, die Zeile „davon ND" steht nicht mehr durchgängig auf null
6. Konto 701 erscheint in FY2022 auf der Passivseite und in FY2023 auf der Aktivseite, mit Vermerk und Frage
7. Der QA-Tab existiert und weist jede Einzelprüfung aus

## Zeigen

Recon FY2022 und FY2023, den Eigenkapitalblock des Lead NA über alle vier Spalten, den QA-Tab vollständig, die Review-Queue, den Benchmark-Tab mit der neuen Abweichungszahl, und eine Liste der angewandten Seitenwechsel.

Dieselben vier Eingangsdateien wie beim ersten Lauf. Die Hausconvention v2.8 ersetzt die bisherige Datei.
