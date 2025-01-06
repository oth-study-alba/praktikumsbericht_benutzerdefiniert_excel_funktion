# VerteilungBetriebsleistung

```excel
VerteilungBetriebsleistung=LAMBDA(                                // Lambda-Funktion Wrapper um die Parameter der Benutzerdefinierten Excel Funktion festzulegen
    tabelle_produktlinie_verteilung,                              // Parameter 1: Eingabetabelle für die Produktlinienverteilung
    tabelle_filter,                                               // Parameter 2: Eingabetabelle für für eventuelle Filterung nach Linien oder Einzelmaschinenfertigung
    tabelle_prognose,                                             // Parameter 3: Eingabetabelle für Prognostizierte Werte der Anlagenverteilung
    MAKEARRAY(                                                    // Erstellung einer n x n Matrix welche bei Aufruf zurückgegeben wird.
        ANZAHL2(INDEX(tabelle_produktlinie_verteilung, , 1)) + 1, // Anzahl der Zeilen in der Produktlinienverteilungstabelle + 1 um Platz für die Jahreszahlen in der ersten Reihe zu haben
        ANZAHL2(INDEX(tabelle_prognose, , 1)) + 1,                // Anzahl der Spalten in der Prognosetabelle + 1 um Platz zu haben für die Bezeichnungen der Produktlinien in der ersten Spalte
        LAMBDA(zeile_produktlinie, spalte_jahr_ae,                // Lambda-Funktion Wrapper welcher für MAKERARRAY benötigt wird.
            WENN(
                (zeile_produktlinie = 1) * (spalte_jahr_ae = 1), "", // Lässt die Oberste Linke Spalte Leer
                WENN(
                    zeile_produktlinie = 1, 
                    INDEX(tabelle_prognose, spalte_jahr_ae - 1, 1) + 1, // Schreibt die Betriebsleistungsjahre welche um Eins höher sind als Auftragseingangsjare in die erste Zeile
                    WENN(
                        spalte_jahr_ae = 1, 
                        INDEX(tabelle_produktlinie_verteilung, zeile_produktlinie - 1, 1), // Schreibt die Bezeichnungen der Produktlinie in Parameter 1 in die erste Spalte
                        SUMMENPRODUKT( // Berechne das Summenprodukt von zwei Arrays um den jeweiligen Produktlinienwert pro Jahr zu erhalten.
                            INDEX(
                                INDEX( // Rückgabewert ein Array aus der Prognose Tabelle ohne Spalten und Reihen Bezeichnung zum Rechnen
                                    tabelle_prognose,
                                    SEQUENZ(ANZAHL2(INDEX(tabelle_prognose, , 1))), 
                                    SEQUENZ(1, 4, 2) 
                                ),
                                XVERGLEICH( // Sucht die Zeile welche die Werte des aktuell zu berechnenden Jahres enthält
                                    INDEX(tabelle_prognose, spalte_jahr_ae - 1, 1), // Vergleiche den Prognosewert für die Spalte
                                    INDEX(tabelle_prognose, 0, 1) // Mit der ersten Spalte der Prognosetabelle
                                ),
                                0
                            ),
                            INDEX(
                                INDEX( // Rückgabewert ein Array aus der Produktlinien Verteilungs Tabelle ohne Spalten und Reihen Bezeichnung zum Rechnen
                                    tabelle_produktlinie_verteilung,
                                    SEQUENZ(ANZAHL2(INDEX(tabelle_produktlinie_verteilung, , 1))), // Sequenz der Zeilen in der Produktlinienverteilungstabelle
                                    SEQUENZ(1, 4, 2) // Sequenz der Spalten in der Produktlinienverteilungstabelle
                                ),
                                XVERGLEICH( // Sucht die Zeile welche die Werte der aktuell zu berechnenden Produktlinie enthält
                                    INDEX(tabelle_produktlinie_verteilung, zeile_produktlinie - 1, 1), // Vergleiche den Produktlinienverteilungswert für die Zeile
                                    INDEX(tabelle_produktlinie_verteilung, 0, 1) // Mit der ersten Spalte der Produktlinienverteilungstabelle
                                ),
                                0
                            ) * 
                            INDEX( // Multiplikation mit dem jeweiligen Faktor der Produktlinie falls eine Filterung nach Einzelmaschinen oder Linien benötigt wird.
                                INDEX( // Rückgabewert ein Array aus der Filter Tabelle ohne Spalten und Reihen Bezeichnung zum Rechnen
                                    tabelle_filter,
                                    SEQUENZ(ANZAHL2(INDEX(tabelle_filter, , 1))), // Sequenz der Zeilen in der Filtertabelle
                                    SEQUENZ(1, 1, 2) // Sequenz der Spalten in der Filtertabelle
                                ),
                                XVERGLEICH( // Sucht die Zeile welche die Werte der aktuell zu berechnenden Produktlinie enthält
                                    INDEX(tabelle_produktlinie_verteilung, zeile_produktlinie - 1, 1), // Vergleiche den Produktlinienverteilungswert für die Zeile
                                    INDEX(tabelle_filter, 0, 1) // Mit der ersten Spalte der Filtertabelle
                                ),
                                0
                            ) / 10000 // Teilt das Ergebniss um den Faktor 1000 da die Werte welche miteinander Verechnet werden alle Prozentzahlen im Formato von Dezimalzahlen <1 sind
                        )
                    )
                )
            )
        )
    )
)
