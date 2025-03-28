1. Output file is created in the script's directory, using the following naming convention "LATUS_enviaTEL_ERZ_Reporting_Entsorgungen_Tabelle_für_PBI_{DDMMYYYY}.xlsx".
2. Input file is selected with a pop-up window.
3. Input worksheet has to be in the first position on the left.
4. The script begins execution from row 7 in the input file.
5. Whenever there is a value in column D other than "Gesamt", it copies values from that row to the output from columns:
- Kategoriegruppe - column "B" (variable is reassigned if there is a change in value),
- Lagerplatz - column "D",
- Haufwerk - column "E",
- Ergebnis Beprobung - column "I",
- Kosten Beprobung - column "K",
- Tonnage - column "G",
- Angebot Preis/t - column "M".
6. Datum (Version) is populated with the current date of execution.
