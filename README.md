# EsercitazioneFinale_Excel

**IT**  
Il progetto di seguito si basa nell'elaborare un file excel per la regiona Marche, la quale ha richiesto di sviluppare un sistema per la gestione delle strutture ricettive del territorio. 
L'obiettivo è stato quello di rielaborare un file Excel che potesse essere pronto all'uso per l'utente finale per ricerche puntuali sulle strutture ricettive.  
Per lo sviluppo sono stati messi a disposizione due file:
1. elencostrutture.xlsx contenente lʼelenco di tutte le informazioni sulle strutture (categoria, nome, indirizzo, ecc.);
2. prezzimedi.csv contenente lʼelenco dei prezzi medi delle strutture per città.

Per iniziare è stato necessario pulire il file con Power Query e successivamente applicare delle condizioni per la pulizia dei dati (tra cui: ricerca di dati duplicati, analizzare eventuali righe vuote, maiuscole/minuscole, rimuovere gli spazi iniziali, eliminare le informazioni ripetute in alcune celle come la via, sostituire il numero di telefono con la mail e viceversa ove necessario, etc), sono stati eseguiti merge per dei controlli incrociati e in aggiunta è stato applicato un indice alle strutture, in quanto esistono strutture con il medesimo nome in diverse città.

È stato chiesto di:
1. Creare un menù a tendina, all'interno del quale scegliendo il Nome Struttura, si avrà il numero totale delle strutture della regione e il numero di strutture presenti nella città della Struttura CERCATA (attraverso l'indice applicato vengono restituite tutte le strutture con lo stesso nome, ma presenti nelle diverse città, senza veniva restituita solo la prima struttura..  
2. Creare una tabella pivot per mostrare il numero di strutture per CATEGORIA filtrabile per CITTÀ.
3. Caricare la tabella presente nel foglio PREZZI del file PREZZIMEDI. Creare la relazione tra le due tabelle attraverso Power Pivot.


**EN**  
This project involved developing a user-friendly Excel-based system for managing accommodation facilities in Italy's Marche region. The goal was to transform raw data into an interactive tool for targeted searches on hospitality structures.

*Data Sources*:
1. elencostrutture.xlsx: Facility details (category, name, address, etc.)
2. prezzimedi.csv: Average prices by city

*Data Processing*:
Performed comprehensive cleaning using Power Query transformations:
1. Standardized text fields (trim, case consistency)
2. Removed duplicates & empty rows
3. Created unique facility IDs (for same-name facilities in different cities)
4. Cross-validated contact info (phone ↔ email)
5. Merged datasets for relational analysis

*Key Features Developed*:
1. Dynamic Dropdown Menu:
    - Search by facility name to view:
    - Total regional facilities
    - Facilities in the selected city (using the index to show all matching names)
2. Interactive Pivot Table:
    - Facilities by category
    - City-based filtering
3. Data Integration:
    - Imported price data from PREZZIMEDI
    - Established table relationships via Power Pivot

**Technical Stack:**  

- 🔄 **Excel** Advanced, Power Query, Power Pivot
