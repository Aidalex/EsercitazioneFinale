# EsercitazioneFinale_Excel

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
