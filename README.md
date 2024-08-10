# EightTables Data Processing Script

## Gerarchia

EightTables/
│
├── tab_originale.xlsx     # File excel contenente i dati da elaborare
├── script.py              # Script principale per l'elaborazione dei dati
├── README.md              # Questo file
├── LICENSE                # Licenza
└── EightTables_output/    # Directory in cui verranno salvati i risultati elaborati

## Descrizione

Questo script Python è stato creato per elaborare dati da file Excel contenenti informazioni sui segnali di telecomunicazione. Lo script elabora i dati presenti in diversi fogli del file Excel, generando tabelle di output per ciascun foglio e salvandole in una directory dedicata.

Il processo di elaborazione include:
- Mappatura di canali e ARFCN (Absolute Radio Frequency Channel Number) agli operatori e bande corrispondenti.
- Calcolo delle medie di vari parametri di segnale come RSRP, RSCP, Rx Level, Ec/N0 e SINR.
- Selezione dei valori massimi per ciascun operatore e banda.
- Esportazione dei risultati elaborati in file Excel separati per ciascun foglio elaborato.

## Funzionalità

- **Mappatura automatica**: Lo script associa automaticamente i canali e gli ARFCN ai rispettivi operatori e bande.
- **Elaborazione avanzata**: Per ogni foglio, vengono calcolate le medie dei principali parametri di segnale, che vengono poi filtrati e salvati.
- **Barra di avanzamento**: Una barra di avanzamento mostra il progresso complessivo dell'elaborazione.
- **Supporto per dati GSM**: Lo script gestisce dati GSM e li integra nei risultati finali.

## Prerequisiti

### Librerie necessarie

Assicurati di avere installato Python e le seguenti librerie prima di eseguire lo script:

- `pandas`
- `numpy`
- `openpyxl`
- `tqdm`
- `tkinter`

Puoi installare tutte le dipendenze necessarie utilizzando il comando:

```bash
pip install pandas numpy openpyxl tqdm tk
```

### Formato del File di Input

Il file Excel da fornire in input deve seguire un formato specifico. Ogni foglio all'interno del file deve contenere tabelle con le misurazioni da elaborare. 

- Ogni foglio deve essere denominato in modo da riflettere la direzione relativa delle misurazioni.
- Lo script elaborerà i fogli partendo dall'ultimo e risalendo fino al primo, con l'eccezione del foglio denominato "Foglio1", che verrà escluso dall'elaborazione.
- È fondamentale che i dati all'interno di ciascun foglio siano strutturati correttamente e che le colonne rispettino i nomi previsti dallo script, per garantire un'elaborazione accurata.

Assicurati che il file di input sia configurato correttamente per evitare errori durante l'elaborazione.
