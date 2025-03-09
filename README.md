# EightTables 2G/3G/4G/5G Data Processing

## Introduzione
Questo progetto è stato sviluppato in collaborazione con **Selektra Italia S.r.l.** e consiste in un sistema automatizzato per la conversione ed elaborazione di dati provenienti da misurazioni GSM (2G), UMTS (3G), LTE (4G) e 5G effettuate sul campo. Lo scopo principale del progetto è generare report strutturati per una rapida analisi tecnica.

## Struttura del Progetto

```
EightTables_Project/
├── main.py
└── functions/
    ├── EightTables_GSM.py
    └── EightTables_GSM_5G.py
```

- **`EightTables_GSM.py`**: gestisce la conversione ed elaborazione dati GSM (2G), UMTS (3G) e LTE (4G).
- **`EightTables_GSM_5G.py`**: estende il primo script aggiungendo il supporto alle misure 5G.
- **`main.py`**: script principale per selezionare quale elaborazione effettuare.

## Requisiti
- Python 3.x
- Librerie Python necessarie:
  - pandas
  - numpy
  - openpyxl
  - tqdm
- tkinter (incluso in Python standard)

## Installazione delle Dipendenze
Utilizzare il seguente comando per installare tutte le dipendenze necessarie:
```bash
pip install pandas numpy openpyxl tqdm
```

## Utilizzo
Lanciare l'applicazione tramite il file `main.py` con:
```bash
python main.py
```
All'avvio verrà richiesto tramite interfaccia grafica di selezionare il file Excel da elaborare. Successivamente, l'utente sceglierà se elaborare dati esclusivamente GSM, UMTS e LTE oppure includere anche il 5G.

Lo script produrrà automaticamente degli output Excel suddivisi per ogni direzione rilevata, salvati nella cartella `EightTables_output` posizionata nella stessa directory del file di input selezionato.

## Formato Dati Input/Output
- **Input**: file Excel con dati GSM, UMTS, LTE e/o 5G da elaborare, divisi in fogli corrispondenti alle direzioni.
- **Output**: file Excel riepilogativi con colonne:
  - OPERATORE BANDA
  - CANALE
  - PCI
  - CAMPIONI
  - RSRP-RSCP
  - RSRQ-EC/NO
  - RSSI-RXLEV
  - SINR
  - DIREZIONE
  - CELL ID

## Struttura del Codice
- Mappatura canali e frequenze agli operatori corrispondenti.
- Aggregazione dati medi per tecnologia (GSM, UMTS, LTE, 5G).
- Calcolo valori medi di misure (RSRP, RSCP, RSSI, SINR, ecc.).

## Collaborazione
Questo progetto è stato realizzato in stretta collaborazione con l'azienda:

**Selektra Italia S.r.l.**

## Licenza
Questo progetto è distribuito sotto la licenza **Apache 2.0**. Per ulteriori dettagli consultare il file [LICENSE](LICENSE).
