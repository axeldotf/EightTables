# README per EightTables
## Descrizione

Questo script Python è progettato per elaborare un file Excel contenente dati relativi a canali e ARFCN (Absolute Radio Frequency Channel Number) utilizzati per la gestione delle reti mobili LTE e GSM. L'obiettivo principale è di riorganizzare i dati, mappare i canali e ARFCN agli operatori e bande di frequenza, e salvare i risultati in file Excel distinti per ogni foglio del file originale.

## Funzionalità

1. **Selezione del file Excel**: Lo script richiede all'utente di selezionare un file Excel tramite una finestra di dialogo.
2. **Elaborazione dei fogli**: Per ogni foglio (eccetto 'Foglio1') del file Excel selezionato:
   - Rinomina le colonne secondo una mappatura predefinita.
   - Mappa i valori unici di 'Ch' (canale) e 'ARFCN' (numero di canale radio assoluto) agli operatori e bande di frequenza corrispondenti.
   - Calcola e aggrega statistiche sui dati.
   - Salva i dati elaborati in nuovi file Excel, uno per ogni foglio originale.
3. **Barra di avanzamento**: Mostra una barra di avanzamento per monitorare il progresso dell'elaborazione.

## Dipendenze

Questo script richiede le seguenti librerie Python:
- `pandas`
- `openpyxl`
- `tqdm`
- `tkinter` (incluso nella libreria standard di Python)

Per installare le librerie necessarie, è possibile utilizzare `pip`:
```sh
pip install pandas openpyxl tqdm
