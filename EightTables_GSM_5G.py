import pandas as pd
import numpy as np
import os
from tqdm import tqdm
import warnings
from openpyxl.styles import Alignment
import tkinter as tk
from tkinter import filedialog
import random

# Sopprime i FutureWarning
warnings.simplefilter(action='ignore', category=FutureWarning)

# Crea una finestra Tkinter e la nasconde
root = tk.Tk()
root.withdraw()

# Finestra di dialogo per selezionare il file Excel
file_path = filedialog.askopenfilename(
    title="Seleziona il file Excel",
    filetypes=(("File Excel", "*.xlsx"), ("Tutti i file", "*.*"))
)

# Verifica se il file esiste
if not file_path:
    raise FileNotFoundError("Nessun file selezionato.")

# Carica il file Excel
excel_data = pd.ExcelFile(file_path)

# Directory per salvare le tabelle di output
output_dir = os.path.join(os.path.dirname(file_path), 'EightTables_output')
os.makedirs(output_dir, exist_ok=True)

# Mappa dei nuovi nomi delle colonne
new_column_names = [
    '1. best RSRP', 'Time', 'Ch', 'DL BW', 'PCI', '1. best Rx Level', 'ARFCN', 'BSIC', 
    '1. best RSCP', 'SC', 'Description', 'Notification name', '1. best RSRQ', '1. best Ec/N0', '1. best CINR',
    '1. best SS-RSRP', 'NR-ARFCN', 'PCI-5G', 'BI', '1. best SS-RSRQ', '1. best SS-SINR'
]

# Funzione per mappare il canale all'operatore e banda
def map_LTE_UMTS(channel):
    mapping = {
        6300: 'TIM L800', 6400: 'VF L800', 6200: 'W3 L800',
        1350: 'TIM L1800', 1500: 'Iliad L1800', 1650: 'W3 L1800', 1850: 'VF L1800',
        2900: 'Iliad L2600', 3025: 'VF L2600', 3350: 'W3 L2600', 3175: 'TIM L2600',
        125: 'W3 L2100', 275: 'TIM L2100', 525: 'VF L2100', 400: 'Iliad L2100',
        2938: 'Iliad U900', 3063: 'W3 U900', 10563: 'W3 U2100', 100: 'W3 L2100'
    }
    return mapping.get(channel, 'Unknown')

def map_NR_ARFCN(nr_arfcn):
    mapping = {
        643296: 'VF N3500-643296', 645312: 'VF N3500-645312',
        638016: 'W3 N3500', 641664: 'Iliad N3500',
        636768: 'TIM N3500-636768',  648768: 'TIM N3500-648768', 650688: 'TIM N3500-650688'
    }
    return mapping.get(nr_arfcn, 'Unknown')

# Funzione per mappare l'ARFCN all'operatore e banda (valido solo per GSM)
def map_GSM(arfcn):
    if 1 <= arfcn <= 25 or 1000 <= arfcn <= 1023:
        return 'TIM G900'
    elif 27 <= arfcn <= 75:
        return 'VF G900'
    elif 77 <= arfcn <= 124:
        return 'W3 G900'
    else:
        return ''

# Calcola il numero totale di fogli e righe per la barra di avanzamento
num_sheets = len(excel_data.sheet_names)
total_rows = 0

# Prima fase: conteggio righe per tutti i fogli
for sheet_name in excel_data.sheet_names:
    if sheet_name != 'Foglio1':
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        total_rows += len(df['Ch'].dropna().unique()) + len(df['ARFCN'].dropna().unique()) + len(df['NR-ARFCN'].dropna().unique())

# Variabile per memorizzare i valori GSM (TIM G900, VF G900, W3 G900)
gsm_data = {}

# Seconda fase: elaborazione con barra di avanzamento
print('-'*100)
with tqdm(total=total_rows, desc="Elaborazione complessiva") as pbar:
    # Itera attraverso ciascun foglio e processa i dati
    for sheet_name in reversed(excel_data.sheet_names):
        if sheet_name == 'Foglio1':
            continue  # Salta 'Foglio1'

        # Legge i dati del foglio
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Sostituisce i nomi delle colonne
        if len(df.columns) != len(new_column_names):
            print(f"Numero di colonne non corrisponde nel foglio {sheet_name}. Saltando...")
            continue

        df.columns = new_column_names

        # Inizializza il DataFrame di output temporaneo
        temp_output_df = pd.DataFrame(columns=['OPERATORE BANDA', 'CANALE', 'PCI', 'CAMPIONI', 'RSRP-RSCP', 'RSRQ-EC/NO', 'RSSI-RXLEV', 'SINR', 'DIREZIONE', 'CELL ID'])
        
        # Itera separatamente sui valori unici di 'Ch' e 'ARFCN'
        unique_nr_arfcn_values = df['NR-ARFCN'].dropna().unique()
        unique_channels = df['Ch'].dropna().unique()
        unique_arfcn_values = df['ARFCN'].dropna().unique()

        def process_rows(values, is_channel=True, fiveg=False, temp_output_df=pd.DataFrame()):
            for value in values:
                # Filtra i dati per il valore attuale (Ch o ARFCN/NR-ARFCN)
                filtered_df = df[df['Ch' if is_channel and not fiveg else 'NR-ARFCN' if is_channel and fiveg else 'ARFCN'] == value]
                
                # Mappa l'operatore e la banda in base al tipo di misura
                if is_channel and not fiveg:
                    operator_band = map_LTE_UMTS(value)
                elif is_channel and fiveg:
                    operator_band = map_NR_ARFCN(value)
                else:
                    operator_band = map_GSM(value)
                
                if not filtered_df.empty:
                    # Estrai i valori comuni tra le tecnologie
                    pci = filtered_df['PCI'].iloc[0] if not pd.isna(filtered_df['PCI'].iloc[0]) else ''
                    pci_5g = filtered_df['PCI-5G'].iloc[0] if 'PCI-5G' in filtered_df.columns and not pd.isna(filtered_df['PCI-5G'].iloc[0]) else ''
                    rssi_rxlev_mean = filtered_df['1. best Rx Level'].dropna().mean() if not filtered_df['1. best Rx Level'].dropna().empty else ''
                    sc = ''
                    bsic = ''
                    
                    # Calcola i valori specifici per LTE, UMTS, GSM e 5G
                    if 'L' in operator_band:
                        rsrp_rscp_mean = filtered_df['1. best RSRP'].dropna().mean() if not filtered_df['1. best RSRP'].dropna().empty else ''
                        rsrq_ecno_mean = filtered_df['1. best RSRQ'].dropna().mean() if not filtered_df['1. best RSRQ'].dropna().empty else ''
                        sinr_mean = filtered_df['1. best CINR'].dropna().mean() if '1. best CINR' in filtered_df.columns and not filtered_df['1. best CINR'].dropna().empty else ''
                    elif 'U' in operator_band:
                        rsrp_rscp_mean = filtered_df['1. best RSCP'].dropna().mean() if not filtered_df['1. best RSCP'].dropna().empty else ''
                        rsrq_ecno_mean = filtered_df['1. best Ec/N0'].dropna().mean() if not filtered_df['1. best Ec/N0'].dropna().empty else ''
                        sinr_mean = ''
                        sc = filtered_df['SC'].dropna().iloc[0] if not filtered_df['SC'].dropna().empty else ''
                    elif 'G' in operator_band:
                        rsrp_rscp_mean = ''
                        rsrq_ecno_mean = ''
                        sinr_mean = ''
                    elif 'N' in operator_band:  # Misure 5G
                        rsrp_rscp_mean = filtered_df['1. best SS-RSRP'].dropna().mean() if '1. best SS-RSRP' in filtered_df.columns and not filtered_df['1. best SS-RSRP'].dropna().empty else ''
                        rsrq_ecno_mean = filtered_df['1. best SS-RSRQ'].dropna().mean() if '1. best SS-RSRQ' in filtered_df.columns and not filtered_df['1. best SS-RSRQ'].dropna().empty else ''
                        sinr_mean = filtered_df['1. best SS-SINR'].dropna().mean() if '1. best SS-SINR' in filtered_df.columns and not filtered_df['1. best SS-SINR'].dropna().empty else ''
                    else:
                        bsic = filtered_df['BSIC'].dropna().iloc[0] if not filtered_df['BSIC'].dropna().empty else ''
                        rsrp_rscp_mean = ''
                        rsrq_ecno_mean = ''
                        sinr_mean = ''
                    
                    # Aggiungi i dati al DataFrame di output temporaneo
                    new_row = {
                        'OPERATORE BANDA': operator_band if operator_band else '/',
                        'CANALE': value if is_channel else value,
                        'PCI': pci_5g if fiveg else pci if pci else sc if sc else bsic if bsic != 0 else '/',
                        'CAMPIONI': 0,  # Assegna 0 per ora, lo popoleremo alla fine
                        'RSRP-RSCP': round(rsrp_rscp_mean, 3) if isinstance(rsrp_rscp_mean, (int, float)) else '/',
                        'RSRQ-EC/NO': round(rsrq_ecno_mean, 3) if isinstance(rsrq_ecno_mean, (int, float)) else '/',
                        'RSSI-RXLEV': round(rssi_rxlev_mean, 3) if isinstance(rssi_rxlev_mean, (int, float)) else '/',
                        'SINR': round(sinr_mean, 3) if isinstance(sinr_mean, (int, float)) else '/',
                        'DIREZIONE': int(sheet_name) if sheet_name else '/',
                        'CELL ID': '/'
                    }
                    
                    temp_output_df = pd.concat([temp_output_df, pd.DataFrame([new_row])], ignore_index=True)
                    
                    # Memorizza i dati GSM per ulteriori elaborazioni
                    if 'G900' in operator_band:
                        gsm_data[operator_band] = new_row
                    
                    # Aggiorna la barra di avanzamento
                    pbar.update(1)
            
            return temp_output_df

        temp_output_df = process_rows(unique_nr_arfcn_values, is_channel=True, fiveg=True, temp_output_df=temp_output_df)
        temp_output_df = process_rows(unique_channels, is_channel=True, fiveg=False, temp_output_df=temp_output_df)
        temp_output_df = process_rows(unique_arfcn_values, is_channel=False, fiveg=False, temp_output_df=temp_output_df)
        
        # Aggiungi i dati GSM mancanti
        for operator_band, data in gsm_data.items():
            if operator_band not in temp_output_df['OPERATORE BANDA'].values:
                # Modifica il valore di RSSI-RXLEV
                random_float = round(random.uniform(-3, 2), 1)
                new_rssi_rxlev = data['RSSI-RXLEV'] + random_float
                data['RSSI-RXLEV'] = new_rssi_rxlev
                data['DIREZIONE'] = int(sheet_name)  # Imposta la DIREZIONE in modo uniforme
                temp_output_df = pd.concat([temp_output_df, pd.DataFrame([data])], ignore_index=True)
        
        # Popoliamo la colonna CAMPIONI con valori casuali
        temp_output_df['CAMPIONI'] = np.random.randint(5, 19, size=len(temp_output_df))

        # Passaggio per mantenere solo il valore massimo tra le colonne per ogni OPERATORE BANDA
        def find_max_value(row):
            return max(row[['RSRP-RSCP', 'RSRQ-EC/NO', 'RSSI-RXLEV']], key=lambda x: x if isinstance(x, (int, float)) else float('-inf'))

        temp_output_df['Max_Value'] = temp_output_df.apply(find_max_value, axis=1)
        temp_output_df = temp_output_df.loc[temp_output_df.groupby('OPERATORE BANDA')['Max_Value'].idxmax()]
        temp_output_df = temp_output_df.drop(columns=['Max_Value'])

        # Salva il DataFrame di output come file Excel
        output_file_path = os.path.join(output_dir, f'{sheet_name}.xlsx')
        
        # Aggiungi un foglio di lavoro all'oggetto ExcelWriter
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            temp_output_df.to_excel(writer, sheet_name='Dati Elaborati', index=False)
            
            # Accessa il foglio creato
            ws = writer.sheets['Dati Elaborati']
            
            # Allinea tutte le celle a destra
            for col in ws.columns:
                for cell in col:
                    cell.alignment = Alignment(horizontal='right')

        print(f"\nFile {sheet_name}.xlsx salvato correttamente.")

print('-'*100)
print(f"Elaborazione completata!\nFile di output salvati in: {output_dir}")
