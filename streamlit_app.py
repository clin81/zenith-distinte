import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from streamlit_gsheets import GSheetsConnection

# --- CONFIGURAZIONE PAGINA ---
st.set_page_config(page_title="Zenith Prato - Gestione Distinte", layout="wide")
TEMPLATE_FILE = "distinta_vuota.xlsx"

# --- CONNESSIONE GOOGLE SHEETS ---
conn = st.connection("gsheets", type=GSheetsConnection)

def carica_db():
    try:
        df = conn.read(ttl="0")
        colonne_necessarie = ["Nominativo", "Tipo", "Maglia", "GG", "MM", "AA", "FIGC"]
        if df is None or df.empty:
            return pd.DataFrame(columns=colonne_necessarie)
        
        # Sincronizza colonne mancanti
        for col in colonne_necessarie:
            if col not in df.columns:
                df[col] = ""
        
        # SBLOCCO MODIFICA MAGLIA: Forza il tipo numerico per l'editor
        df['Maglia'] = pd.to_numeric(df['Maglia'], errors='coerce')
        
        return df[colonne_necessarie]
    except Exception as e:
        st.error(f"Errore caricamento: {e}")
        return pd.DataFrame(columns=["Nominativo", "Tipo", "Maglia", "GG", "MM", "AA", "FIGC"])

def salva_db(df):
    try:
        df = df.fillna("") # Pulizia celle vuote
        conn.update(data=df)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Errore salvataggio: {e}")
        return False

def safe_write(ws, cell_coord, value):
    from openpyxl.cell.cell import MergedCell
    cell = ws[cell_coord]
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell_coord in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                return
    else:
        cell.value = value

def compila_template(players_df, staff_df, info):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active 
    
    # Scrittura Intestazione
    safe_write(ws, 'G7', f"Zenith Prato S.S.D.R.L. Vs {info['avversario']}")
    safe_write(ws, 'G8', f"Data: {info['data']} - Ora: {info['ora']}")
    safe_write(ws, 'G9', info['campo'])

    # Scrittura Giocatori (Riga 12)
    r_idx = 12 
    for _, row in players_df.iterrows():
        safe_write(ws, f'C{r_idx}', row.get('Maglia', ''))
        safe_write(ws, f'D{r_idx}', row.get('
