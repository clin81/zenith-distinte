import streamlit as st
import pandas as pd
import os
from io import BytesIO
from openpyxl import load_workbook

# 1. CONFIGURAZIONE
st.set_page_config(page_title="Zenith Prato - Distinte", layout="wide")

DB_FILE = "Database_Tesserati.csv"
TEMPLATE_FILE = "distinta_vuota.xlsx"

def carica_db():
    if not os.path.exists(DB_FILE):
        # Se il file non esiste, restituiamo un DF vuoto con le colonne giuste
        return pd.DataFrame(columns=["Tipo", "Maglia", "GG", "MM", "AA", "Nominativo", "FIGC"])
    try:
        return pd.read_csv(DB_FILE, dtype=str).fillna("")
    except:
        return pd.DataFrame(columns=["Tipo", "Maglia", "GG", "MM", "AA", "Nominativo", "FIGC"])

def safe_write(ws, cell_coord, value):
    """Scrive in una cella gestendo i blocchi di celle unite."""
    from openpyxl.cell.cell import MergedCell
    cell = ws[cell_coord]
    if isinstance(cell, MergedCell):
        # Se la cella è unita, cerchiamo la cella 'madre' del blocco
        for merged_range in ws.merged_cells.ranges:
            if cell_coord in merged_range:
                # Scriviamo nella cella in alto a sinistra del range unito
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                return
    else:
        cell.value = value

def compila_template(players_df, staff_df, info):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active 

    # --- DATI GARA ---
    # 1. Avversario in G7 (mantiene il testo fisso "Vs")
    testo_precedente = ws['G7'].value if ws['G7'].value else "Zenith Prato S.S.D.R.L. Vs "
    safe_write(ws, 'G7', f"{testo_precedente} {info['avversario']}")
    
    # 2. Data e Ora insieme in G8
    safe_write(ws, 'G8', f"Data: {info['data']} - Ora: {info['ora']}")
    
    # 3. Campo in G9
    safe_write(ws, 'G9', info['campo'])

    # 4. Altri dati (opzionali, se vuoi mantenerli anche in B7/B8)
    safe_write(ws, 'B7', f"Gara: ZENITH PRATO vs {info['avversario']}")
    safe_write(ws, 'B8', f"Data: {info['data']}")

    # --- GIOCATORI: INIZIO RIGA 12 ---
    r_idx = 12 
    for _, row in players_df.iterrows():
        safe_write(ws, f'C{r_idx}', row['Maglia'])
        safe_write(ws, f'D{r_idx}', row['GG'])
        safe_write(ws, f'E{r_idx}', row['MM'])
        safe_write(ws, f'F{r_idx}', row['AA'])
        safe_write(ws, f'G{r_idx}', row['Nominativo'])
        safe_write(ws, f'I{r_idx}', row['FIGC'])
        r_idx += 1

    # --- STAFF: INIZIO RIGA 39 ---
    s_idx = 39
    for _, row in staff_df.iterrows():
        safe_write(ws, f'C{s_idx}', row['Maglia']) 
        safe_write(ws, f'G{s_idx}', row['Nominativo'])
        safe_write(ws, f'I{s_idx}', row['FIGC'])
        s_idx += 1

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# --- INTERFACCIA STREAMLIT ---
st.title("⚽ Zenith Prato - Generatore Distinte")

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"❌ Errore: Il file '{TEMPLATE_FILE}' non è presente su GitHub. Caricalo per continuare.")
else:
    if 'data' not in st.session_state:
        st.session_state.data = carica_db()

    # Sidebar per i dettagli della gara
    st.sidebar.header("Dati Gara")
    info = {
        "avversario": st.sidebar.text_input("Avversario", "SQUADRA OSPITE"),
        "data": st.sidebar.text_input("Data", "14/04/2026"),
        "ora": st.sidebar.text_input("Ora", "10:30"),
        "campo": st.sidebar.text_input("Campo", "Chiavacci")
    }

    df = st.session_state.data
    
    if df.empty:
        st.warning("Il database tesserati è vuoto. Carica i dati nel file CSV.")
    else:
        col1, col2 = st.columns(2)
        
        with col1:
            giocatori = df[df['Tipo'] == 'Giocatore']
            scelti_p = st.multiselect("Seleziona i Giocatori", giocatori['Nominativo'].tolist())
            
        with col2:
            staff = df[df['Tipo'] == 'Staff']
            scelti_s = st.multiselect("Seleziona lo Staff", staff['Nominativo'].tolist())

        if scelti_p:
            df_p = giocatori[giocatori['Nominativo'].isin(scelti_p)]
            df_s = staff[staff['Nominativo'].isin(scelti_s)]
            
            # Generazione del file
            try:
                file_xlsx = compila_template(df_p, df_s, info)
                st.success("Distinta pronta!")
                st.download_button(
                    label="📥 Scarica Distinta Compilata",
                    data=file_xlsx,
                    file_name=f"Distinta_{info['avversario']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Errore durante la creazione del file: {e}")
