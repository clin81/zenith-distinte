import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from streamlit_gsheets import GSheetsConnection

# --- CONFIGURAZIONE PAGINA ---
st.set_page_config(page_title="Zenith Prato - Gestione Distinte", layout="wide")

TEMPLATE_FILE = "distinta_vuota.xlsx"

# --- CONNESSIONE GOOGLE SHEETS ---
# Crea la connessione (le credenziali devono essere nei Secrets di Streamlit)
conn = st.connection("gsheets", type=GSheetsConnection)

def carica_db():
    try:
        # Legge i dati dal foglio Google
        return conn.read(ttl="0") 
    except Exception as e:
        st.error(f"Errore nel caricamento da Google Sheets: {e}")
        return pd.DataFrame(columns=["Tipo", "Maglia", "GG", "MM", "AA", "Nominativo", "FIGC"])

def salva_db(df):
    try:
        conn.update(data=df)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Errore nel salvataggio su Google Sheets: {e}")
        return False

# --- FUNZIONE DI SCRITTURA SICURA ---
def safe_write(ws, cell_coord, value):
    """Scrive gestendo le celle unite (Merged Cells)."""
    from openpyxl.cell.cell import MergedCell
    cell = ws[cell_coord]
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell_coord in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                return
    else:
        cell.value = value

# --- FUNZIONE COMPILAZIONE EXCEL ---
def compila_template(players_df, staff_df, info):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active 

    # 1. Intestazione G7 (Aggiunge avversario al testo esistente)
    testo_fisso = ws['G7'].value if ws['G7'].value else "Zenith Prato S.S.D.R.L. Vs "
    safe_write(ws, 'G7', f"{testo_fisso} {info['avversario']}")
    
    # 2. Data e Ora in G8
    safe_write(ws, 'G8', f"Data: {info['data']} - Ora: {info['ora']}")
    
    # 3. Campo in G9
    safe_write(ws, 'G9', info['campo'])

    # 4. Giocatori (Inizio riga 12)
    r_idx = 12 
    for _, row in players_df.iterrows():
        safe_write(ws, f'C{r_idx}', row['Maglia'])
        safe_write(ws, f'D{r_idx}', row['GG'])
        safe_write(ws, f'E{r_idx}', row['MM'])
        safe_write(ws, f'F{r_idx}', row['AA'])
        safe_write(ws, f'G{r_idx}', row['Nominativo'])
        safe_write(ws, f'I{r_idx}', row['FIGC'])
        r_idx += 1

    # 5. Staff (Inizio riga 39)
    s_idx = 39
    for _, row in staff_df.iterrows():
        safe_write(ws, f'C{s_idx}', row['Maglia']) 
        safe_write(ws, f'G{s_idx}', row['Nominativo'])
        safe_write(ws, f'I{s_idx}', row['FIGC'])
        s_idx += 1

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- INTERFACCIA UTENTE ---
st.title("⚽ Zenith Prato - Sistema Distinte")

tab_distinta, tab_database = st.tabs(["📋 Genera Distinta", "⚙️ Gestione Anagrafica"])

# --- TABELLA 2: GESTIONE DATABASE ---
with tab_database:
    st.header("Modifica o Aggiungi Tesserati")
    df_db = carica_db()
    
    # Editor interattivo per aggiungere/modificare/eliminare
    df_editato = st.data_editor(
        df_db, 
        num_rows="dynamic", 
        use_container_width=True,
        key="db_editor"
    )
    
    if st.button("💾 Salva modifiche su Google Sheets"):
        if salva_db(df_editato):
            st.success("Database aggiornato permanentemente!")

# --- TABELLA 1: GENERAZIONE DISTINTA ---
with tab_distinta:
    st.sidebar.header("Dati della Gara")
    info = {
        "avversario": st.sidebar.text_input("Squadra Avversaria", "SQUADRA OSPITE"),
        "data": st.sidebar.text_input("Data", "15/04/2026"),
        "ora": st.sidebar.text_input("Ora Inizio", "10:30"),
        "campo": st.sidebar.text_input("Nome Campo", "Chiavacci")
    }

    df_lavoro = carica_db()
    
    if df_lavoro.empty:
        st.warning("Il database è vuoto. Vai nella scheda 'Gestione Anagrafica' per inserire i dati.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            giocatori = df_lavoro[df_lavoro['Tipo'] == 'Giocatore']
            scelti_p = st.multiselect("Seleziona Giocatori", giocatori['Nominativo'].tolist())
        with col2:
            staff = df_lavoro[df_lavoro['Tipo'] == 'Staff']
            scelti_s = st.multiselect("Seleziona Staff", staff['Nominativo'].tolist())

        if st.button("🚀 Genera File Excel"):
            if not scelti_p:
                st.error("Seleziona almeno un giocatore!")
            else:
                df_p = giocatori[giocatori['Nominativo'].isin(scelti_p)]
                df_s = staff[staff['Nominativo'].isin(scelti_s)]
                
                excel_final = compila_template(df_p, df_s, info)
                
                st.download_button(
                    label="📥 Scarica Distinta Compilata",
                    data=excel_final,
                    file_name=f"Distinta_{info['avversario']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
