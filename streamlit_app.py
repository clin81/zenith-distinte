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
        # Legge i dati dal foglio Google
        df = conn.read(ttl="0")
        
        # Se il foglio è vuoto o mancano colonne, definiamo la struttura base
        colonne_necessarie = ["Tipo", "Maglia", "GG", "MM", "AA", "Nominativo", "FIGC"]
        if df is None or df.empty:
            return pd.DataFrame(columns=colonne_necessarie)
        
        # Assicuriamoci che tutte le colonne necessarie esistano (per evitare errori nel data_editor)
        for col in colonne_necessarie:
            if col not in df.columns:
                df[col] = None
        return df
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

    # 1. Intestazione G7
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
        # Lo staff solitamente non ha maglia, usiamo la colonna C per il ruolo se presente
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
    st.info("💡 Scrivi nell'ultima riga per aggiungere. Seleziona una cella per modificare.")
    
    df_db = carica_db()
    
    # Inizializziamo la configurazione come vuota
    config_colonne = {}

    # Controllo di sicurezza: se Streamlit supporta column_config, lo usiamo
    if hasattr(st, "column_config"):
        config_colonne = {
            "Tipo": st.column_config.SelectColumn(
                "Tipo",
                options=["Giocatore", "Staff"],
                required=True,
            ),
            "Maglia": st.column_config.NumberColumn("N° Maglia", format="%d"),
            "GG": st.column_config.NumberColumn("Giorno", format="%02d"),
            "MM": st.column_config.NumberColumn("Mese", format="%02d"),
            "AA": st.column_config.NumberColumn("Anno", format="%d"),
        }

    # L'editor ora non crasha più: se config_colonne è vuoto, mostrerà una tabella standard
    df_editato = st.data_editor(
        df_db, 
        num_rows="dynamic", 
        use_container_width=True,
        key="db_editor",
        column_config=config_colonne
    )
    
    if st.button("💾 Salva modifiche su Google Sheets"):
        if salva_db(df_editato):
            st.success("Database aggiornato con successo!")
            st.rerun()

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
        st.warning("⚠️ Il database è vuoto. Vai nella scheda 'Gestione Anagrafica' per inserire i primi nomi.")
    else:
        # Filtriamo i dati (Case Sensitive: Giocatore/Staff)
        giocatori = df_lavoro[df_lavoro['Tipo'] == 'Giocatore']
        staff = df_lavoro[df_lavoro['Tipo'] == 'Staff']

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Seleziona Giocatori")
            if giocatori.empty:
                st.write("Nessun giocatore in archivio.")
                scelti_p = []
            else:
                scelti_p = st.multiselect("Cerca per nome:", giocatori['Nominativo'].tolist())
        
        with col2:
            st.subheader("Seleziona Staff")
            if staff.empty:
                st.write("Nessun membro staff in archivio.")
                scelti_s = []
            else:
                scelti_s = st.multiselect("Cerca per nome:", staff['Nominativo'].tolist())

        st.divider()

        if st.button("🚀 Genera File Excel", use_container_width=True):
            if not scelti_p:
                st.error("Seleziona almeno un giocatore per la distinta!")
