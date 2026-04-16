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
        # Ordine desiderato: Nominativo per primo
        colonne_necessarie = ["Nominativo", "Tipo", "Maglia", "GG", "MM", "AA", "FIGC"]
        if df is None or df.empty:
            return pd.DataFrame(columns=colonne_necessarie)
        
        # Assicuriamoci che le colonne siano nell'ordine giusto
        for col in colonne_necessarie:
            if col not in df.columns:
                df[col] = ""
        return df[colonne_necessarie]
    except Exception as e:
        st.error(f"Errore caricamento: {e}")
        return pd.DataFrame(columns=["Nominativo", "Tipo", "Maglia", "GG", "MM", "AA", "FIGC"])

def salva_db(df):
    try:
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
    safe_write(ws, 'G7', f"Zenith Prato S.S.D.R.L. Vs {info['avversario']}")
    safe_write(ws, 'G8', f"Data: {info['data']} - Ora: {info['ora']}")
    safe_write(ws, 'G9', info['campo'])

    r_idx = 12 
    for _, row in players_df.iterrows():
        safe_write(ws, f'C{r_idx}', row.get('Maglia', ''))
        safe_write(ws, f'D{r_idx}', row.get('GG', ''))
        safe_write(ws, f'E{r_idx}', row.get('MM', ''))
        safe_write(ws, f'F{r_idx}', row.get('AA', ''))
        safe_write(ws, f'G{r_idx}', row.get('Nominativo', ''))
        safe_write(ws, f'I{r_idx}', row.get('FIGC', ''))
        r_idx += 1

    s_idx = 39
    for _, row in staff_df.iterrows():
        safe_write(ws, f'C{s_idx}', row.get('Maglia', '')) 
        safe_write(ws, f'G{s_idx}', row.get('Nominativo', ''))
        safe_write(ws, f'I{s_idx}', row.get('FIGC', ''))
        s_idx += 1

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

st.title("⚽ Zenith Prato - Distinta 2016")
tab_distinta, tab_database = st.tabs(["📋 Genera Distinta", "⚙️ Gestione Anagrafica"])

# --- TABELLA 2: GESTIONE DATABASE ---
with tab_database:
    st.header("Modifica o Aggiungi Tesserati")
    st.info("💡 Inserisci il nome nella prima colonna. Nella colonna 'Tipo' scrivi esattamente 'Giocatore' o 'Staff'.")
    
    df_db = carica_db()
    
    # Versione stabile senza configurazioni avanzate per evitare AttributeError
    df_editato = st.data_editor(
        df_db, 
        num_rows="dynamic", 
        use_container_width=True, 
        key="db_editor"
    )
    
    if st.button("💾 Salva modifiche su Google Sheets"):
        if salva_db(df_editato):
            st.success("Database aggiornato con successo!")
            st.rerun()

# --- TABELLA 1: GENERAZIONE DISTINTA ---
with tab_distinta:
    st.header("📝 Compilazione Gara")
    
    # Pannello superiore per i dettagli della gara
    with st.container(border=True):
        st.subheader("Dettagli dell'Incontro")
        c1, c2 = st.columns(2)
        with c1:
            avversario = st.text_input("Squadra Avversaria", value="SQUADRA OSPITE", help="Nome della squadra contro cui si gioca")
            luogo = st.text_input("Campo di gioco", value="Chiavacci", help="Nome del campo o stadio")
        with c2:
            data_gara = st.text_input("Data (es. 15/04/2026)", value="15/04/2026")
            ora_gara = st.text_input("Ora Inizio (es. 10:30)", value="10:30")

    st.divider()
    
    # Raccolta dei dati per la funzione di compilazione
    info = {
        "avversario": avversario,
        "campo": luogo,
        "data": data_gara,
        "ora": ora_gara
    }

    df_lavoro = carica_db()
    if not df_lavoro.empty:
        # Filtro per separare Giocatori e Staff
        giocatori = df_lavoro[df_lavoro['Tipo'].astype(str).str.lower() == 'giocatore']
        staff = df_lavoro[df_lavoro['Tipo'].astype(str).str.lower() == 'staff']

        st.subheader("Selezione Tesserati")
        col1, col2 = st.columns(2)
        with col1:
            scelti_p = st.multiselect("Seleziona Giocatori", giocatori['Nominativo'].tolist())
        with col2:
            scelti_s = st.multiselect("Seleziona Staff", staff['Nominativo'].tolist())

        # Bottone finale con i dati aggiornati
        if st.button("🚀 Genera e Scarica Distinta Excel", use_container_width=True):
            if scelti_p:
                excel_final = compila_template(
                    giocatori[giocatori['Nominativo'].isin(scelti_p)], 
                    staff[staff['Nominativo'].isin(scelti_s)], 
                    info
                )
                st.download_button(
                    label="📥 Clicca qui per il Download",
                    data=excel_final,
                    file_name=f"Distinta_{avversario}_{data_gara.replace('/', '-')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.error("Seleziona almeno un giocatore per procedere!")
