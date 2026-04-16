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
        
        for col in colonne_necessarie:
            if col not in df.columns:
                df[col] = ""
        
        # SBLOCCO MAGLIA: Forza il tipo numerico per evitare blocchi nell'editor
        df['Maglia'] = pd.to_numeric(df['Maglia'], errors='coerce')
        return df[colonne_necessarie]
    except Exception as e:
        st.error(f"Errore caricamento: {e}")
        return pd.DataFrame(columns=["Nominativo", "Tipo", "Maglia", "GG", "MM", "AA", "FIGC"])

def salva_db(df):
    try:
        df = df.fillna("")
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
    
    # Intestazione
    safe_write(ws, 'G7', f"Zenith Prato S.S.D.R.L. Vs {info['avversario']}")
    safe_write(ws, 'G8', f"Data: {info['data']} - Ora: {info['ora']}")
    safe_write(ws, 'G9', info['campo'])

    # Giocatori (Riga 12)
    r_idx = 12 
    for _, row in players_df.iterrows():
        safe_write(ws, f'C{r_idx}', row.get('Maglia', ''))
        safe_write(ws, f'D{r_idx}', row.get('GG', ''))
        safe_write(ws, f'E{r_idx}', row.get('MM', ''))
        safe_write(ws, f'F{r_idx}', row.get('AA', ''))
        safe_write(ws, f'G{r_idx}', row.get('Nominativo', ''))
        safe_write(ws, f'I{r_idx}', row.get('FIGC', ''))
        r_idx += 1

    # Staff (Riga 39)
    s_idx = 39
    for _, row in staff_df.iterrows():
        safe_write(ws, f'G{s_idx}', row.get('Nominativo', ''))
        safe_write(ws, f'I{s_idx}', row.get('FIGC', ''))
        s_idx += 1

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

st.title("⚽ Zenith Prato - Sistema Distinte")
tab_distinta, tab_database = st.tabs(["📋 Genera Distinta", "⚙️ Gestione Anagrafica"])

# --- TABELLA 2: GESTIONE DATABASE ---
with tab_database:
    st.header("Anagrafica Tesserati")
    df_db = carica_db()
    
    # Logica di protezione: usiamo column_config solo se supportato
    config_sicura = {}
    if hasattr(st, "column_config"):
        try:
            config_sicura = {
                "Maglia": st.column_config.NumberColumn("N° Maglia", format="%d", min_value=1),
                "Tipo": st.column_config.SelectColumn("Tipo", options=["Giocatore", "Staff"], required=True)
            }
        except:
            config_sicura = {}

    df_editato = st.data_editor(
        df_db, 
        num_rows="dynamic", 
        use_container_width=True, 
        key="db_editor",
        column_config=config_sicura
    )
    
    if st.button("💾 Salva modifiche"):
        if salva_db(df_editato):
            st.success("Database aggiornato con successo!")
            st.rerun()

# --- TABELLA 1: GENERAZIONE DISTINTA ---
with tab_distinta:
    st.header("📝 Dati della Gara")
    with st.container(border=True):
        c1, c2 = st.columns(2)
        with c1:
            avversario = st.text_input("Squadra Avversaria", "SQUADRA OSPITE")
            campo = st.text_input("Luogo/Campo", "Chiavacci")
        with c2:
            data_g = st.text_input("Data (GG/MM/AAAA)", "15/04/2026")
            ora_g = st.text_input("Ora Inizio", "10:30")

    info = {"avversario": avversario, "campo": campo, "data": data_g, "ora": ora_g}
    df_lavoro = carica_db()
    
    if not df_lavoro.empty:
        giocatori = df_lavoro[df_lavoro['Tipo'].astype(str).str.lower() == 'giocatore']
        staff = df_lavoro[df_lavoro['Tipo'].astype(str).str.lower() == 'staff']

        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            scelti_p = st.multiselect("Seleziona Giocatori", giocatori['Nominativo'].tolist())
        with col2:
            scelti_s = st.multiselect("Seleziona Staff", staff['Nominativo'].tolist())

        if st.button("🚀 Genera Distinta Excel", use_container_width=True):
            if scelti_p:
                excel_final = compila_template(
                    giocatori[giocatori['Nominativo'].isin(scelti_p)], 
                    staff[staff['Nominativo'].isin(scelti_s)], 
                    info
                )
                st.download_button("📥 Scarica File", excel_final, f"Distinta_{avversario}.xlsx", use_container_width=True)
            else:
                st.error("Seleziona almeno un giocatore!")
