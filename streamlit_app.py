import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from streamlit_gsheets import GSheetsConnection

# --- CONFIGURAZIONE PAGINA ---
st.set_page_config(page_title="Zenith Prato - Gestione Distinte", layout="wide")
TEMPLATE_FILE = "distinta_vuota.xlsx"

conn = st.connection("gsheets", type=GSheetsConnection)

# --- FUNZIONI DI SERVIZIO ---
def carica_db():
    try:
        df = conn.read(ttl="0")
        # Aggiunte colonne Capitano e Portiere
        colonne_necessarie = ["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC", "Capitano", "Portiere"]
        
        if df is None or df.empty:
            return pd.DataFrame(columns=colonne_necessarie)
        
        for col in colonne_necessarie:
            if col not in df.columns:
                df[col] = False if col in ["Capitano", "Portiere"] else ""
        
        # Pulizia tipi dati
        df['Ruolo'] = df['Ruolo'].astype(str).replace(['nan', 'None', ''], '')
        df['Maglia'] = pd.to_numeric(df['Maglia'], errors='coerce')
        # Assicuriamoci che Capitano e Portiere siano booleani (per le checkbox)
        df['Capitano'] = df['Capitano'].fillna(False).astype(bool)
        df['Portiere'] = df['Portiere'].fillna(False).astype(bool)
        
        return df[colonne_necessarie]
    except Exception as e:
        st.error(f"Errore caricamento: {e}")
        return pd.DataFrame(columns=["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC", "Capitano", "Portiere"])

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
        for m_range in ws.merged_cells.ranges:
            if cell_coord in m_range:
                ws.cell(row=m_range.min_row, column=m_range.min_col).value = value
                return
    else:
        cell.value = value

def compila_template(players_df, staff_df, info):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active 
    
    safe_write(ws, 'G7', f"Zenith Prato Vs {info['avversario']}")
    safe_write(ws, 'G8', f"Data: {info['data']} - Ora: {info['ora']}")
    safe_write(ws, 'G9', info['campo'])

    # Scrittura Giocatori
    r_idx = 12 
    for _, row in players_df.iterrows():
        nome_completo = str(row.get('Nominativo', ''))
        # Aggiunta sigle C e P
        if row.get('Capitano'): nome_completo += " (C)"
        if row.get('Portiere'): nome_completo += " (P)"
        
        safe_write(ws, f'C{r_idx}', row.get('Maglia', ''))
        safe_write(ws, f'D{r_idx}', row.get('GG', ''))
        safe_write(ws, f'E{r_idx}', row.get('MM', ''))
        safe_write(ws, f'F{r_idx}', row.get('AA', ''))
        safe_write(ws, f'G{r_idx}', nome_completo)
        safe_write(ws, f'I{r_idx}', row.get('FIGC', ''))
        r_idx += 1

    # Scrittura Staff
    s_idx = 39
    for _, row in staff_df.iterrows():
        safe_write(ws, f'C{s_idx}', row.get('Ruolo', ''))
        safe_write(ws, f'G{s_idx}', row.get('Nominativo', ''))
        safe_write(ws, f'I{s_idx}', row.get('FIGC', ''))
        s_idx += 1

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- UI ---
st.title("⚽ Zenith Prato - Sistema Distinte")
tab1, tab2 = st.tabs(["📋 Genera Distinta", "⚙️ Gestione Anagrafica"])

with tab2:
    st.header("Anagrafica Tesserati")
    df_db = carica_db()
    
    config_editor = {
        "Tipo": st.column_config.SelectColumn("Tipo", options=["Giocatore", "Staff"], required=True),
        "Maglia": st.column_config.NumberColumn("N°", format="%d"),
        "Capitano": st.column_config.CheckboxColumn("Capitano"),
        "Portiere": st.column_config.CheckboxColumn("Portiere"),
        "Nominativo": st.column_config.TextColumn("Nome e Cognome", width="large")
    }

    df_editato = st.data_editor(
        df_db, 
        num_rows="dynamic", 
        use_container_width=True,
        key="db_editor_v11", # Nuova chiave per forzare hide_index
        column_config=config_editor,
        hide_index=True 
    )
    
    if st.button("💾 Salva modifiche", use_container_width=True):
        if salva_db(df_editato):
            st.success("Dati sincronizzati!")
            st.rerun()

with tab1:
    st.header("📝 Dati della Gara")
    c1, c2 = st.columns(2)
    with c1:
        avv = st.text_input("Squadra Avversaria", "SQUADRA OSPITE")
        cmp = st.text_input("Campo", "Chiavacci")
    with c2:
        dat = st.text_input("Data", "15/04/2026")
        ora = st.text_input("Ora", "10:30")

    df_lavoro = carica_db()
    if not df_lavoro.empty:
        giocatori = df_lavoro[df_lavoro['Tipo'].str.lower() == 'giocatore']
        staff = df_lavoro[df_lavoro['Tipo'].str.lower() == 'staff']
        
        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            scelti_p = st.multiselect("Seleziona Giocatori", giocatori['Nominativo'].tolist())
        with col2:
            scelti_s = st.multiselect("Seleziona Staff", staff['Nominativo'].tolist())

        if st.button("🚀 Scarica Distinta Excel", use_container_width=True):
            if scelti_p:
                file = compila_template(
                    giocatori[giocatori['Nominativo'].isin(scelti_p)], 
                    staff[staff['Nominativo'].isin(scelti_s)], 
                    {"avversario": avv, "campo": cmp, "data": dat, "ora": ora}
                )
                st.download_button("📥 Clicca per il Download", file, f"Distinta_{avv}.xlsx", use_container_width=True)
