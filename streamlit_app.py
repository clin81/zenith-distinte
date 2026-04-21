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
        colonne_necessarie = ["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC", "Capitano", "Portiere"]
        
        if df is None or df.empty:
            return pd.DataFrame(columns=colonne_necessarie)
        
        for col in colonne_necessarie:
            if col not in df.columns:
                df[col] = False if col in ["Capitano", "Portiere"] else ""
        
        df['Ruolo'] = df['Ruolo'].astype(str).replace(['nan', 'None', ''], '')
        df['Maglia'] = pd.to_numeric(df['Maglia'], errors='coerce')
        df['Capitano'] = pd.to_numeric(df['Capitano'], errors='coerce').fillna(0).astype(bool)
        df['Portiere'] = pd.to_numeric(df['Portiere'], errors='coerce').fillna(0).astype(bool)
        
        return df[colonne_necessarie]
    except Exception as e:
        st.error(f"Errore caricamento: {e}")
        return pd.DataFrame(columns=colonne_necessarie)

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

def compila_template(p_df, s_df, info):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active 
    safe_write(ws, 'G7', f"Zenith Prato Vs {info['avversario']}")
    safe_write(ws, 'G8', f"Data: {info['data']} - Ora: {info['ora']}")
    safe_write(ws, 'G9', info['campo'])

    # Giocatori (Riga 12)
    for i, (_, row) in enumerate(p_df.iterrows()):
        r = 12 + i
        nome = str(row.get('Nominativo', ''))
        if row.get('Capitano'): nome += " (C)"
        if row.get('Portiere'): nome += " (P)"
        
        safe_write(ws, f'C{r}', row.get('Maglia', ''))
        safe_write(ws, f'D{r}', row.get('GG', ''))
        safe_write(ws, f'E{r}', row.get('MM', ''))
        safe_write(ws, f'F{r}', row.get('AA', ''))
        safe_write(ws, f'G{r}', nome)
        safe_write(ws, f'I{r}', row.get('FIGC', ''))

    # Staff (Riga 39)
    for i, (_, row) in enumerate(s_df.iterrows()):
        r = 39 + i
        safe_write(ws, f'C{r}', row.get('Ruolo', ''))
        safe_write(ws, f'G{r}', row.get('Nominativo', ''))
        safe_write(ws, f'I{r}', row.get('FIGC', ''))

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# --- UI ---
st.title("⚽ Zenith Prato - Sistema Distinte")
t1, t2 = st.tabs(["📋 Genera Distinta", "⚙️ Gestione Anagrafica"])

with t2:
    st.header("Anagrafica Tesserati")
    df_db = carica_db()
    
    # PROTEZIONE ANTI-CRASH: carichiamo i config solo se disponibili
    config_editor = {}
    if hasattr(st, "column_config"):
        try:
            config_editor = {
                "Tipo": st.column_config.SelectColumn("Tipo", options=["Giocatore", "Staff"]),
                "Maglia": st.column_config.NumberColumn("N°", format="%d"),
                "Capitano": st.column_config.CheckboxColumn("Capitano"),
                "Portiere": st.column_config.CheckboxColumn("Portiere")
            }
        except: config_editor = {}

    df_editato = st.data_editor(
        df_db, 
        num_rows="dynamic", 
        use_container_width=True,
        key="editor_v_final_secure",
        column_config=config_editor,
        hide_index=True 
    )
    
    if st.button("💾 Salva modifiche", use_container_width=True):
        if salva_db(df_editato):
            st.success("Dati sincronizzati!")
            st.rerun()

with t1:
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
        gioc = df_lavoro[df_lavoro['Tipo'].str.lower() == 'giocatore']
        staf = df_lavoro[df_lavoro['Tipo'].str.lower() == 'staff']
        
        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            sel_p = st.multiselect("Seleziona Giocatori", gioc['Nominativo'].tolist())
        with col2:
            sel_s = st.multiselect("Seleziona Staff", staf['Nominativo'].tolist())

        if st.button("🚀 Scarica Distinta Excel", use_container_width=True):
            if sel_p:
                xlsx = compila_template(
                    gioc[gioc['Nominativo'].isin(sel_p)], 
                    staf[staf['Nominativo'].isin(sel_s)], 
                    {"avversario": avv, "campo": cmp, "data": dat, "ora": ora}
                )
                st.download_button("📥 Scarica Ora", xlsx, f"Distinta_{avv}.xlsx", use_container_width=True)
