import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from streamlit_gsheets import GSheetsConnection

# --- CONFIGURAZIONE PAGINA ---
st.set_page_config(page_title="Zenith Prato - Gestione Distinte", layout="wide")
TEMPLATE_FILE = "distinta_vuota.xlsx"

conn = st.connection("gsheets", type=GSheetsConnection)

def carica_db():
    try:
        df = conn.read(ttl="0")
        # Tutte le colonne presenti nel database fisico
        col_db = ["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC", "Capitano", "Portiere", "Titolare"]
        
        if df is None or df.empty:
            return pd.DataFrame(columns=col_db)
        
        for col in col_db:
            if col not in df.columns:
                df[col] = False if col in ["Capitano", "Portiere", "Titolare"] else ""
        
        # Pulizia dati
        df['Ruolo'] = df['Ruolo'].astype(str).replace(['nan', 'None', ''], '')
        df['Nominativo'] = df['Nominativo'].astype(str).replace(['nan', 'None', ''], '')
        for c in ["Capitano", "Portiere", "Titolare"]:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype(bool)
            
        # ORDINAMENTO: Prima Ruolo, poi Nominativo
        df = df.sort_values(by=['Ruolo', 'Nominativo'], ascending=[True, True])
        return df
    except Exception as e:
        st.error(f"Errore caricamento: {e}")
        return pd.DataFrame()

def salva_db(df_editato, df_originale):
    try:
        # Uniamo le modifiche fatte nella tabella visibile con i dati nascosti (Capitano/Portiere)
        # Usiamo il Nominativo come chiave per non perdere i dati
        df_originale.set_index('Nominativo', inplace=True)
        df_editato.set_index('Nominativo', inplace=True)
        
        # Aggiorniamo solo le colonne visibili nell'editor
        df_originale.update(df_editato)
        df_final = df_originale.reset_index()
        
        conn.update(data=df_final)
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

    # Logica Titolari (Righe 12-22) e Riserve (Righe 23-33)
    titolari = p_df[p_df['Titolare'] == True].head(11)
    riserve = p_df[p_df['Titolare'] == False].head(11)

    def scrivi_blocco(lista, start_row):
        for i, (_, row) in enumerate(lista.iterrows()):
            r = start_row + i
            nome = str(row.get('Nominativo', ''))
            if row.get('Capitano'): nome += " (C)"
            if row.get('Portiere'): nome += " (P)"
            safe_write(ws, f'C{r}', row.get('Maglia', ''))
            safe_write(ws, f'D{r}', row.get('GG', ''))
            safe_write(ws, f'E{r}', row.get('MM', ''))
            safe_write(ws, f'F{r}', row.get('AA', ''))
            safe_write(ws, f'G{r}', nome)
            safe_write(ws, f'I{r}', row.get('FIGC', ''))

    scrivi_blocco(titolari, 12)
    scrivi_blocco(riserve, 23)

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
    df_full = carica_db()
    
    # MOSTRA SOLO QUESTE COLONNE (Nasconde Capitano e Portiere)
    colonne_visibili = ["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC", "Titolare"]
    df_vista = df_full[colonne_visibili].copy()
    
    config = {
        "Tipo": st.column_config.SelectColumn("Tipo", options=["Giocatore", "Staff"]),
        "Titolare": st.column_config.CheckboxColumn("Titolare?"),
        "Maglia": st.column_config.NumberColumn("N°", format="%d")
    }

    st.info("L'indice numerico a sinistra è nascosto. Capitano e Portiere rimangono salvati ma non visualizzati qui.")

    df_editato = st.data_editor(
        df_vista, 
        num_rows="dynamic", 
        use_container_width=True,
        key="editor_v_ultimate",
        column_config=config,
        hide_index=True # Forza la rimozione della colonna 0, 1, 2...
    )
    
    if st.button("💾 Salva modifiche", use_container_width=True):
        if salva_db(df_editato, df_full):
            st.success("Dati salvati e ordinati!")
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
