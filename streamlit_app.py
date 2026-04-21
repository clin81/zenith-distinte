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
        colonne_necessarie = ["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC"]
        if df is None or df.empty:
            return pd.DataFrame(columns=colonne_necessarie)
        
        for col in colonne_necessarie:
            if col not in df.columns:
                df[col] = ""
        
        df['Ruolo'] = df['Ruolo'].astype(str).replace(['nan', 'None', ''], '')
        df['Tipo'] = df['Tipo'].astype(str).replace(['nan', 'None', ''], '')
        df['Maglia'] = pd.to_numeric(df['Maglia'], errors='coerce')
        
        return df[colonne_necessarie]
    except Exception as e:
        st.error(f"Errore caricamento database: {e}")
        return pd.DataFrame(columns=["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC"])

def salva_db(df):
    try:
        df = df.fillna("")
        conn.update(data=df)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Errore durante il salvataggio: {e}")
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
    safe_write(ws, 'G7', f"Zenith Prato Vs {info['avversario']}")
    safe_write(ws, 'G8', f"Data: {info['data']} - Ora: {info['ora']}")
    safe_write(ws, 'G9', info['campo'])

    # Giocatori
    r_idx = 12 
    for _, row in players_df.iterrows():
        safe_write(ws, f'C{r_idx}', row.get('Maglia', ''))
        safe_write(ws, f'D{r_idx}', row.get('GG', ''))
        safe_write(ws, f'E{r_idx}', row.get('MM', ''))
        safe_write(ws, f'F{r_idx}', row.get('AA', ''))
        safe_write(ws, f'G{r_idx}', row.get('Nominativo', ''))
        safe_write(ws, f'I{r_idx}', row.get('FIGC', ''))
        r_idx += 1

    # Staff
    s_idx = 39
    for _, row in staff_df.iterrows():
        safe_write(ws, f'C{s_idx}', row.get('Ruolo', ''))
        safe_write(ws, f'G{s_idx}', row.get('Nominativo', ''))
        safe_write(ws, f'I{s_idx}', row.get('FIGC', ''))
        s_idx += 1

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- UI PRINCIPALE ---
st.title("⚽ Zenith Prato - Sistema Distinte")

tab_distinta, tab_database = st.tabs(["📋 Genera Distinta", "⚙️ Gestione Anagrafica"])

# --- TABELLA 1: GENERA DISTINTA ---
with tab_distinta:
    st.header("📝 Dati della Gara")
    c1, c2 = st.columns(2)
    with c1:
        avversario = st.text_input("Squadra Avversaria", "SQUADRA OSPITE")
        campo = st.text_input("Campo di gioco", "Chiavacci")
    with c2:
        data_g = st.text_input("Data", "15/04/2026")
        ora_g = st.text_input("Ora", "10:30")

    info = {"avversario": avversario, "campo": campo, "data": data_g, "ora": ora_g}
    df_lavoro = carica_db()
    
    if not df_lavoro.empty:
        giocatori_list = df_lavoro[df_lavoro['Tipo'].str.lower() == 'giocatore']
        staff_list = df_lavoro[df_lavoro['Tipo'].str.lower() == 'staff']
        
        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            scelti_p = st.multiselect("Seleziona i Giocatori", giocatori_list['Nominativo'].tolist())
        with col2:
            scelti_s = st.multiselect("Seleziona lo Staff", staff_list['Nominativo'].tolist())

        if st.button("🚀 Genera e Scarica Excel", use_container_width=True):
            if scelti_p:
                file_excel = compila_template(
                    giocatori_list[giocatori_list['Nominativo'].isin(scelti_p)], 
                    staff_list[staff_list['Nominativo'].isin(scelti_s)], 
                    info
                )
                st.download_button(
                    label="📥 Scarica File",
                    data=file_excel,
                    file_name=f"Distinta_{avversario}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.warning("Seleziona almeno un giocatore.")

# --- TABELLA 2: GESTIONE ANAGRAFICA ---
with tab_database:
    st.header("Anagrafica Tesserati")
    df_db = carica_db()
    
    config_sicura = {}
    if hasattr(st, "column_config"):
        try:
            config_sicura = {
                "Tipo": st.column_config.SelectColumn("Tipo", options=["Giocatore", "Staff"], required=True),
                "Maglia": st.column_config.NumberColumn("N°", format="%d"),
                "Ruolo": st.column_config.TextColumn("Ruolo Staff")
            }
        except:
            config_sicura = {}

    # Ordinamento manuale rapido
    sort_col = st.selectbox("Ordina per:", df_db.columns, index=0)
    df_db = df_db.sort_values(by=sort_col)

    df_editato = st.data_editor(
        df_db, 
        num_rows="dynamic", 
        use_container_width=True,
        key="db_editor_v5",
        column_config=config_sicura,
        hide_index=True  # TOGLIE LA COLONNA CON I NUMERI
    )
    
    if st.button("💾 Salva modifiche"):
        if salva_db(df_editato):
            st.success("Database aggiornato!")
            st.rerun()
