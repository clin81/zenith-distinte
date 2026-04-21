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
        colonne_necessarie = ["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC"]
        if df is None or df.empty:
            return pd.DataFrame(columns=colonne_necessarie)
        
        for col in colonne_necessarie:
            if col not in df.columns:
                df[col] = ""
        
        # SBLOCCO TESTO: Forza le colonne a essere stringhe per accettare lettere
        df['Ruolo'] = df['Ruolo'].astype(str).replace(['nan', 'None', ''], '')
        df['Tipo'] = df['Tipo'].astype(str).replace(['nan', 'None', ''], '')
        df['Maglia'] = pd.to_numeric(df['Maglia'], errors='coerce')
        
        return df[colonne_necessarie]
    except Exception as e:
        st.error(f"Errore caricamento: {e}")
        return pd.DataFrame(columns=["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC"])

# [Le altre funzioni salva_db, safe_write e compila_template rimangono invariate]

def salva_db(df):
    try:
        df = df.fillna("")
        conn.update(data=df)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Errore salvataggio: {e}")
        return False

# --- UI PRINCIPALE ---
st.title("⚽ Zenith Prato - Sistema Distinte")
tab_distinta, tab_database = st.tabs(["📋 Genera Distinta", "⚙️ Gestione Anagrafica"])

with tab_database:
    st.header("Anagrafica Tesserati")
    df_db = carica_db()
    
    # Configurazione colonne (se la versione di Streamlit lo permette)
    config_sicura = {}
    if hasattr(st, "column_config"):
        try:
            config_sicura = {
                "Tipo": st.column_config.SelectColumn("Tipo", options=["Giocatore", "Staff"]),
                "Maglia": st.column_config.NumberColumn("N° Maglia", format="%d"),
                "Ruolo": st.column_config.TextColumn("Ruolo Staff (Lettere)")
            }
        except:
            config_sicura = {}

    # Ordinamento Manuale
    c_ord1, _ = st.columns([2, 2])
    with c_ord1:
        sort_col = st.selectbox("Ordina tabella per:", df_db.columns, index=0)
    
    df_db = df_db.sort_values(by=sort_col)

    # --- LA MODIFICA È QUI ---
    df_editato = st.data_editor(
        df_db, 
        num_rows="dynamic", 
        width="stretch", 
        key="db_editor_v4",
        column_config=config_sicura,
        hide_index=True  # <--- Questo toglie la prima colonna con i numeri 0, 1, 2...
    )
    
    if st.button("💾 Salva modifiche"):
        if salva_db(df_editato):
            st.success("Dati salvati!")
            st.rerun()

# [Codice per tab_distinta rimane uguale a prima]
