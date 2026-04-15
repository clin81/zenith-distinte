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

def compila_template(players_df, staff_df, info):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active 

    # --- DATI GARA ---
    # NOTA: Se queste celle sono unite, scrivi sempre sulla PRIMA del gruppo.
    # Esempio: Se B7 e C7 sono unite, scrivi in B7.
    try:
        ws['B7'] = f"Gara: ZENITH PRATO vs {info['avversario']}"
        ws['B8'] = f"Data: {info['data']}"
        ws['E8'] = f"Ora: {info['ora']}"  # Se dà errore qui, prova a cambiare in 'D8'
        ws['G8'] = f"Campo: {info['campo']}" # Se dà errore qui, prova a cambiare in 'F8'
    except AttributeError:
        st.error("Errore nelle celle dell'intestazione (celle unite). Controlla le coordinate B7, B8, E8.")

    # --- GIOCATORI (Inizio riga 18) ---
    r_idx = 18
    for _, row in players_df.iterrows():
        # Scriviamo nelle colonne C, D, E, F, G, I
        ws[f'C{r_idx}'] = row['Maglia']
        ws[f'D{r_idx}'] = row['GG']
        ws[f'E{r_idx}'] = row['MM']
        ws[f'F{r_idx}'] = row['AA']
        ws[f'G{r_idx}'] = row['Nominativo']
        ws[f'I{r_idx}'] = row['FIGC']
        r_idx += 1

    # --- STAFF (Inizio riga 35) ---
    s_idx = 35
    for _, row in staff_df.iterrows():
        ws[f'C{s_idx}'] = row['Maglia'] # All./Dir.
        ws[f'G{s_idx}'] = row['Nominativo']
        ws[f'I{s_idx}'] = row['FIGC']
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
