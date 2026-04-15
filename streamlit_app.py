import streamlit as st
import pandas as pd
import os
from io import BytesIO
from openpyxl import load_workbook

# 1. CONFIGURAZIONE
st.set_page_config(page_title="Zenith Prato - Distinte", layout="wide")

DB_FILE = "Database_Tesserati.csv"
TEMPLATE_FILE = "distinta_vuota.xlsx"

# --- FUNZIONI DI CARICAMENTO DB (Rimangono quelle "blindate") ---
def carica_db():
    # (Codice omesso per brevità, usa quello del messaggio precedente)
    # Assicura che il DB sia sempre leggibile
    if not os.path.exists(DB_FILE):
        return pd.DataFrame(columns=["Tipo", "Maglia", "GG", "MM", "AA", "Nominativo", "FIGC"])
    try:
        return pd.read_csv(DB_FILE, dtype=str).fillna("")
    except:
        return pd.DataFrame(columns=["Tipo", "Maglia", "GG", "MM", "AA", "Nominativo", "FIGC"])

# --- NUOVA FUNZIONE PER COMPILARE IL TEMPLATE ---
def compila_template(players_df, staff_df, info):
    # Carichiamo il file modello
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active # Prende il primo foglio

    # 1. Inseriamo i dati della gara (Modifica le celle A7, A8 ecc. in base al tuo file)
    ws['B7'] = info['avversario']
    ws['B8'] = info['data']
    ws['E8'] = info['ora']
    ws['H8'] = info['campo']

    # 2. Inseriamo i Giocatori
    # Supponiamo che la tabella inizi alla riga 12
    riga_inizio_giocatori = 12
    for i, (_, row) in enumerate(players_df.iterrows()):
        current_row = riga_inizio_giocatori + i
        ws[f'B{current_row}'] = row['Maglia']
        ws[f'C{current_row}'] = row['GG']
        ws[f'D{current_row}'] = row['MM']
        ws[f'E{current_row}'] = row['AA']
        ws[f'F{current_row}'] = row['Nominativo']
        ws[f'H{current_row}'] = row['FIGC']

    # 3. Inseriamo lo Staff
    # Supponiamo che la sezione staff inizi alla riga 35
    riga_inizio_staff = 35
    for i, (_, row) in enumerate(staff_df.iterrows()):
        current_row = riga_inizio_staff + i
        ws[f'B{current_row}'] = row['Maglia'] # Ruolo (All/Dir)
        ws[f'F{current_row}'] = row['Nominativo']
        ws[f'H{current_row}'] = row['FIGC']

    # Salvataggio in memoria
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- INTERFACCIA STREAMLIT ---
st.title("⚽ Compilatore Distinte (da Template)")

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"⚠️ Attenzione: Il file '{TEMPLATE_FILE}' non è stato trovato su GitHub!")
else:
    if 'data' not in st.session_state:
        st.session_state.data = carica_db()

    # Sidebar per info gara
    st.sidebar.header("Dati Gara")
    info = {
        "avversario": st.sidebar.text_input("Avversario", "SQUADRA OSPITE"),
        "data": st.sidebar.text_input("Data", "14/04/2026"),
        "ora": st.sidebar.text_input("Ora", "10:30"),
        "campo": st.sidebar.text_input("Campo", "Chiavacci")
    }

    df = st.session_state.data
    c1, c2 = st.columns(2)

    with c1:
        giocatori = df[df['Tipo'] == 'Giocatore']
        scelti_p = st.multiselect("Seleziona Giocatori", giocatori['Nominativo'].tolist())
        df_p = giocatori[giocatori['Nominativo'].isin(scelti_p)]

    with c2:
        staff = df[df['Tipo'] == 'Staff']
        scelti_s = st.multiselect("Seleziona Staff", staff['Nominativo'].tolist())
        df_s = staff[staff['Nominativo'].isin(scelti_s)]

    if not df_p.empty:
        # Generazione file usando il template
        file_pronto = compila_template(df_p, df_s, info)
        
        st.download_button(
            label="📥 Scarica Distinta Compilata",
            data=file_pronto,
            file_name=f"Distinta_{info['avversario']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
