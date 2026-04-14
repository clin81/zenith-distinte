import streamlit as st
import pandas as pd
import os
from io import BytesIO
import xlsxwriter

# 1. CONFIGURAZIONE
st.set_page_config(page_title="Zenith Prato - Distinte", layout="wide")

DB_FILE = "Database_Tesserati.csv"

# Dati di backup pronti all'uso
DATA_INIZIALE = [
    ["Giocatore", "7", "26", "05", "2016", "BARDAZZI CESARE", "4157212"],
    ["Giocatore", "21", "05", "11", "2016", "BODDI EDOARDO", "3757322"],
    ["Giocatore", "20", "24", "08", "2016", "BONINI TOMMASO", "3734578"],
    ["Giocatore", "18", "11", "04", "2016", "BURGASSI PIETRO", "4183427"],
    ["Giocatore", "6", "12", "01", "2016", "CELA MECHAN MATTEO", "4615671"],
    ["Giocatore", "15", "15", "10", "2016", "CHEREJI RAUL", "4222451"],
    ["Giocatore", "26", "24", "08", "2016", "DAI ZIJUN", "4817374"],
    ["Giocatore", "25", "25", "07", "2016", "DE SIMONE CHRISTIAN", "3639492"],
    ["Giocatore", "11", "07", "06", "2016", "DE STEFANO DENIS", "3686390"],
    ["Giocatore", "2", "16", "06", "2016", "EL AGAD MOHAMED", "4338659"],
    ["Giocatore", "13", "11", "01", "2016", "EL RHAZIRI ADAM", "4817233"],
    ["Giocatore", "14", "07", "04", "2016", "FALCONE MATTEO", "4257438"],
    ["Giocatore", "30", "15", "08", "2016", "GRASSI MATTIA", "4771538"],
    ["Giocatore", "5", "02", "02", "2016", "IDIAKE MESHACH", "3806275"],
    ["Giocatore", "19", "19", "07", "2016", "INNOCENTI ALESSANDRO", "3845616"],
    ["Giocatore", "23", "23", "09", "2016", "IORDAN ERIC STEFAN", "4817262"],
    ["Giocatore", "9", "29", "01", "2016", "MEMAJ ANDREA", "3645627"],
    ["Giocatore", "22", "27", "09", "2016", "PALATTELLA FRANCESCO", "3672869"],
    ["Giocatore", "17", "10", "17", "2016", "POCCIANTI LEONARDO", "4615632"],
    ["Giocatore", "8", "16", "01", "2016", "PRECI DEIVIS", "3974022"],
    ["Giocatore", "4", "11", "22", "2016", "SOLENNI VASCO", "4608714"],
    ["Giocatore", "29", "10", "02", "2016", "SPIRIDON ELIA", "4840330"],
    ["Giocatore", "10", "24", "08", "2016", "VELTRI ANDREA", "3734574"],
    ["Giocatore", "16", "10", "06", "2016", "XIAO KEVIN", "3750385"],
    ["Giocatore", "28", "24", "08", "2016", "ZHANG ZIXUAN", "4817381"],
    ["Staff", "All.", "", "", "", "GALEOTTI NICCOLO", "2297889"],
    ["Staff", "All.", "", "", "", "TRINGALI PAOLO", "209196881"],
    ["Staff", "Dir.", "", "", "", "BONINI EMILIANO", "210428276"],
    ["Staff", "Dir.", "", "", "", "DE SIMONE VINCENZO", "210410671"],
    ["Staff", "Dir.", "", "", "", "INNOCENTI CLAUDIO", "210428277"],
    ["Staff", "Dir.", "", "", "", "MEMAJ VLADIMIR", "210410675"]
]

def carica_db():
    # Se il file non esiste, lo creiamo
    if not os.path.exists(DB_FILE):
        df = pd.DataFrame(DATA_INIZIALE, columns=["Tipo", "Maglia", "GG", "MM", "AA", "Nominativo", "FIGC"])
        df.to_csv(DB_FILE, index=False)
        return df
    
    try:
        # Proviamo a leggere. Se è vuoto, pd.read_csv solleverà EmptyDataError
        return pd.read_csv(DB_FILE, dtype=str).fillna("")
    except Exception:
        # Se il file è vuoto o corrotto, lo resettiamo con i dati iniziali
        df = pd.DataFrame(DATA_INIZIALE, columns=["Tipo", "Maglia", "GG", "MM", "AA", "Nominativo", "FIGC"])
        df.to_csv(DB_FILE, index=False)
        return df

def genera_excel(players_df, staff_df, info):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = workbook.add_worksheet('Distinta')

    # FORMATI
    fmt_box = workbook.add_format({'border': 2, 'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    fmt_head = workbook.add_format({'border': 1, 'bold': True, 'bg_color': '#D9D9D9', 'align': 'center'})
    fmt_cell = workbook.add_format({'border': 1, 'align': 'center'})
    fmt_name = workbook.add_format({'border': 1, 'align': 'left', 'indent': 1})

    ws.set_column('F:F', 35) 
    
    ws.merge_range('A1:E5', "F.I.G.C. L.N.D.\nZENITH PRATO\nU10 PULCINI 2016", fmt_box)
    ws.merge_range('F1:J5', "STAGIONE 2025/2026\nDistinta Atleti", fmt_box)

    ws.write('A7', f"Gara: ZENITH PRATO vs {info['avversario']}")
    ws.write('A8', f"Data: {info['data']}  Ora: {info['ora']}  Campo: {info['campo']}")

    cols = ["Tit/Ris", "Maglia", "GG", "MM", "AA", "Nominativo", "Cap/Vice", "Matricola", "A", "E"]
    for i, c in enumerate(cols):
        ws.write(10, i, c, fmt_head)

    r = 11
    for _, row in players_df.iterrows():
        ws.write(r, 0, "", fmt_cell)
        ws.write(r, 1, row['Maglia'], fmt_cell)
        ws.write(r, 2, row['GG'], fmt_cell)
        ws.write(r, 3, row['MM'], fmt_cell)
        ws.write(r, 4, row['AA'], fmt_cell)
        ws.write(r, 5, row['Nominativo'], fmt_name)
        ws.write(r, 6, "", fmt_cell)
        ws.write(r, 7, row['FIGC'], fmt_cell)
        ws.write(r, 8, "", fmt_cell)
        ws.write(r, 9, "", fmt_cell)
        r += 1

    r += 2
    ws.write(r, 0, "Ruolo", fmt_head)
    ws.merge_range(r, 1, r, 6, "Nominativo", fmt_head)
    ws.write(r, 7, "Matricola", fmt_head)
    
    r += 1
    for _, row in staff_df.iterrows():
        ws.write(r, 0, row['Maglia'], fmt_cell)
        ws.merge_range(r, 1, r, 6, row['Nominativo'], fmt_name)
        ws.write(r, 7, row['FIGC'], fmt_cell)
        r += 1

    workbook.close()
    return output.getvalue()

# LOGICA UI
st.title("⚽ Zenith Prato - Distinte")

# Carichiamo i dati nel session_state una sola volta
if 'data' not in st.session_state:
    st.session_state.data = carica_db()

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
    file_ex = genera_excel(df_p, df_s, info)
    st.download_button("📥 Scarica Distinta Excel", file_ex, f"Distinta_{info['avversario']}.xlsx")
else:
    st.info("Scegli i giocatori per generare il file.")
