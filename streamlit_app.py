import streamlit as st
import pandas as pd
import os
from io import BytesIO
import xlsxwriter

# 1. CONFIGURAZIONE E COSTANTI
st.set_page_config(page_title="Zenith Prato - Gestione Distinte", layout="wide")
DB_FILE = "Database_Tesserati.csv"

# Dati iniziali (U10 - 2016) da usare se il file non esiste
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

# 2. FUNZIONI DI SERVIZIO
def carica_o_crea_db():
    if not os.path.exists(DB_FILE):
        df = pd.DataFrame(DATA_INIZIALE, columns=["Tipo", "Maglia", "GG", "MM", "AA", "Nominativo", "FIGC"])
        df.to_csv(DB_FILE, index=False)
        return df
    return pd.read_csv(DB_FILE, dtype=str).fillna("")

def genera_excel(players_df, staff_df, info_gara):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Distinta')

    # FORMATI
    header_box = workbook.add_format({'border': 2, 'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'font_size': 11})
    table_header = workbook.add_format({'border': 1, 'bold': True, 'bg_color': '#D9D9D9', 'align': 'center', 'valign': 'vcenter', 'font_size': 9})
    cell_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 10})
    name_fmt = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter', 'font_size': 10, 'indent': 1})
    bold_txt = workbook.add_format({'bold': True, 'font_size': 10})

    # SETUP COLONNE
    col_widths = [8, 8, 4, 4, 6, 35, 10, 18, 4, 4]
    for i, w in enumerate(col_widths):
        worksheet.set_column(i, i, w)

    # 1. INTESTAZIONE
    worksheet.merge_range('A1:E5', "F.I.G.C. L.N.D.\nZENITH PRATO S.S.D.R.L.\nU10 PULCINI 2016", header_box)
    worksheet.merge_range('F1:J5', "STAGIONE 2025/2026\nDistinta Atleti", header_box)

    # 2. DATI GARA
    worksheet.write('A7', f"Gara: ZENITH PRATO vs {info_gara['avversario']}", bold_txt)
    worksheet.write('A8', f"Data: {info_gara['data']}   Ora: {info_gara['ora']}   Campo: {info_gara['campo']}", bold_txt)

    # 3. TABELLA GIOCATORI
    headers = ["Tit/Ris", "Maglia", "GG", "MM", "AA", "Nominativo", "Cap/Vice", "Matricola FIGC", "A", "E"]
    for col, text in enumerate(headers):
        worksheet.write(10, col, text, table_header)

    row_idx = 11
    for _, p in players_df.iterrows():
        worksheet.write(row_idx, 0, "", cell_fmt)
        worksheet.write(row_idx, 1, p['Maglia'], cell_fmt)
        worksheet.write(row_idx, 2, p['GG'], cell_fmt)
        worksheet.write(row_idx, 3, p['MM'], cell_fmt)
        worksheet.write(row_idx, 4, p['AA'], cell_fmt)
        worksheet.write(row_idx, 5, p['Nominativo'], name_fmt)
        worksheet.write(row_idx, 6, "", cell_fmt)
        worksheet.write(row_idx, 7, p['FIGC'], cell_fmt)
        worksheet.write(row_idx, 8, "", cell_fmt)
        worksheet.write(row_idx, 9, "", cell_fmt)
        row_idx += 1

    # 4. TABELLA STAFF
    row_idx += 2
    staff_headers = ["Ruolo", "Nominativo", "Matricola FIGC", "", "", "", "", "", "A", "E"]
    for col, text in enumerate(staff_headers):
        if text: worksheet.write(row_idx, col, text, table_header)
    
    row_idx += 1
    for _, s in staff_df.iterrows():
        worksheet.write(row_idx, 0, s['Maglia'], cell_fmt) # Uso colonna Maglia per il Ruolo
        worksheet.merge_range(row_idx, 1, row_idx, 6, s['Nominativo'], name_fmt)
        worksheet.write(row_idx, 7, s['FIGC'], cell_fmt)
        worksheet.write(row_idx, 8, "", cell_fmt)
        worksheet.write(row_idx, 9, "", cell_fmt)
        row_idx += 1

    # 5. FIRME
    row_idx += 3
    worksheet.write(row_idx, 0, "L'ARBITRO: ___________________________", bold_txt)
    worksheet.write(row_idx, 5, "IL DIRIGENTE: ___________________________", bold_txt)

    workbook.close()
    return output.getvalue()

# 3. INTERFACCIA STREAMLIT
st.title("⚽ Zenith Prato - Generatore Distinte")

df = carica_o_crea_db()

# Sidebar per info gara
st.sidebar.header("Dati Gara")
avv = st.sidebar.text_input("Avversario", "________________")
data_g = st.sidebar.text_input("Data", "DD/MM/YYYY")
ora_g = st.sidebar.text_input("Ora", "00:00")
campo_g = st.sidebar.text_input("Campo", "Chiavacci")

info_gara = {"avversario": avv, "data": data_g, "ora": ora_g, "campo": campo_g}

# Selezione Giocatori e Staff
giocatori_list = df[df['Tipo'] == 'Giocatore']
staff_list = df[df['Tipo'] == 'Staff']

col1, col2 = st.columns(2)

with col1:
    st.subheader("Convocati")
    scelti_p = st.multiselect("Seleziona i giocatori", giocatori_list['Nominativo'].tolist())
    df_scelti_p = giocatori_list[giocatori_list['Nominativo'].isin(scelti_p)]

with col2:
    st.subheader("Staff in panchina")
    scelti_s = st.multiselect("Seleziona lo staff", staff_list['Nominativo'].tolist())
    df_scelti_s = staff_list[staff_list['Nominativo'].isin(scelti_s)]

# Bottone di Download
if not df_scelti_p.empty:
    excel_data = genera_excel(df_scelti_p, df_scelti_s, info_gara)
    st.download_button(
        label="📥 Scarica Distinta Excel",
        data=excel_data,
        file_name=f"Distinta_Zenith_{avv}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.warning("Seleziona almeno un giocatore per generare il file.")
