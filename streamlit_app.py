import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from streamlit_gsheets import GSheetsConnection

# --- CONFIGURAZIONE ---
st.set_page_config(page_title="Zenith Prato - Gestione Distinte", layout="wide")
TEMPLATE_FILE = "distinta_vuota.xlsx"

conn = st.connection("gsheets", type=GSheetsConnection)

def carica_db():
    try:
        # Lettura diretta
        df = conn.read(ttl="0")
        
        if df is None or df.empty:
            return pd.DataFrame(columns=["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC", "Capitano", "Portiere"])

        # Pulizia robusta delle colonne
        df.columns = [str(c).strip() for c in df.columns]
        
        # Assicuriamoci che esistano tutte le colonne
        colonne_necessarie = ["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC", "Capitano", "Portiere"]
        for col in colonne_necessarie:
            if col not in df.columns:
                df[col] = False if col in ["Capitano", "Portiere"] else ""
        
        # Formattazione per i filtri
        df['Tipo'] = df['Tipo'].astype(str).str.strip().fillna("Giocatore")
        df['Nominativo'] = df['Nominativo'].astype(str).str.strip()
        
        return df.sort_values(by=['Tipo', 'Nominativo'], ascending=[False, True])
    
    except Exception as e:
        if "429" in str(e):
            st.error("🚫 **Limite Google Superato.** Attendi 60 secondi esatti senza toccare l'app per resettare il permesso.")
        else:
            st.error(f"Errore: {e}")
        return pd.DataFrame()

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
    
    # Intestazione
    safe_write(ws, 'G7', f"Zenith Prato Vs {info['avversario']}")
    safe_write(ws, 'G8', f"Data: {info['data']} - Ora: {info['ora']}")
    safe_write(ws, 'G9', info['campo'])

    # Giocatori ordinati per maglia
    p_df = p_df.sort_values(by='Maglia', ascending=True)

    # Scrittura Giocatori dalla riga 12
    for i, (_, row) in enumerate(p_df.iterrows()):
        r = 12 + i
        if r > 38: break
        
        nome_completo = f"{row.get('Nominativo', '')}"
        if row.get('Capitano'): nome_completo += " (C)"
        if row.get('Portiere'): nome_completo += " (P)"
        
        safe_write(ws, f'C{r}', row.get('Maglia', ''))
        safe_write(ws, f'D{r}', row.get('GG', ''))
        safe_write(ws, f'E{r}', row.get('MM', ''))
        safe_write(ws, f'F{r}', row.get('AA', ''))
        safe_write(ws, f'G{r}', nome_completo)
        safe_write(ws, f'I{r}', row.get('FIGC', ''))

    # Scrittura Staff dalla riga 39
    for i, (_, row) in enumerate(s_df.iterrows()):
        r = 39 + i
        safe_write(ws, f'C{r}', row.get('Ruolo', ''))       # Colonna C
        safe_write(ws, f'D{r}', row.get('Nominativo', ''))  # Colonna D
        safe_write(ws, f'I{r}', row.get('FIGC', ''))        # Colonna I

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# --- INTERFACCIA ---
st.title("⚽ Zenith Prato - Gestione Distinte")
t1, t2 = st.tabs(["📋 Genera Distinta", "⚙️ Gestione Anagrafica"])

with t2:
    st.header("Anagrafica")
    df = carica_db()
    if not df.empty:
        # Editor dati
        col_edit = ["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC"]
        df_edit = st.data_editor(df[col_edit], num_rows="dynamic", use_container_width=True, hide_index=True)
        
        if st.button("💾 Salva modifiche"):
            if conn.update(data=df_edit):
                st.success("Salvato!")
                st.rerun()

with t1:
    st.header("Gara")
    c1, c2 = st.columns(2)
    with c1:
        avv = st.text_input("Avversario", "SQUADRA OSPITE")
        cmp = st.text_input("Campo", "Chiavacci")
    with c2:
        dat = st.text_input("Data", "15/04/2026")
        ora = st.text_input("Ora", "10:30")

    df_gara = carica_db()
    if not df_gara.empty:
        gioc = df_gara[df_gara['Tipo'].str.contains('giocatore', case=False, na=False)]
        staf = df_gara[df_gara['Tipo'].str.contains('staff', case=False, na=False)]
        
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            sel_p = st.multiselect("Seleziona Giocatori", gioc['Nominativo'].tolist())
        with col_s2:
            sel_s = st.multiselect("Seleziona Staff", staf['Nominativo'].tolist())

        if st.button("🚀 Scarica Excel"):
            if sel_p:
                xlsx = compila_template(
                    gioc[gioc['Nominativo'].isin(sel_p)], 
                    staf[staf['Nominativo'].isin(sel_s)], 
                    {"avversario": avv, "campo": cmp, "data": dat, "ora": ora}
                )
                st.download_button("📥 Clicca qui per il download", xlsx, f"Distinta_{avv}.xlsx")
