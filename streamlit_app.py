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
        # Leggiamo i dati dal foglio Google
        df = conn.read(ttl="0")
        
        if df is None or df.empty:
            # Ritorna un dataframe vuoto con le colonne necessarie se il foglio non risponde
            return pd.DataFrame(columns=["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC", "Capitano", "Portiere"])

        # --- SOLUZIONE AL KEYERROR ---
        # Pulizia radicale dei nomi delle colonne: rimuove spazi, tabulazioni e va a capo
        df.columns = [str(c).strip().replace('\n', '').replace('\r', '') for c in df.columns]
        
        # Lista delle colonne che l'app si aspetta di trovare
        colonne_richieste = ["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC", "Capitano", "Portiere"]
        
        # Se una colonna fondamentale manca, la aggiungiamo noi vuota
        for col in colonne_richieste:
            if col not in df.columns:
                df[col] = False if col in ["Capitano", "Portiere"] else ""
        
        # Pulizia dei dati nelle celle
        df['Tipo'] = df['Tipo'].astype(str).str.strip().fillna("Giocatore")
        df['Nominativo'] = df['Nominativo'].astype(str).str.strip().replace(['nan', 'None', ''], '')
        df['Ruolo'] = df['Ruolo'].astype(str).str.strip().replace(['nan', 'None', ''], '')
        df['Maglia'] = pd.to_numeric(df['Maglia'], errors='coerce')
        
        # Conversione sicura per i Booleani (Capitano/Portiere)
        for c in ["Capitano", "Portiere"]:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype(bool)
        
        # Ordina per Tipo (Staff prima, Giocatori dopo) e poi per Nome
        return df.sort_values(by=['Tipo', 'Nominativo'], ascending=[False, True])
    except Exception as e:
        st.error(f"Errore critico durante il caricamento dati: {e}")
        return pd.DataFrame()

def salva_db(df):
    try:
        conn.update(data=df)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Errore durante il salvataggio su Google Sheets: {e}")
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
    
    # Intestazione Gara
    safe_write(ws, 'G7', f"Zenith Prato Vs {info['avversario']}")
    safe_write(ws, 'G8', f"Data: {info['data']} - Ora: {info['ora']}")
    safe_write(ws, 'G9', info['campo'])

    # Ordinamento Giocatori per Numero di Maglia (Lista Piatta)
    p_df = p_df.sort_values(by='Maglia', ascending=True)

    # Scrittura Giocatori (Inizio riga 12)
    for i, (_, row) in enumerate(p_df.iterrows()):
        r = 12 + i
        if r > 38: break # Limite massimo prima dello staff
        
        nome = f"{row.get('Nominativo', '')}"
        if row.get('Capitano'): nome += " (C)"
        if row.get('Portiere'): nome += " (P)"
        
        safe_write(ws, f'C{r}', row.get('Maglia', ''))
        safe_write(ws, f'D{r}', row.get('GG', ''))
        safe_write(ws, f'E{r}', row.get('MM', ''))
        safe_write(ws, f'F{r}', row.get('AA', ''))
        safe_write(ws, f'G{r}', nome)
        safe_write(ws, f'I{r}', row.get('FIGC', ''))

    # Scrittura Staff (Inizio riga 39)
    for i, (_, row) in enumerate(s_df.iterrows()):
        r = 39 + i
        # Ruolo in C, Nominativo in D, FIGC in I
        safe_write(ws, f'C{r}', row.get('Ruolo', ''))       
        safe_write(ws, f'D{r}', row.get('Nominativo', ''))  
        safe_write(ws, f'I{r}', row.get('FIGC', ''))        

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# --- INTERFACCIA ---
st.title("⚽ Zenith Prato - Sistema Distinte")
t1, t2 = st.tabs(["📋 Genera Distinta", "⚙️ Gestione Anagrafica"])

with t2:
    st.header("Anagrafica Tesserati")
    df = carica_db()
    
    if not df.empty:
        # Filtro Giocatori sicuro (ignora maiuscole/minuscole)
        giocatori_df = df[df['Tipo'].str.contains('giocatore', case=False, na=False)]
        giocatori_nomi = giocatori_df['Nominativo'].tolist()
        
        st.subheader("🏆 Ruoli Speciali")
        c1, c2 = st.columns(2)
        with c1:
            cap_sel = df[df['Capitano'] == True]['Nominativo'].iloc[0] if not df[df['Capitano'] == True].empty else "Nessuno"
            idx_cap = (giocatori_nomi.index(cap_sel) + 1) if cap_sel in giocatori_nomi else 0
            capitano = st.selectbox("Seleziona Capitano", ["Nessuno"] + giocatori_nomi, index=idx_cap)
        with c2:
            por_sel = df[df['Portiere'] == True]['Nominativo'].iloc[0] if not df[df['Portiere'] == True].empty else "Nessuno"
            idx_por = (giocatori_nomi.index(por_sel) + 1) if por_sel in giocatori_nomi else 0
            portiere = st.selectbox("Seleziona Portiere", ["Nessuno"] + giocatori_nomi, index=idx_por)

        st.subheader("📝 Dati Anagrafici")
        col_visibili = ["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC"]
        df_edit = st.data_editor(df[col_visibili], num_rows="dynamic", use_container_width=True, hide_index=True, key="editor_v6")
        
        if st.button("💾 Salva modifiche Database", use_container_width=True):
            df_edit['Capitano'] = df_edit['Nominativo'] == capitano
            df_edit['Portiere'] = df_edit['Nominativo'] == portiere
            if salva_db(df_edit):
                st.success("Database aggiornato con successo!")
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
        gioc = df_lavoro[df_lavoro['Tipo'].str.contains('giocatore', case=False, na=False)]
        staf = df_lavoro[df_lavoro['Tipo'].str.contains('staff', case=False, na=False)]
        
        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            sel_p = st.multiselect("Seleziona Giocatori convocati", gioc['Nominativo'].tolist())
        with col2:
            sel_s = st.multiselect("Seleziona Staff presente", staf['Nominativo'].tolist())

        if st.button("🚀 Scarica Distinta Excel", use_container_width=True):
            if sel_p:
                xlsx = compila_template(
                    gioc[gioc['Nominativo'].isin(sel_p)], 
                    staf[staf['Nominativo'].isin(sel_s)], 
                    {"avversario": avv, "campo": cmp, "data": dat, "ora": ora}
                )
                st.download_button("📥 Scarica Ora", xlsx, f"Distinta_{avv}.xlsx", use_container_width=True)
            else:
                st.warning("Seleziona almeno un giocatore per generare il file.")
