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
        
        # --- PULIZIA COLONNE (Risolve il KeyError) ---
        if df is not None and not df.empty:
            # Rimuove spazi bianchi dai nomi delle colonne e li rende coerenti
            df.columns = [str(c).strip() for c in df.columns]
        
        col_db = ["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC", "Capitano", "Portiere", "Titolare"]
        
        if df is None or df.empty:
            return pd.DataFrame(columns=col_db)
            
        # Assicuriamoci che tutte le colonne esistano, altrimenti le creiamo vuote
        for col in col_db:
            if col not in df.columns:
                df[col] = False if col in ["Capitano", "Portiere", "Titolare"] else ""
        
        # Forziamo i tipi di dati per evitare errori nei selettori
        df['Tipo'] = df['Tipo'].astype(str).fillna("Giocatore")
        df['Nominativo'] = df['Nominativo'].astype(str).replace(['nan', 'None', ''], '')
        df['Ruolo'] = df['Ruolo'].astype(str).replace(['nan', 'None', ''], '')
        
        for c in ["Capitano", "Portiere", "Titolare"]:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype(bool)
        
        return df.sort_values(by=['Tipo', 'Nominativo'], ascending=[False, True])
    except Exception as e:
        st.error(f"Errore tecnico nel caricamento: {e}")
        return pd.DataFrame(columns=["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC", "Capitano", "Portiere", "Titolare"])

def salva_db(df):
    try:
        conn.update(data=df)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Errore nel salvataggio: {e}")
        return False

# ... (Le funzioni safe_write e compila_template rimangono identiche a prima)

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

    # --- LOGICA GIOCATORI ---
    # Titolari: Righe 12-22
    titolari = p_df[p_df['Titolare'] == True].head(11)
    # Riserve: Righe 23-33
    riserve = p_df[p_df['Titolare'] == False].head(11)

    def scrivi_blocco(lista, start_row):
        for i, (_, row) in enumerate(lista.iterrows()):
            r = start_row + i
            nome = f"{row.get('Nominativo', '')}"
            if row.get('Capitano'): nome += " (C)"
            if row.get('Portiere'): nome += " (P)"
            
            safe_write(ws, f'C{r}', row.get('Maglia', ''))  # N° Maglia
            safe_write(ws, f'D{r}', row.get('GG', ''))      # GG
            safe_write(ws, f'E{r}', row.get('MM', ''))      # MM
            safe_write(ws, f'F{r}', row.get('AA', ''))      # AA
            safe_write(ws, f'G{r}', nome)                   # Nominativo
            safe_write(ws, f'I{r}', row.get('FIGC', ''))    # Matricola

    # Scrittura Atleti
    scrivi_blocco(titolari, 12)
    scrivi_blocco(riserve, 23)

    # --- LOGICA STAFF (Dalla riga 39) ---
    for i, (_, row) in enumerate(s_df.iterrows()):
        r = 39 + i
        # Spostati nelle colonne corrette per lo staff
        safe_write(ws, f'C{r}', row.get('Ruolo', ''))       # Ruolo nella colonna C
        safe_write(ws, f'D{r}', row.get('Nominativo', ''))  # Nominativo nella colonna D
        safe_write(ws, f'I{r}', row.get('FIGC', ''))        # Matricola nella colonna I

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# --- UI ---
st.title("⚽ Zenith Prato - Sistema Distinte")
t1, t2 = st.tabs(["📋 Genera Distinta", "⚙️ Gestione Anagrafica"])

with t2:
    st.header("Anagrafica Tesserati")
    df = carica_db()
    
    # Protezione: se df è vuoto, mostriamo un messaggio
    if df.empty:
        st.warning("Il database è vuoto o non accessibile. Controlla il foglio Google.")
    
    st.subheader("🏆 Ruoli Speciali")
    # Usiamo str.lower() per i confronti per essere sicuri
    giocatori_df = df[df['Tipo'].str.contains('giocatore', case=False, na=False)]
    giocatori_nomi = giocatori_df['Nominativo'].tolist()
    
    c1, c2, c3 = st.columns(3)
    with c1:
        capitano = st.selectbox("Seleziona Capitano", ["Nessuno"] + giocatori_nomi, 
                               index=(giocatori_nomi.index(df[df['Capitano']]['Nominativo'].iloc[0]) + 1) if not df[df['Capitano']].empty else 0)
    with c2:
        portiere = st.selectbox("Seleziona Portiere", ["Nessuno"] + giocatori_nomi,
                               index=(giocatori_nomi.index(df[df['Portiere']]['Nominativo'].iloc[0]) + 1) if not df[df['Portiere']].empty else 0)
    with c3:
        titolari_sel = st.multiselect("Seleziona gli 11 Titolari", giocatori_nomi,
                                    default=df[df['Titolare']]['Nominativo'].tolist())

    st.subheader("📝 Dati Anagrafici")
    col_visibili = ["Nominativo", "Tipo", "Ruolo", "Maglia", "GG", "MM", "AA", "FIGC"]
    df_vista = df[col_visibili].reset_index(drop=True)
    
    df_edit = st.data_editor(
        df_vista, 
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="final_safe_editor"
    )
    
    if st.button("💾 Salva modifiche Database", use_container_width=True):
        if len(titolari_sel) > 11:
            st.error(f"Errore: Hai selezionato {len(titolari_sel)} titolari. Massimo 11.")
        else:
            df_edit['Capitano'] = df_edit['Nominativo'] == capitano
            df_edit['Portiere'] = df_edit['Nominativo'] == portiere
            df_edit['Titolare'] = df_edit['Nominativo'].isin(titolari_sel)
            if salva_db(df_edit):
                st.success("Database aggiornato!")
                st.rerun()

with t1:
    # ... (La parte tab1 rimane uguale, carica_db si occuperà di tutto)
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
            sel_p = st.multiselect("Seleziona Convocati (Titolari + Riserve)", gioc['Nominativo'].tolist())
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
