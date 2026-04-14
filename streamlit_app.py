import streamlit as st
import pandas as pd
import os
from io import BytesIO
import xlsxwriter

# Configurazione Pagina
st.set_page_config(page_title="Zenith Prato 2016 - Gestore Distinte", layout="wide")

DB_FILE = "Database_Tesserati.csv"

# --- 1. FUNZIONE PER INIZIALIZZARE IL DATABASE ---
def inizializza_database_completo():
    dati = [
        # STAFF
        ['Staff', 'Allenatore', '', '', '', '', 'GALEOTTI NICCOLO', '2297889'],
        ['Staff', 'Allenatore', '', '', '', '', 'TRINGALI PAOLO', '209196881'],
        ['Staff', 'D. Accompagnatore', '', '', '', '', 'BONINI EMILIANO', '210428276'],
        ['Staff', 'D. Accompagnatore', '', '', '', '', 'DE SIMONE VINCENZO', '210410671'],
        ['Staff', 'D. Accompagnatore', '', '', '', '', 'INNOCENTI CLAUDIO', '210428277'],
        ['Staff', 'D. Accompagnatore', '', '', '', '', 'MEMAJ VLADIMIR', '210410675'],
        # CALCIATORI
        ['Giocatore', '', '7', '', '', '2016', 'BARDAZZI CESARE', '4157212'],
        ['Giocatore', '', '21', '', '', '2016', 'BODDI EDOARDO', '3757322'],
        ['Giocatore', '', '20', '', '', '2016', 'BONINI TOMMASO', '3734578'],
        ['Giocatore', '', '18', '', '', '2016', 'BURGASSI PIETRO', '4183427'],
        ['Giocatore', '', '6', '', '', '2016', 'CELA MECHAN MATTEO', '4615671'],
        ['Giocatore', '', '15', '', '', '2016', 'CHEREJI RAUL', '4222451'],
        ['Giocatore', '', '26', '', '', '2016', 'DAI ZIJUN', '4817374'],
        ['Giocatore', '', '25', '', '', '2016', 'DE SIMONE CHRISTIAN', '3639492'],
        ['Giocatore', '', '11', '', '', '2016', 'DE STEFANO DENIS', '3686390'],
        ['Giocatore', '', '2', '', '', '2016', 'EL AGAD MOHAMED', '4338659'],
        ['Giocatore', '', '13', '', '', '2016', 'EL RHAZIRI ADAM', '4817233'],
        ['Giocatore', '', '14', '', '', '2016', 'FALCONE MATTEO', '4257438'],
        ['Giocatore', '', '30', '', '', '2016', 'GRASSI MATTIA', '4771538'],
        ['Giocatore', '', '5', '', '', '2016', 'IDIAKE MESHACH', '3806275'],
        ['Giocatore', '', '19', '', '', '2016', 'INNOCENTI ALESSANDRO', '3845616'],
        ['Giocatore', '', '23', '', '', '2016', 'IORDAN ERIC STEFAN', '4817262'],
        ['Giocatore', '', '9', '', '', '2016', 'MEMAJ ANDREA', '3645627'],
        ['Giocatore', '', '', '', '', '2016', 'MOLINARO EMANUELE', '3806316'], 
        ['Giocatore', '', '22', '', '', '2016', 'PALATTELLA FRANCESCO', '3672869'],
        ['Giocatore', '', '17', '', '', '2016', 'POCCIANTI LEONARDO', '4615632'],
        ['Giocatore', '', '8', '', '', '2016', 'PRECI DEIVIS', '3974022'],
        ['Giocatore', '', '4', '', '', '2016', 'SOLENNI VASCO', '4608714'],
        ['Giocatore', '', '29', '', '', '2016', 'SPIRIDON ELIA', '4840330'],
        ['Giocatore', '', '10', '', '', '2016', 'VELTRI ANDREA', '3734574'],
        ['Giocatore', '', '16', '', '', '2016', 'XIAO KEVIN', '3750385'],
        ['Giocatore', '', '28', '', '', '2016', 'ZHANG ZIXUAN', '4817381'],
    ]
    df = pd.DataFrame(dati, columns=['Tipo', 'Ruolo', 'Maglia', 'GG', 'MM', 'AA', 'Nominativo', 'FIGC'])
    df = df.sort_values(by=['Tipo', 'Nominativo']) 
    df.to_csv(DB_FILE, index=False)
    return df

# --- 2. GESTIONE CARICAMENTO DATI ---

# Controlla se il file esiste sul server, altrimenti crealo subito
if not os.path.exists(DB_FILE):
    inizializza_database_completo()

# Ora che siamo sicuri che il file esiste, lo carichiamo in memoria
if 'data' not in st.session_state:
    try:
        st.session_state.data = pd.read_csv(DB_FILE, dtype=str).fillna("")
    except Exception as e:
        # Se per qualche motivo il caricamento fallisce, resettiamo il database
        st.error(f"Errore nel caricamento del database: {e}")
        st.session_state.data = inizializza_database_completo()

# --- 3. GENERAZIONE EXCEL ---
def genera_distinta_excel(players, staff):
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('Distinta')
    head_fmt = wb.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#D3D3D3', 'font_size': 9})
    cell_fmt = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 9})
    box_fmt = wb.add_format({'border': 2, 'valign': 'top', 'text_wrap': True})

    ws.merge_range('A1:E4', "F.I.G.C. L.N.D.\nZENITH PRATO S.S.D.R.L.\nU10 PULCINI 2016", box_fmt)
    ws.merge_range('F1:J4', "STAGIONE 2025/2026\nDistinta Atleti", box_fmt)

    headers = ["Maglia", "GG", "MM", "AA", "Nominativo", "Matricola FIGC"]
    for i, h in enumerate(headers):
        ws.write(10, [2,3,4,5,6,8][i], h, head_fmt)
    
    for r, row in enumerate(players.values.tolist()):
        ws.write(11+r, 2, row[3], cell_fmt)
        ws.write(11+r, 3, row[4], cell_fmt)
        ws.write(11+r, 4, row[5], cell_fmt)
        ws.write(11+r, 5, row[6], cell_fmt)
        ws.write(11+r, 6, row[7], cell_fmt)
        ws.write(11+r, 8, row[8], cell_fmt)

    ws.write(40, 1, "Ruolo", head_fmt)
    ws.merge_range('D41:G41', "Nominativo", head_fmt)
    ws.write(40, 7, "FIGC", head_fmt)
    
    for r, row in enumerate(staff.values.tolist()):
        ws.merge_range(41+r, 1, 41+r, 2, row[2], cell_fmt)
        ws.merge_range(41+r, 3, 41+r, 6, row[7], cell_fmt)
        ws.write(41+r, 7, row[8], cell_fmt)

    ws.set_column('G:G', 35)
    wb.close()
    return output.getvalue()

# --- 4. INTERFACCIA (Verticale) ---
st.title("⚽ Zenith Prato 2016 - Gestione Gara")

df_sel = st.session_state.data.copy()
df_sel.insert(0, 'Seleziona', False)

# 1. Tabella Giocatori (SOPRA)
st.subheader("🏃 Calciatori")
edit_p = st.data_editor(
    df_sel[df_sel['Tipo']=='Giocatore'], 
    hide_index=True, 
    use_container_width=True, 
    key="p_ed"
)

st.markdown("---") # Una linea sottile di separazione

# 2. Tabella Staff (SOTTO)
st.subheader("👨‍💼 Staff Tecnico e Dirigenziale")
edit_s = st.data_editor(
    df_sel[df_sel['Tipo']=='Staff'], 
    hide_index=True, 
    use_container_width=True, 
    key="s_ed"
)

# Area Pulsanti
st.divider()
c1, c2 = st.columns(2)
with c1:
    if st.button("💾 SALVA MODIFICHE ANAGRAFICA", use_container_width=True):
        updated = pd.concat([edit_p.drop('Seleziona', axis=1), edit_s.drop('Seleziona', axis=1)])
        st.session_state.data = updated
        st.session_state.data.to_csv(DB_FILE, index=False)
        st.success("Dati aggiornati correttamente!")

with c2:
    if st.button("🚀 GENERA DISTINTA EXCEL", use_container_width=True):
        conv_p = edit_p[edit_p['Seleziona']]
        conv_s = edit_s[edit_s['Seleziona']]
        if not conv_p.empty:
            excel_out = genera_distinta_excel(conv_p, conv_s)
            st.download_button("📥 Scarica il file per la stampa", excel_out, "Distinta_Zenith.xlsx")
        else:
            st.warning("Seleziona i convocati prima di generare il file.")

# Sidebar per Reset
if st.sidebar.button("🗑️ Reset Database"):
    if os.path.exists(DB_FILE): os.remove(DB_FILE)
    st.session_state.clear()
    st.rerun()
