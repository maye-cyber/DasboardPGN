import streamlit as st
import pandas as pd
import plotly.express as px
import os
from fpdf import FPDF
import io
import plotly.io as pio

# CONFIGURACION DE LA PAGINA
st.set_page_config(
    page_title="Monitor PGN - Dashboard",
    layout="wide"
)

# PALETA DE COLORES CORPORATIVA PGN
COLOR_AZUL = "#003366"
COLOR_AMARILLO = "#FFCC00"
COLOR_NEGRO = "#212529"
COLOR_FONDO = "#F8F9FA"

# CSS PARA ESTILO INSTITUCIONAL
st.markdown(f"""
    <style>
    .stApp {{ background-color: {COLOR_FONDO}; }}
    .main-header {{
        background-color: {COLOR_AZUL}; padding: 25px; border-radius: 10px;
        color: white; text-align: center; margin-bottom: 25px; border-bottom: 5px solid {COLOR_AMARILLO};
    }}
    .main-header h1 {{ color: white !important; font-weight: bold; margin: 0; padding: 0; }}
    .main-header p {{ color: {COLOR_AMARILLO} !important; font-size: 1.3rem; margin: 5px 0 0 0; font-weight: 500; }}
    [data-testid="stMetricValue"] {{ color: {COLOR_AZUL} !important; font-size: 1.8rem !important; }}
    [data-testid="stMetricLabel"] {{ color: {COLOR_NEGRO} !important; font-weight: bold; }}
    div[data-testid="stPopover"] button {{ border-radius: 8px; border: 1px solid {COLOR_AZUL}; color: {COLOR_AZUL}; width: 100%; }}
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="main-header">
        <h1>Monitoreo y Alarmas</h1>
        <p>Dashboard</p>
    </div>
""", unsafe_allow_html=True)

# SISTEMAS OFICIALES RECONOCIDOS
SISTEMAS_OFICIALES = sorted(["SIM", "X-ROAD", "SIRI", "PORTAL WEB", "INTRANET", "HOMINIS", "INSAP", "APPS PORTAL", "APPS INTRANET", "STRATEGOS", "SIAF", "GESTOR DOKUS", "DOKUS", "ITA", "SIGDEA PORTAL EMPLEADO", "SIGDEA SEDE ELECTRONICA", "SIGDEA", "ALFA", "APP MOVIL", "APPS EXTERNAS", "NUEVA SEDE ELECTRONICA", "IGA", "REGLA DE NEGOCIO", "SIM HOMINIS"], key=len, reverse=True)

def extraer_alarmas(valor):
    val = str(valor).strip().upper()
    if val in ['OK', 'O.K.', 'NAN', 'NONE', '', 'NA', 'NINGUNO']: return []
    for sep in [',', ';', '\n', ' Y ', ' E ', ' - ']: val = val.replace(sep, '|')
    partes = [p.strip() for p in val.split('|') if p.strip()]
    encontrados = []
    for p in partes:
        match_oficial = False
        for oficial in SISTEMAS_OFICIALES:
            if oficial in p: encontrados.append(oficial); match_oficial = True; break
        if not match_oficial and p not in ['OK', 'O.K.', 'ALERTA', 'ALARMA']:
            p_l = p.replace("SERVICIO ", "").strip()
            if len(p_l) > 1: encontrados.append(p_l)
    return list(dict.fromkeys([e.upper() for e in encontrados]))

@st.cache_data
def cargar_y_procesar_todo(archivo_path_or_buf):
    try:
        dict_hojas = pd.read_excel(archivo_path_or_buf, sheet_name=None)
        lista_final = []
        for mes_hoja, df_hoja in dict_hojas.items():
            df_hoja.columns = [c.strip() for c in df_hoja.columns]
            mes_nombre = str(mes_hoja).strip().capitalize()
            for col in ['Monitoreo fecha', 'Horario control']:
                if col in df_hoja.columns: df_hoja[col] = df_hoja[col].ffill()
            for _, row in df_hoja.iterrows():
                app_val = row.get('Aplicativo', 'OK')
                apps = extraer_alarmas(app_val)
                if not apps:
                    lista_final.append({**row.to_dict(), 'Sistemas': 'OK', 'Es_Alarma': False, 'Mes': mes_nombre})
                else:
                    for app in apps:
                        lista_final.append({**row.to_dict(), 'Sistemas': app, 'Es_Alarma': True, 'Mes': mes_nombre})
        df_final = pd.DataFrame(lista_final)
        if 'Horario control' in df_final.columns:
            def limpiar_hora(h):
                h_s = str(h).lower()
                if any(x in h_s for x in ['08:', '8 am', '8:00']): return '8 am'
                if any(x in h_s for x in ['12:', '12 pm', '12:00']): return '12 pm'
                if any(x in h_s for x in ['16:', '4 pm', '16:00']): return '4 pm'
                return None
            df_final['Horario_Normalizado'] = df_final['Horario control'].apply(limpiar_hora)
        return df_final
    except Exception as e:
        return None

def generar_pdf_completo(total_a, s_top, m_top, fig_s, fig_m, fig_h):
    pdf = FPDF()
    pdf.add_page(); pdf.set_fill_color(0, 51, 102); pdf.rect(0, 0, 210, 40, 'F')
    pdf.set_text_color(255, 255, 255); pdf.set_font("Arial", "B", 18); pdf.cell(200, 20, "REPORTE EJECUTIVO - PGN", ln=True, align="C")
    pdf.set_text_color(0, 0, 0); pdf.ln(20); pdf.set_font("Arial", "B", 14); pdf.cell(200, 10, "Indicadores Clave", ln=True)
    pdf.set_font("Arial", "", 12)
    def c(t): return str(t).encode('latin-1', 'replace').decode('latin-1')
    pdf.cell(200, 8, c(f"- Total Alarmas: {total_a}"), ln=True)
    pdf.cell(200, 8, c(f"- Sistema Top: {s_top}"), ln=True)
    pdf.cell(200, 8, c(f"- Mes Top: {m_top}"), ln=True)
    try:
        pdf.add_page(); pdf.image(io.BytesIO(pio.to_image(fig_s, format="png", width=1200, height=800)), x=15, y=30, w=180)
        pdf.add_page(); pdf.image(io.BytesIO(pio.to_image(fig_m, format="png", width=1200, height=600)), x=15, y=30, w=180); pdf.image(io.BytesIO(pio.to_image(fig_h, format="png", width=1200, height=600)), x=15, y=140, w=180)
    except: pass
    return pdf.output()

# FLUJO
ruta_excel = "c:/DSB/monitoreos 2025.xlsx"
archivo_subido = st.sidebar.file_uploader("Cargar Excel", type=["xlsx"])
data_source = archivo_subido if archivo_subido else (ruta_excel if os.path.exists(ruta_excel) else None)

if data_source:
    df = cargar_y_procesar_todo(data_source)
    if df is not None:
        c1, c2, c_pdf, c_f1, c_f2 = st.columns([1, 1, 0.6, 0.7, 0.7])
        with c_f2:
            with st.popover("Sistemas"):
                s_alarmas = sorted(df[df['Es_Alarma']]['Sistemas'].unique().tolist())
                s_sel = st.multiselect("Sistemas:", s_alarmas, default=s_alarmas)
        with c_f1:
            with st.popover("Meses"):
                ord_m = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
                m_disp = sorted(df['Mes'].unique().tolist(), key=lambda x: ord_m.index(x) if x in ord_m else 99)
                m_sel = st.multiselect("Meses:", m_disp, default=m_disp)
        
        df_f = df[df['Mes'].isin(m_sel)].copy()
        df_a = df_f[(df_f['Es_Alarma']) & (df_f['Sistemas'].isin(s_sel))].copy()
        total_a = len(df_a)
        
        if total_a > 0:
            top_s_val = df_a.groupby('Sistemas').size().idxmax(); count_s = df_a.groupby('Sistemas').size().max(); p_s = (count_s / total_a) * 100
            top_m_val = df_a.groupby('Mes').size().idxmax(); count_m = df_a.groupby('Mes').size().max(); p_m = (count_m / total_a) * 100
            
            # FIGURAS
            fig_s = px.bar(df_a.groupby('Sistemas').size().reset_index(name='C').sort_values('C', ascending=False), x='Sistemas', y='C', text_auto=True, color='Sistemas', title="Ranking de Incidencias por Sistema", color_discrete_sequence=px.colors.qualitative.Bold)
            fig_s.update_layout(showlegend=False, xaxis_tickangle=-45, template="plotly_white")
            fig_m = px.bar(df_f.groupby('Mes')['Es_Alarma'].sum().reset_index(name='C').sort_values('Mes', key=lambda x: x.map({v: i for i, v in enumerate(ord_m)})), x='Mes', y='C', text_auto=True, color='Mes', title="Incidencias por Mes", color_discrete_sequence=px.colors.sequential.YlOrRd)
            fig_m.update_layout(showlegend=False, template="plotly_white")
            fig_h = px.bar(df_a.dropna(subset=['Horario_Normalizado']).groupby('Horario_Normalizado').size().reset_index(name='C').sort_values('Horario_Normalizado', key=lambda x: x.map({'8 am': 1, '12 pm': 2, '4 pm': 3})), x='Horario_Normalizado', y='C', text_auto=True, color='Horario_Normalizado', color_discrete_map={'8 am': '#003366', '12 pm': '#FFCC00', '4 pm': '#212529'}, title="Incidencias por Horario", template="plotly_white")
            fig_h.update_layout(showlegend=False)

            with c1: st.metric("SISTEMA MAS AFECTADO", top_s_val, delta=f"{count_s} ({p_s:.1f}%)", delta_color="inverse")
            with c2: st.metric("MES MAS CRITICO", top_m_val, delta=f"{count_m} ({p_m:.1f}%)", delta_color="inverse")
            with c_pdf:
                st.download_button("Exportar PDF", data=bytes(generar_pdf_completo(total_a, top_s_val, top_m_val, fig_s, fig_m, fig_h)), file_name="Reporte.pdf", mime="application/pdf")
            
            st.markdown("---")
            cl, cr = st.columns(2)
            with cl: st.plotly_chart(fig_s, use_container_width=True)
            with cr: st.plotly_chart(fig_m, use_container_width=True)
            st.markdown("---")
            st.plotly_chart(fig_h, use_container_width=True)
        else:
            st.warning("No hay alarmas registradas para los filtros seleccionados.")
else:
    st.info("Sube un archivo Excel para comenzar.")
