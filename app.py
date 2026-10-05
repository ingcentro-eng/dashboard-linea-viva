import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as os
from datetime import datetime

# ---------------------------------------------------------
# Configuración de Página
# ---------------------------------------------------------
st.set_page_config(
    page_title="Dashboard Línea Viva & Operaciones",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS Personalizados para acabado elegante y profesional
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .metric-card {
        background-color: #ffffff;
        padding: 18px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        border-left: 5px solid #0056b3;
        margin-bottom: 15px;
    }
    .metric-title {
        font-size: 0.85rem;
        color: #6c757d;
        font-weight: 600;
        text-transform: uppercase;
    }
    .metric-value {
        font-size: 1.8rem;
        color: #1e293b;
        font-weight: 700;
    }
    .stMultiSelect, .stSelectbox {
        font-size: 0.9rem;
    }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# Carga de Datos
# ---------------------------------------------------------
@st.cache_data(ttl=300)
def load_data():
    try:
        df = pd.read_excel('COPIA PROGRAMADOR SEGUIMIENTO.xlsx', sheet_name='Hoja1')
        
        # Limpieza básica
        df['GRUPO'] = df['GRUPO'].fillna('SIN GRUPO').astype(str).str.strip()
        df['MES'] = df['MES'].fillna('SIN MES').astype(str).str.strip().str.upper()
        df['JEFE DE CUADRILLA'] = df['JEFE DE CUADRILLA'].fillna('NO ASIGNADO').astype(str).str.strip()
        df['Estado'] = df['Estado'].fillna('Pendiente').astype(str).str.strip()
        df['DILIGENCIADA EN SISPROING'] = df['DILIGENCIADA EN SISPROING'].fillna('SIN DILIGENCIAR').astype(str).str.strip()
        
        if 'FECHA' in df.columns:
            df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
            df['DIA'] = df['FECHA'].dt.strftime('%d/%m/%Y')
            df['DIA_NUM'] = df['FECHA'].dt.day
        else:
            df['DIA'] = 'N/A'
            df['DIA_NUM'] = 0
            
        return df
    except Exception as e:
        st.error(f"Error al cargar el archivo 'COPIA PROGRAMADOR SEGUIMIENTO.xlsx': {e}")
        return pd.DataFrame()

df_raw = load_data()

if df_raw.empty:
    st.stop()

# ---------------------------------------------------------
# Filtros en Barra Lateral
# ---------------------------------------------------------
st.sidebar.image("https://img.icons8.com/fluency/96/flash-on.png", width=60)
st.sidebar.title("Filtros de Control")

# 1. Filtro Tipo de Cuadrilla (Opción de incluir Pesadas)
tipo_cuadrilla = st.sidebar.radio(
    "Tipo de Cuadrilla a Visualizar:",
    options=["Solo Línea Viva (LINV)", "Todas las Cuadrillas (Incluye CUAD y CR)"],
    index=0,
    help="Por defecto se visualizan únicamente las cuadrillas de Línea Viva (LINV). Seleccione 'Todas' para incluir pesadas."
)

if tipo_cuadrilla == "Solo Línea Viva (LINV)":
    df_filtered = df_raw[df_raw['GRUPO'].str.startswith('LINV')].copy()
else:
    df_filtered = df_raw.copy()

# 2. Filtro por Mes
meses_disponibles = sorted(df_filtered['MES'].unique().tolist())
# Orden cronológico preferente si existen los meses
orden_meses = ["MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE"]
meses_ordenados = [m for m in orden_meses if m in meses_disponibles] + [m for m in meses_disponibles if m not in orden_meses]

sel_mes = st.sidebar.multiselect("Filtrar por Mes:", options=meses_ordenados, default=meses_ordenados)

if sel_mes:
    df_filtered = df_filtered[df_filtered['MES'].isin(sel_mes)]

# 3. Filtro por Cuadrilla
cuadrillas_disponibles = sorted(df_filtered['GRUPO'].unique().tolist())
sel_grupo = st.sidebar.multiselect("Filtrar por Cuadrilla (Grupo):", options=cuadrillas_disponibles, default=cuadrillas_disponibles)

if sel_grupo:
    df_filtered = df_filtered[df_filtered['GRUPO'].isin(sel_grupo)]

# 4. Filtro por Jefe de Cuadrilla
jefes_disponibles = sorted(df_filtered['JEFE DE CUADRILLA'].unique().tolist())
sel_jefe = st.sidebar.multiselect("Filtrar por Jefe de Cuadrilla:", options=jefes_disponibles, default=jefes_disponibles)

if sel_jefe:
    df_filtered = df_filtered[df_filtered['JEFE DE CUADRILLA'].isin(sel_jefe)]

# ---------------------------------------------------------
# Encabezado Principal
# ---------------------------------------------------------
st.title("⚡ Dashboard de Seguimiento de Actividades — Línea Viva")
st.markdown("Plataforma interactiva para el control operacional, estado de solicitudes, reprogramaciones y diligenciamiento en **Sisproing**.")

st.divider()

# ---------------------------------------------------------
# KPIs Principales (Tarjetas de Control)
# ---------------------------------------------------------
tot_actividades = len(df_filtered)
ejecutadas = len(df_filtered[df_filtered['Estado'].str.lower().isin(['finalizado', 'ejecutado'])])
en_ejecucion = len(df_filtered[df_filtered['Estado'].str.lower() == 'en ejecución'])
reprogramadas = len(df_filtered[df_filtered['Estado'].str.lower().isin(['re-programado', 'reprogramado'])])
diligenciadas = len(df_filtered[df_filtered['DILIGENCIADA EN SISPROING'].str.upper() == 'DILIGENCIADA'])

pct_diligenciado = (diligenciadas / tot_actividades * 100) if tot_actividades > 0 else 0

c1, c2, c3, c4, c5 = st.columns(5)

with c1:
    st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Total Actividades</div>
            <div class="metric-value">{tot_actividades:,}</div>
        </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
        <div class="metric-card" style="border-left-color: #28a745;">
            <div class="metric-title">Finalizadas / Ejecutadas</div>
            <div class="metric-value">{ejecutadas:,}</div>
        </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown(f"""
        <div class="metric-card" style="border-left-color: #17a2b8;">
            <div class="metric-title">En Ejecución</div>
            <div class="metric-value">{en_ejecucion:,}</div>
        </div>
    """, unsafe_allow_html=True)

with c4:
    st.markdown(f"""
        <div class="metric-card" style="border-left-color: #ffc107;">
            <div class="metric-title">Re-Programadas</div>
            <div class="metric-value">{reprogramadas:,}</div>
        </div>
    """, unsafe_allow_html=True)

with c5:
    st.markdown(f"""
        <div class="metric-card" style="border-left-color: #6f42c1;">
            <div class="metric-title">Diligenciadas SISPROING</div>
            <div class="metric-value">{diligenciadas:,} <span style="font-size: 1rem; color:#6c757d;">({pct_diligenciado:.1f}%)</span></div>
        </div>
    """, unsafe_allow_html=True)

st.write("")

# ---------------------------------------------------------
# Gráficos Interactivos (Plotly)
# ---------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📊 Visión Mensual y Diaria", "👥 Desglose por Cuadrilla / Jefe", "📋 Tabla Detallada y Exportación"])

with tab1:
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.subheader("Estado de Actividades por Mes")
        df_mes_estado = df_filtered.groupby(['MES', 'Estado']).size().reset_index(name='Cantidad')
        fig_mes = px.bar(
            df_mes_estado, 
            x='MES', 
            y='Cantidad', 
            color='Estado',
            barmode='stack',
            text='Cantidad',
            color_discrete_sequence=px.colors.qualitative.Safe,
            labels={'MES': 'Mes', 'Cantidad': 'N° de Actividades'}
        )
        fig_mes.update_layout(xaxis_title="", yaxis_title="Cantidad", legend_title="Estado")
        st.plotly_chart(fig_mes, use_container_width=True)
        
    with col_g2:
        st.subheader("Estado de Diligenciamiento en SISPROING")
        df_sis = df_filtered['DILIGENCIADA EN SISPROING'].value_counts().reset_index()
        df_sis.columns = ['Estado Sisproing', 'Cantidad']
        fig_sis = px.pie(
            df_sis, 
            names='Estado Sisproing', 
            values='Cantidad',
            hole=0.4,
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        fig_sis.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig_sis, use_container_width=True)

    st.subheader("Evolución Diaria de Actividades y Reprogramaciones")
    if 'FECHA' in df_filtered.columns and not df_filtered['FECHA'].dropna().empty:
        df_dia = df_filtered.groupby([df_filtered['FECHA'].dt.date, 'Estado']).size().reset_index(name='Cantidad')
        df_dia['FECHA'] = pd.to_datetime(df_dia['FECHA'])
        
        fig_dia = px.line(
            df_dia,
            x='FECHA',
            y='Cantidad',
            color='Estado',
            markers=True,
            labels={'FECHA': 'Fecha', 'Cantidad': 'Actividades por Día'}
        )
        fig_dia.update_xaxes(dtick="M1", tickformat="%b %d")
        st.plotly_chart(fig_dia, use_container_width=True)

with tab2:
    col_c1, col_c2 = st.columns(2)
    
    with col_c1:
        st.subheader("Actividades por Cuadrilla (Grupo)")
        df_grupo = df_filtered.groupby(['GRUPO', 'Estado']).size().reset_index(name='Cantidad')
        fig_grupo = px.bar(
            df_grupo,
            y='GRUPO',
            x='Cantidad',
            color='Estado',
            orientation='h',
            text='Cantidad',
            labels={'GRUPO': 'Cuadrilla / Grupo'}
        )
        fig_grupo.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="Cantidad")
        st.plotly_chart(fig_grupo, use_container_width=True)
        
    with col_c2:
        st.subheader("Desempeño por Jefe de Cuadrilla")
        df_jefe = df_filtered.groupby(['JEFE DE CUADRILLA', 'Estado']).size().reset_index(name='Cantidad')
        fig_jefe = px.bar(
            df_jefe,
            y='JEFE DE CUADRILLA',
            x='Cantidad',
            color='Estado',
            orientation='h',
            text='Cantidad',
            labels={'JEFE DE CUADRILLA': 'Jefe de Cuadrilla'}
        )
        fig_jefe.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="Cantidad")
        st.plotly_chart(fig_jefe, use_container_width=True)

with tab3:
    st.subheader("Registro Detallado de Programación y Seguimiento")
    st.markdown("Utilice el buscador para filtrar registros específicos o exporte la vista en formato CSV.")
    
    # Columnas a mostrar
    cols_display = ['MES', 'FECHA', 'GRUPO', 'JEFE DE CUADRILLA', 'OM', 'AVISO (VP)', 'Estado', 'DILIGENCIADA EN SISPROING', 'OBSERVACIONES']
    cols_exist = [c for c in cols_display if c in df_filtered.columns]
    
    df_show = df_filtered[cols_exist].copy()
    if 'FECHA' in df_show.columns:
        df_show['FECHA'] = df_show['FECHA'].dt.strftime('%Y-%m-%d')
        
    st.dataframe(df_show, use_container_width=True, height=400)
    
    # Botón de Descarga
    csv_data = df_show.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Exportar Datos Filtrados a CSV",
        data=csv_data,
        file_name=f"seguimiento_linea_viva_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv"
    )

# Pie de Página
st.markdown("---")
st.markdown("<p style='text-align: center; color: #888; font-size: 0.85rem;'>Sistema de Control y Seguimiento de Operaciones — Celsia / Proyectos de Ingeniería S.A.</p>", unsafe_allow_html=True)
