import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# ---------------------------------------------------------
# Configuración de Página
# ---------------------------------------------------------
st.set_page_config(
    page_title="Control Operacional — Línea Viva",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS Personalizados para acabado corporativo elegante
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .metric-card {
        background-color: #ffffff;
        padding: 16px;
        border-radius: 10px;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.05);
        border-left: 5px solid #0284c7;
        margin-bottom: 12px;
    }
    .metric-title {
        font-size: 0.8rem;
        color: #64748b;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .metric-value {
        font-size: 1.7rem;
        color: #0f172a;
        font-weight: 800;
        margin-top: 4px;
    }
    .alert-box-danger {
        background-color: #fef2f2;
        border: 1px solid #fecaca;
        border-left: 5px solid #ef4444;
        padding: 12px;
        border-radius: 8px;
        color: #991b1b;
        margin-bottom: 15px;
    }
    .alert-box-warning {
        background-color: #fffbeb;
        border: 1px solid #fef3c7;
        border-left: 5px solid #f59e0b;
        padding: 12px;
        border-radius: 8px;
        color: #92400e;
        margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# Carga y Procesamiento de Datos
# ---------------------------------------------------------
@st.cache_data(ttl=60) # Refresco rápido para detectar cambios en el Excel cargado a GitHub
def load_data():
    try:
        df = pd.read_excel('COPIA PROGRAMADOR SEGUIMIENTO.xlsx', sheet_name='Hoja1')
        
        # Limpieza básica de datos
        df['GRUPO'] = df['GRUPO'].fillna('SIN GRUPO').astype(str).str.strip()
        df['MES'] = df['MES'].fillna('SIN MES').astype(str).str.strip().str.upper()
        df['JEFE DE CUADRILLA'] = df['JEFE DE CUADRILLA'].fillna('NO ASIGNADO').astype(str).str.strip()
        df['Estado'] = df['Estado'].fillna('Pendiente').astype(str).str.strip()
        df['DILIGENCIADA EN SISPROING'] = df['DILIGENCIADA EN SISPROING'].fillna('SIN DILIGENCIAR').astype(str).str.strip()
        
        # Manejo de Fechas y Cálculo de Días Transcurridos
        if 'FECHA' in df.columns:
            df['FECHA_DT'] = pd.to_datetime(df['FECHA'], errors='coerce')
            df['FECHA_STR'] = df['FECHA_DT'].dt.strftime('%Y-%m-%d')
            
            # Cálculo de días transcurridos respecto a HOY
            hoy = pd.to_datetime(datetime.now().date())
            df['DIAS_EN_EJECUCION'] = (hoy - df['FECHA_DT']).dt.days
            df['DIAS_EN_EJECUCION'] = df['DIAS_EN_EJECUCION'].apply(lambda x: max(0, x) if pd.notnull(x) else 0)
        else:
            df['FECHA_STR'] = 'N/A'
            df['DIAS_EN_EJECUCION'] = 0
            
        return df
    except Exception as e:
        st.error(f"Error al cargar el archivo 'COPIA PROGRAMADOR SEGUIMIENTO.xlsx': {e}")
        return pd.DataFrame()

df_raw = load_data()

if df_raw.empty:
    st.stop()

# ---------------------------------------------------------
# Barra Lateral y Filtros
# ---------------------------------------------------------
st.sidebar.image("https://img.icons8.com/fluency/96/flash-on.png", width=50)
st.sidebar.title("Filtros de Control")

# 1. Filtro Tipo de Cuadrilla (Opción de exclusión pesadas)
tipo_cuadrilla = st.sidebar.radio(
    "Tipo de Cuadrilla:",
    options=["Solo Línea Viva (LINV)", "Todas las Cuadrillas (Incluye CUAD y CR)"],
    index=0
)

if tipo_cuadrilla == "Solo Línea Viva (LINV)":
    df = df_raw[df_raw['GRUPO'].str.startswith('LINV')].copy()
else:
    df = df_raw.copy()

# 2. Filtro por Mes
meses_disponibles = sorted(df['MES'].unique().tolist())
orden_meses = ["MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE"]
meses_ordenados = [m for m in orden_meses if m in meses_disponibles] + [m for m in meses_disponibles if m not in orden_meses]

sel_mes = st.sidebar.multiselect("Mes de Operación:", options=meses_ordenados, default=meses_ordenados)

if sel_mes:
    df = df[df['MES'].isin(sel_mes)]

# 3. Filtro por Cuadrilla / Grupo
cuadrillas = sorted(df['GRUPO'].unique().tolist())
sel_grupo = st.sidebar.multiselect("Cuadrilla (Grupo):", options=cuadrillas, default=cuadrillas)

if sel_grupo:
    df = df[df['GRUPO'].isin(sel_grupo)]

# 4. Filtro por Jefe de Cuadrilla
jefes = sorted(df['JEFE DE CUADRILLA'].unique().tolist())
sel_jefe = st.sidebar.multiselect("Jefe de Cuadrilla:", options=jefes, default=jefes)

if sel_jefe:
    df = df[df['JEFE DE CUADRILLA'].isin(sel_jefe)]

# ---------------------------------------------------------
# Encabezado Principal
# ---------------------------------------------------------
st.title("⚡ Control y Seguimiento Diario — Línea Viva")
st.caption(f"Última sincronización con datos de GitHub: **{datetime.now().strftime('%Y-%m-%d %H:%M')}**")

st.divider()

# ---------------------------------------------------------
# KPIs de Control General
# ---------------------------------------------------------
tot_act = len(df)
ejecutadas = len(df[df['Estado'].str.lower().isin(['finalizado', 'ejecutado'])])
en_ejecucion = len(df[df['Estado'].str.lower() == 'en ejecución'])
reprogramadas = len(df[df['Estado'].str.lower().isin(['re-programado', 'reprogramado'])])
canceladas = len(df[df['Estado'].str.lower() == 'cancelado'])

# Filtro especial: Finalizados pero SIN Diligenciar
fin_sin_diligenciar = df[(df['Estado'].str.lower().isin(['finalizado', 'ejecutado'])) & (df['DILIGENCIADA EN SISPROING'].str.upper() != 'DILIGENCIADA')]

k1, k2, k3, k4, k5, k6 = st.columns(6)

with k1:
    st.markdown(f'<div class="metric-card"><div class="metric-title">Total Programado</div><div class="metric-value">{tot_act:,}</div></div>', unsafe_allow_html=True)
with k2:
    st.markdown(f'<div class="metric-card" style="border-left-color: #10b981;"><div class="metric-title">Finalizadas</div><div class="metric-value">{ejecutadas:,}</div></div>', unsafe_allow_html=True)
with k3:
    st.markdown(f'<div class="metric-card" style="border-left-color: #f59e0b;"><div class="metric-title">En Ejecución</div><div class="metric-value">{en_ejecucion:,}</div></div>', unsafe_allow_html=True)
with k4:
    st.markdown(f'<div class="metric-card" style="border-left-color: #ef4444;"><div class="metric-title">Fin. SIN Diligenciar</div><div class="metric-value">{len(fin_sin_diligenciar):,}</div></div>', unsafe_allow_html=True)
with k5:
    st.markdown(f'<div class="metric-card" style="border-left-color: #8b5cf6;"><div class="metric-title">Re-Programadas</div><div class="metric-value">{reprogramadas:,}</div></div>', unsafe_allow_html=True)
with k6:
    st.markdown(f'<div class="metric-card" style="border-left-color: #64748b;"><div class="metric-title">Canceladas</div><div class="metric-value">{canceladas:,}</div></div>', unsafe_allow_html=True)

st.write("")

# ---------------------------------------------------------
# Pestañas del Aplicativo Operativo
# ---------------------------------------------------------
tab_auditoria, tab_fines_semana, tab_evolucion, tab_tabla = st.tabs([
    "🚨 Auditoría Operativa Crítica", 
    "📅 Control Fines de Semana (Sáb/Dom)", 
    "📊 Estado por Mes y Cuadrilla", 
    "📋 Base Completa de Seguimiento"
])

# ---------------------------------------------------------
# PESTAÑA 1: Auditoría Operativa Crítica (LO SOLICITADO)
# ---------------------------------------------------------
with tab_auditoria:
    st.subheader("📌 Actividades Prioritarias que Requieren Atención")
    
    col_a1, col_a2 = st.columns(2)
    
    with col_a1:
        st.markdown('<div class="alert-box-danger"><b>🔴 1. Actividades FINALIZADAS pero SIN DILIGENCIAR en Sisproing</b><br>Obras ejecutadas en campo pendietes por registro formal.</div>', unsafe_allow_html=True)
        cols_fin = ['FECHA_STR', 'GRUPO', 'JEFE DE CUADRILLA', 'OM', 'AVISO (VP)', 'DILIGENCIADA EN SISPROING', 'OBSERVACIONES']
        cols_exist_fin = [c for c in cols_fin if c in fin_sin_diligenciar.columns]
        st.dataframe(fin_sin_diligenciar[cols_exist_fin], use_container_width=True, height=250)
        
    with col_a2:
        st.markdown('<div class="alert-box-warning"><b>🟡 2. Actividades EN EJECUCIÓN (Antigüedad al día de hoy)</b><br>Monitoreo de días transcurridos desde el inicio registrado.</div>', unsafe_allow_html=True)
        df_en_ejec = df[df['Estado'].str.lower() == 'en ejecución'].copy()
        cols_ejec = ['FECHA_STR', 'DIAS_EN_EJECUCION', 'GRUPO', 'JEFE DE CUADRILLA', 'OM', 'AVISO (VP)', 'DILIGENCIADA EN SISPROING']
        cols_exist_ejec = [c for c in cols_ejec if c in df_en_ejec.columns]
        st.dataframe(df_en_ejec[cols_exist_ejec].sort_values(by='DIAS_EN_EJECUCION', ascending=False), use_container_width=True, height=250)

    st.divider()

    col_a3, col_a4 = st.columns(2)
    
    with col_a3:
        st.subheader("🟠 3. Actividades Re-Programadas")
        df_reprog = df[df['Estado'].str.lower().isin(['re-programado', 'reprogramado'])].copy()
        cols_rep = ['FECHA_STR', 'GRUPO', 'JEFE DE CUADRILLA', 'OM', 'AVISO (VP)', 'OBSERVACIONES']
        cols_exist_rep = [c for c in cols_rep if c in df_reprog.columns]
        st.dataframe(df_reprog[cols_exist_rep], use_container_width=True, height=220)
        
    with col_a4:
        st.subheader("❌ 4. Actividades Canceladas")
        df_canc = df[df['Estado'].str.lower() == 'cancelado'].copy()
        cols_canc = ['FECHA_STR', 'GRUPO', 'JEFE DE CUADRILLA', 'OM', 'AVISO (VP)', 'OBSERVACIONES']
        cols_exist_canc = [c for c in cols_canc if c in df_canc.columns]
        st.dataframe(df_canc[cols_exist_canc], use_container_width=True, height=220)

# ---------------------------------------------------------
# PESTAÑA 2: Control de Disponibilidad Fines de Semana
# ---------------------------------------------------------
with tab_fines_semana:
    st.subheader("📅 Disponibilidad y Programación de Fines de Semana (Sábados y Domingos)")
    st.markdown("Verificación de la regla operativa de turno: **Una cuadrilla LINV permanece disponible/activa cada semana**.")
    
    if 'FECHA_DT' in df.columns:
        df_fin_semana = df[df['FECHA_DT'].dt.dayofweek >= 5].copy() # 5 = Sábado, 6 = Domingo
        df_fin_semana['DÍA_NOMBRE'] = df_fin_semana['FECHA_DT'].dt.strftime('%A (%Y-%m-%d)')
        
        # Tabla dinámica cruzada por Fecha y Cuadrilla
        pivot_fds = pd.crosstab(
            index=[df_fin_semana['FECHA_STR'], df_fin_semana['JEFE DE CUADRILLA']],
            columns=df_fin_semana['GRUPO'],
            values=df_fin_semana['Estado'],
            aggfunc=lambda x: ', '.join(x.unique())
        ).fillna('Sin Registro / Descanso')
        
        st.dataframe(pivot_fds, use_container_width=True, height=450)

# ---------------------------------------------------------
# PESTAÑA 3: Gráficos de Estado por Mes y Cuadrilla
# ---------------------------------------------------------
with tab_evolucion:
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.subheader("Estado por Mes")
        df_mes = df.groupby(['MES', 'Estado']).size().reset_index(name='Cantidad')
        fig_mes = px.bar(
            df_mes, x='MES', y='Cantidad', color='Estado', barmode='stack', text='Cantidad',
            color_discrete_sequence=px.colors.qualitative.Set2
        )
        fig_mes.update_layout(xaxis_title="", yaxis_title="N° Actividades")
        st.plotly_chart(fig_mes, use_container_width=True)
        
    with col_g2:
        st.subheader("Diligenciamiento Sisproing por Jefe")
        df_sis_jefe = df.groupby(['JEFE DE CUADRILLA', 'DILIGENCIADA EN SISPROING']).size().reset_index(name='Cantidad')
        fig_sis = px.bar(
            df_sis_jefe, y='JEFE DE CUADRILLA', x='Cantidad', color='DILIGENCIADA EN SISPROING', orientation='h',
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        fig_sis.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="Cantidad")
        st.plotly_chart(fig_sis, use_container_width=True)

# ---------------------------------------------------------
# PESTAÑA 4: Tabla Completa de Datos y Exportación
# ---------------------------------------------------------
with tab_tabla:
    st.subheader("📋 Registro Total de Seguimiento")
    
    cols_show = ['MES', 'FECHA_STR', 'GRUPO', 'JEFE DE CUADRILLA', 'OM', 'AVISO (VP)', 'Estado', 'DILIGENCIADA EN SISPROING', 'OBSERVACIONES']
    cols_exist_show = [c for c in cols_show if c in df.columns]
    
    st.dataframe(df[cols_exist_show], use_container_width=True, height=400)
    
    csv_data = df[cols_exist_show].to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Descargar Reporte Completo en CSV",
        data=csv_data,
        file_name=f"reporte_linea_viva_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )

# Pie de Página
st.markdown("---")
st.markdown("<p style='text-align: center; color: #94a3b8; font-size: 0.8rem;'>Sistema de Control Operacional — Celsia / Proyectos de Ingeniería S.A.</p>", unsafe_allow_html=True)
