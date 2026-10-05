import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os

# ---------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA Y TEMA CORPORATIVO
# ---------------------------------------------------------
st.set_page_config(
    page_title="Control Operacional Línea Viva — Celsia / PROING",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS Limpios y Ejecutivos
st.markdown("""
    <style>
    .main {
        background-color: #f8fafc;
    }
    
    /* Tarjetas KPI Corporativas */
    .kpi-card {
        background-color: #ffffff;
        padding: 16px;
        border-radius: 10px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
        border-left: 5px solid #0f172a;
        transition: all 0.2s ease-in-out;
    }
    .kpi-title {
        font-size: 0.72rem;
        color: #64748b;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .kpi-value {
        font-size: 1.7rem;
        color: #0f172a;
        font-weight: 800;
        line-height: 1.2;
        margin-top: 4px;
    }
    .kpi-subtext {
        font-size: 0.72rem;
        color: #94a3b8;
        margin-top: 4px;
    }

    /* Banners de Alerta */
    .alert-header {
        padding: 10px 14px;
        border-radius: 8px;
        font-weight: 700;
        font-size: 0.82rem;
        margin-bottom: 10px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .alert-amber { background-color: #fef3c7; border: 1px solid #fde68a; color: #92400e; }
    .alert-sky { background-color: #e0f2fe; border: 1px solid #bae6fd; color: #0369a1; }
    .alert-indigo { background-color: #e0e7ff; border: 1px solid #c7d2fe; color: #3730a3; }
    .alert-rose { background-color: #ffe4e6; border: 1px solid #fecdd3; color: #9f1239; }

    /* Tabs Estilizados */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        background-color: #ffffff;
        padding: 6px 10px;
        border-radius: 10px;
        border: 1px solid #e2e8f0;
    }
    .stTabs [data-baseweb="tab"] {
        height: 40px;
        border-radius: 6px;
        font-size: 0.82rem;
        font-weight: 600;
        color: #475569;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ff6600 !important;
        color: #ffffff !important;
    }
    </style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------
# 2. CARGA Y PROCESAMIENTO INTELIGENTE DE DATOS
# ---------------------------------------------------------
@st.cache_data(ttl=60)
def load_and_process_data():
    excel_path = 'COPIA PROGRAMADOR SEGUIMIENTO.xlsx'
    
    if not os.path.exists(excel_path):
        st.error(f"⚠️ No se encontró el archivo base '{excel_path}' en el repositorio.")
        return pd.DataFrame(), pd.DataFrame()
        
    try:
        # Carga del Excel principal
        df = pd.read_excel(excel_path, sheet_name='Hoja1')
        
        # Limpieza de campos
        df['PROCESO'] = df['PROCESO'].fillna('Sin Especificar').astype(str).str.strip()
        df['GRUPO'] = df['GRUPO'].fillna('SIN GRUPO').astype(str).str.strip()
        df['MES'] = df['MES'].fillna('SIN MES').astype(str).str.strip().str.upper()
        df['JEFE DE CUADRILLA'] = df['JEFE DE CUADRILLA'].fillna('NO ASIGNADO').astype(str).str.strip()
        df['Estado'] = df['Estado'].fillna('Pendiente').astype(str).str.strip()
        df['DILIGENCIADA EN SISPROING'] = df['DILIGENCIADA EN SISPROING'].fillna('SIN DILIGENCIAR').astype(str).str.strip()
        df['OM'] = df['OM'].fillna('').astype(str).str.strip()
        df['AVISO (VP)'] = df['AVISO (VP)'].fillna('').astype(str).str.strip()
        df['OBSERVACIONES'] = df['OBSERVACIONES'].fillna('').astype(str).str.strip()
        
        # Manejo de Fechas y Antigüedad
        if 'FECHA' in df.columns:
            df['FECHA_DT'] = pd.to_datetime(df['FECHA'], errors='coerce')
            df['FECHA_STR'] = df['FECHA_DT'].dt.strftime('%Y-%m-%d')
            
            hoy = pd.to_datetime(datetime.now().date())
            df['DIAS_EN_EJECUCION'] = (hoy - df['FECHA_DT']).dt.days
            df['DIAS_EN_EJECUCION'] = df['DIAS_EN_EJECUCION'].apply(lambda x: max(0, x) if pd.notnull(x) else 0)
            df['ES_FIN_SEMANA'] = df['FECHA_DT'].dt.dayofweek.isin([5, 6]) # Sábado (5) o Domingo (6)
            df['ES_DIA_HABIL'] = df['FECHA_DT'].dt.dayofweek.isin([0, 1, 2, 3, 4]) # Lunes a Viernes
        else:
            df['FECHA_STR'] = 'N/A'
            df['DIAS_EN_EJECUCION'] = 0
            df['ES_FIN_SEMANA'] = False
            df['ES_DIA_HABIL'] = False

        # Regla Unificada de Cancelación
        df['ES_CANCELADO'] = (
            (df['Estado'].str.lower() == 'cancelado') | 
            (df['DILIGENCIADA EN SISPROING'].str.upper().str.startswith('CANCELAR'))
        )
        
        # Carga opcional del CSV Navegador de Incidentes
        df_inc = pd.DataFrame()
        csv_files = [f for f in os.listdir('.') if f.startswith('Navegador de incidentes') and f.endswith('.csv')]
        if csv_files:
            try:
                df_inc = pd.read_csv(csv_files[0], skiprows=5, sep=';', encoding='utf-8', on_bad_lines='skip')
            except Exception:
                pass

        return df, df_inc

    except Exception as e:
        st.error(f"Error procesando los datos: {e}")
        return pd.DataFrame(), pd.DataFrame()

df_raw, df_inc = load_and_process_data()

if df_raw.empty:
    st.stop()


# ---------------------------------------------------------
# 3. BARRA LATERAL (FILTROS MULTINIVEL DE CONTROL)
# ---------------------------------------------------------
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/flash-on.png", width=44)
    st.title("Filtros de Control")
    st.caption("Sistema Operativo — Línea Viva")
    st.divider()

    # 1. Segmentación de Cuadrillas
    segmentacion = st.radio(
        "Segmentación de Cuadrillas:",
        options=["⚡ Solo Línea Viva (LINV)", "🚜 Todas (Incluye Pesadas CUAD y CR)"],
        index=0
    )

    if segmentacion.startswith("⚡"):
        df_base = df_raw[df_raw['GRUPO'].str.startswith('LINV')].copy()
    else:
        df_base = df_raw.copy()

    # 2. NUEVA SEGMENTACIÓN POR PROCESO
    procesos_disp = sorted(df_base['PROCESO'].unique().tolist())
    sel_proceso = st.multiselect("Segmentación por Proceso:", options=procesos_disp, default=procesos_disp)
    if sel_proceso:
        df_base = df_base[df_base['PROCESO'].isin(sel_proceso)]

    # 3. Filtro por Mes
    meses_ord = ["MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE"]
    meses_disp = sorted(df_base['MES'].unique().tolist())
    meses_final = [m for m in meses_ord if m in meses_disp] + [m for m in meses_disp if m not in meses_ord]
    
    sel_mes = st.multiselect("Mes de Operación:", options=meses_final, default=meses_final)
    if sel_mes:
        df_base = df_base[df_base['MES'].isin(sel_mes)]

    # 4. Filtro por Cuadrilla (Grupo)
    grupos_disp = sorted(df_base['GRUPO'].unique().tolist())
    sel_grupo = st.multiselect("Cuadrilla (Grupo):", options=grupos_disp, default=grupos_disp)
    if sel_grupo:
        df_base = df_base[df_base['GRUPO'].isin(sel_grupo)]

    # 5. Filtro por Jefe de Cuadrilla
    jefes_disp = sorted(df_base['JEFE DE CUADRILLA'].unique().tolist())
    sel_jefe = st.multiselect("Jefe de Cuadrilla:", options=jefes_disp, default=jefes_disp)
    if sel_jefe:
        df_base = df_base[df_base['JEFE DE CUADRILLA'].isin(sel_jefe)]

    # 6. Filtro Avanzado por ESTADO
    estados_disp = sorted(df_base['Estado'].unique().tolist())
    sel_estado = st.multiselect("Estado de la Actividad:", options=estados_disp, default=estados_disp)
    if sel_estado:
        df_base = df_base[df_base['Estado'].isin(sel_estado)]

    # 7. Filtro Avanzado por DILIGENCIADA EN SISPROING
    sisproing_disp = sorted(df_base['DILIGENCIADA EN SISPROING'].unique().tolist())
    sel_sisproing = st.multiselect("Diligenciada en Sisproing:", options=sisproing_disp, default=sisproing_disp)
    if sel_sisproing:
        df_base = df_base[df_base['DILIGENCIADA EN SISPROING'].isin(sel_sisproing)]

    st.divider()
    st.markdown(f"**Registros activos:** `{len(df_base):,}`")


# ---------------------------------------------------------
# 4. ENCABEZADO Y TARJETAS DE KPIS PRINCIPALES
# ---------------------------------------------------------
st.title("⚡ Control y Seguimiento Operacional — Línea Viva")
st.caption(f"Sincronización de datos al día de hoy: **{datetime.now().strftime('%d/%m/%Y %H:%M')}** | Celsia S.A. E.S.P. — PROING")

st.write("")

# Cálculos de KPIs
total_act = len(df_base)
finalizados = len(df_base[df_base['Estado'].str.lower().isin(['finalizado', 'ejecutado']) & (~df_base['ES_CANCELADO'])])
en_ejecucion = len(df_base[df_base['Estado'].str.lower() == 'en ejecución'])
fin_sin_diligenciar = len(df_base[df_base['Estado'].str.lower().isin(['finalizado', 'ejecutado']) & (df_base['DILIGENCIADA EN SISPROING'].str.upper() != 'DILIGENCIADA') & (~df_base['ES_CANCELADO'])])
reprogramados = len(df_base[df_base['Estado'].str.lower().isin(['re-programado', 'reprogramado'])])
cancelados = len(df_base[df_base['ES_CANCELADO']])

# Render Tarjetas KPI
c1, c2, c3, c4, c5, c6 = st.columns(6)

with c1:
    st.markdown(f"""
        <div class="kpi-card" style="border-left-color: #334155;">
            <div class="kpi-title">Total Actividades</div>
            <div class="kpi-value">{total_act:,}</div>
            <div class="kpi-subtext">Programadas en consulta</div>
        </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
        <div class="kpi-card" style="border-left-color: #059669;">
            <div class="kpi-title">Finalizadas</div>
            <div class="kpi-value">{finalizados:,}</div>
            <div class="kpi-subtext">{(finalizados/total_act*100 if total_act else 0):.1f}% ejecutadas</div>
        </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown(f"""
        <div class="kpi-card" style="border-left-color: #d97706;">
            <div class="kpi-title">En Ejecución</div>
            <div class="kpi-value">{en_ejecucion:,}</div>
            <div class="kpi-subtext">En desarrollo activo</div>
        </div>
    """, unsafe_allow_html=True)

with c4:
    st.markdown(f"""
        <div class="kpi-card" style="border-left-color: #0284c7;">
            <div class="kpi-title">Fin. Sin Diligenciar</div>
            <div class="kpi-value">{fin_sin_diligenciar:,}</div>
            <div class="kpi-subtext">Pendientes Sisproing</div>
        </div>
    """, unsafe_allow_html=True)

with c5:
    st.markdown(f"""
        <div class="kpi-card" style="border-left-color: #4f46e5;">
            <div class="kpi-title">Re-Programadas</div>
            <div class="kpi-value">{reprogramados:,}</div>
            <div class="kpi-subtext">N° reprogramaciones</div>
        </div>
    """, unsafe_allow_html=True)

with c6:
    st.markdown(f"""
        <div class="kpi-card" style="border-left-color: #dc2626;">
            <div class="kpi-title">Canceladas</div>
            <div class="kpi-value">{cancelados:,}</div>
            <div class="kpi-subtext">Por sistema o cliente</div>
        </div>
    """, unsafe_allow_html=True)

st.write("")


# ---------------------------------------------------------
# 5. PESTAÑAS DETALLADAS DEL APLICATIVO
# ---------------------------------------------------------
t_audit, t_sin_prog, t_fin, t_fds, t_charts, t_data = st.tabs([
    "🚨 Auditoría Operativa Crítica",
    "⚠️ Alerta: Días Hábiles Sin Programación",
    "✅ Actividades Finalizadas",
    "📅 Control Fines de Semana (Sáb/Dom)",
    "📊 Análisis por Proceso y Desempeño",
    "📋 Base Completa de Seguimiento"
])

# ---------------------------------------------------------
# PESTAÑA 1: AUDITORÍA OPERATIVA CRÍTICA
# ---------------------------------------------------------
with t_audit:
    st.markdown("### 📌 Módulo de Control de Imprevistos y Auditoría Crítica")
    st.caption("Identificación inmediata de órdenes estancadas, inconclusas o con inconsistencias de registro.")

    col_a1, col_a2 = st.columns(2)

    with col_a1:
        st.markdown('<div class="alert-header alert-amber">🟡 1. Actividades EN EJECUCIÓN (Antigüedad al día de hoy)</div>', unsafe_allow_html=True)
        df_ejec = df_base[df_base['Estado'].str.lower() == 'en ejecución'].copy()
        if not df_ejec.empty:
            cols_ej = ['PROCESO', 'FECHA_STR', 'DIAS_EN_EJECUCION', 'GRUPO', 'JEFE DE CUADRILLA', 'OM', 'AVISO (VP)', 'DILIGENCIADA EN SISPROING']
            st.dataframe(
                df_ejec[cols_ej].sort_values(by='DIAS_EN_EJECUCION', ascending=False),
                column_config={
                    "PROCESO": "Proceso",
                    "FECHA_STR": "Fecha Prog.",
                    "DIAS_EN_EJECUCION": st.column_config.NumberColumn("Días Activos", format="%d días"),
                    "GRUPO": "Cuadrilla",
                    "JEFE DE CUADRILLA": "Jefe Cuadrilla",
                    "DILIGENCIADA EN SISPROING": "Sisproing"
                },
                use_container_width=True,
                height=260,
                hide_index=True
            )
        else:
            st.info("No hay actividades registradas actualmente en estado 'En Ejecución'.")

    with col_a2:
        st.markdown('<div class="alert-header alert-sky">🔵 2. Actividades FINALIZADAS pero SIN DILIGENCIAR en Sisproing</div>', unsafe_allow_html=True)
        df_fsd = df_base[
            df_base['Estado'].str.lower().isin(['finalizado', 'ejecutado']) & 
            (df_base['DILIGENCIADA EN SISPROING'].str.upper() != 'DILIGENCIADA') &
            (~df_base['ES_CANCELADO'])
        ].copy()
        
        if not df_fsd.empty:
            cols_fsd = ['PROCESO', 'FECHA_STR', 'GRUPO', 'JEFE DE CUADRILLA', 'OM', 'AVISO (VP)', 'DILIGENCIADA EN SISPROING', 'OBSERVACIONES']
            st.dataframe(
                df_fsd[cols_fsd],
                column_config={
                    "PROCESO": "Proceso",
                    "FECHA_STR": "Fecha",
                    "GRUPO": "Cuadrilla",
                    "DILIGENCIADA EN SISPROING": "Estado Sisproing"
                },
                use_container_width=True,
                height=260,
                hide_index=True
            )
        else:
            st.success("¡Excelente! Todas las actividades finalizadas están debidamente diligenciadas en Sisproing.")

    st.divider()

    col_a3, col_a4 = st.columns(2)

    with col_a3:
        st.markdown('<div class="alert-header alert-indigo">🟣 3. Actividades Re-Programadas</div>', unsafe_allow_html=True)
        df_rep = df_base[df_base['Estado'].str.lower().isin(['re-programado', 'reprogramado'])].copy()
        if not df_rep.empty:
            cols_rep = ['PROCESO', 'FECHA_STR', 'GRUPO', 'JEFE DE CUADRILLA', 'OM', 'AVISO (VP)', 'OBSERVACIONES']
            st.dataframe(
                df_rep[cols_rep],
                column_config={"PROCESO": "Proceso", "FECHA_STR": "Fecha", "GRUPO": "Cuadrilla"},
                use_container_width=True,
                height=240,
                hide_index=True
            )
        else:
            st.info("No hay actividades reprogramadas en la selección actual.")

    with col_a4:
        st.markdown('<div class="alert-header alert-rose">🔴 4. Actividades Canceladas (Unificado Sisproing y Estado)</div>', unsafe_allow_html=True)
        df_canc = df_base[df_base['ES_CANCELADO']].copy()
        if not df_canc.empty:
            cols_canc = ['PROCESO', 'FECHA_STR', 'GRUPO', 'JEFE DE CUADRILLA', 'OM', 'AVISO (VP)', 'Estado', 'DILIGENCIADA EN SISPROING']
            st.dataframe(
                df_canc[cols_canc],
                column_config={"PROCESO": "Proceso", "FECHA_STR": "Fecha", "DILIGENCIADA EN SISPROING": "Sisproing"},
                use_container_width=True,
                height=240,
                hide_index=True
            )
        else:
            st.info("No hay actividades canceladas en el filtro aplicado.")


# ---------------------------------------------------------
# PESTAÑA 2: ALERTA DÍAS HÁBILES SIN PROGRAMACIÓN (LO SOLICITADO)
# ---------------------------------------------------------
with t_sin_prog:
    st.markdown("### ⚠️ Monitoreo de Días Hábiles Sin Programación de Cuadrillas (Lunes a Viernes)")
    st.caption("Detección de fechas laborales de Lunes a Viernes en las que una cuadrilla LINV **no registra programación** o figura en estado **'No labora'** (excluyendo Sábados, Domingos y Festivos de Colombia).")

    # Lista Oficial Festivos Colombia 2026
    festivos_colombia_2026 = [
        '2026-01-01', '2026-01-06', '2026-03-23', '2026-04-02', '2026-04-03',
        '2026-05-01', '2026-05-18', '2026-06-08', '2026-06-15', '2026-06-29',
        '2026-07-20', '2026-08-07', '2026-08-17', '2026-10-12', '2026-11-02',
        '2026-11-16', '2026-12-08', '2026-12-25'
    ]
    festivos_dt = pd.to_datetime(festivos_colombia_2026)

    if 'FECHA_DT' in df_base.columns and not df_base['FECHA_DT'].dropna().empty:
        min_date = df_base['FECHA_DT'].min()
        max_date = df_base['FECHA_DT'].max()

        rango_dias = pd.date_range(min_date, max_date)
        # Excluir Fines de semana y festivos
        dias_habiles = rango_dias[(rango_dias.dayofweek < 5) & (~rango_dias.isin(festivos_dt))]

        grupos_linv = sorted(df_base[df_base['GRUPO'].str.startswith('LINV')]['GRUPO'].unique())

        alertas_sin_prog = []

        dias_nombre_es = {'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles', 'Thursday': 'Jueves', 'Friday': 'Viernes'}

        for g in grupos_linv:
            df_g = df_base[df_base['GRUPO'] == g]
            jefe_nom = df_g['JEFE DE CUADRILLA'].iloc[0] if not df_g.empty else 'N/A'

            for d in dias_habiles:
                records_d = df_g[df_g['FECHA_DT'] == d]
                dia_str = d.strftime('%Y-%m-%d')
                nom_dia = dias_nombre_es.get(d.strftime('%A'), d.strftime('%A'))

                if records_d.empty:
                    alertas_sin_prog.append({
                        'FECHA': dia_str,
                        'DÍA': nom_dia,
                        'MES': d.strftime('%B').upper(),
                        'GRUPO': g,
                        'JEFE DE CUADRILLA': jefe_nom,
                        'ESTADO / CAUSA': '🔴 Sin registro en el programador'
                    })
                else:
                    estados_set = set(records_d['Estado'].str.lower().unique())
                    if estados_set == {'no labora'}:
                        alertas_sin_prog.append({
                            'FECHA': dia_str,
                            'DÍA': nom_dia,
                            'MES': records_d['MES'].iloc[0],
                            'GRUPO': g,
                            'JEFE DE CUADRILLA': records_d['JEFE DE CUADRILLA'].iloc[0],
                            'ESTADO / CAUSA': '🟡 Registrado formalmente como No Labora'
                        })

        df_alertas_sp = pd.DataFrame(alertas_sin_prog)

        if not df_alertas_sp.empty:
            c_sp1, c_sp2 = st.columns([3, 1])
            with c_sp1:
                st.warning(f"Se detectaron **{len(df_alertas_sp)} días hábiles** en que las cuadrillas LINV estuvieron sin programación activa.")
            with c_sp2:
                csv_sp = df_alertas_sp.to_csv(index=False).encode('utf-8')
                st.download_button("📥 Exportar Alertas CSV", data=csv_sp, file_name="alertas_dias_sin_programacion.csv", mime="text/csv")

            st.dataframe(
                df_alertas_sp,
                column_config={
                    "FECHA": "Fecha Hábil",
                    "DÍA": "Día Semana",
                    "MES": "Mes",
                    "GRUPO": "Cuadrilla LINV",
                    "JEFE DE CUADRILLA": "Jefe Responsable",
                    "ESTADO / CAUSA": "Diagnóstico de Alerta"
                },
                use_container_width=True,
                height=420,
                hide_index=True
            )
        else:
            st.success("¡Excelente! Todas las cuadrillas LINV cuentan con programación continua de Lunes a Viernes.")
    else:
        st.info("No hay suficientes datos de fechas para realizar el cálculo de días hábiles.")


# ---------------------------------------------------------
# PESTAÑA 3: ACTIVIDADES FINALIZADAS
# ---------------------------------------------------------
with t_fin:
    st.markdown("### ✅ Registro Completo de Obras Ejecutadas y Finalizadas")
    st.caption("Monitoreo de cierre formal de órdenes de trabajo y conciliación en Sisproing.")

    df_finalizadas_full = df_base[
        df_base['Estado'].str.lower().isin(['finalizado', 'ejecutado']) & 
        (~df_base['ES_CANCELADO'])
    ].copy()

    col_f1, col_f2 = st.columns([3, 1])
    with col_f1:
        st.markdown(f"**Total Ejecutadas en Filtro:** `{len(df_finalizadas_full):,}` órdenes de trabajo")
    with col_f2:
        if not df_finalizadas_full.empty:
            csv_fin = df_finalizadas_full.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Exportar Finalizadas a CSV", data=csv_fin, file_name="actividades_finalizadas.csv", mime="text/csv")

    if not df_finalizadas_full.empty:
        cols_show_fin = ['PROCESO', 'MES', 'FECHA_STR', 'GRUPO', 'JEFE DE CUADRILLA', 'OM', 'AVISO (VP)', 'DILIGENCIADA EN SISPROING', 'OBSERVACIONES']
        st.dataframe(
            df_finalizadas_full[cols_show_fin],
            column_config={
                "PROCESO": "Proceso",
                "MES": "Mes",
                "FECHA_STR": "Fecha Ejecución",
                "GRUPO": "Cuadrilla",
                "JEFE DE CUADRILLA": "Jefe de Cuadrilla",
                "DILIGENCIADA EN SISPROING": "Estado Sisproing"
            },
            use_container_width=True,
            height=450,
            hide_index=True
        )
    else:
        st.warning("No se encontraron actividades finalizadas con los criterios seleccionados.")


# ---------------------------------------------------------
# PESTAÑA 4: CONTROL DE FINES DE SEMANA (SÁB / DOM)
# ---------------------------------------------------------
with t_fds:
    st.markdown("### 📅 Matriz Operativa de Fines de Semana")
    st.caption("Verificación de la regla de negocio: **Cobertura continua alternada de cuadrillas LINV los Sábados y Domingos**.")

    df_fds = df_base[df_base['ES_FIN_SEMANA']].copy()

    if not df_fds.empty:
        df_fds['DÍA_SEMANA'] = df_fds['FECHA_DT'].dt.strftime('%A')
        dias_es = {'Saturday': 'Sábado', 'Sunday': 'Domingo'}
        df_fds['DÍA_SEMANA'] = df_fds['DÍA_SEMANA'].map(dias_es).fillna(df_fds['DÍA_SEMANA'])

        cols_fds = ['PROCESO', 'FECHA_STR', 'DÍA_SEMANA', 'MES', 'GRUPO', 'JEFE DE CUADRILLA', 'Estado', 'DILIGENCIADA EN SISPROING', 'OM', 'AVISO (VP)']
        
        st.dataframe(
            df_fds[cols_fds].sort_values(by='FECHA_STR', ascending=False),
            column_config={
                "PROCESO": "Proceso",
                "FECHA_STR": "Fecha",
                "DÍA_SEMANA": "Día",
                "GRUPO": "Cuadrilla",
                "Estado": "Estado Operativo"
            },
            use_container_width=True,
            height=450,
            hide_index=True
        )
    else:
        st.info("No existen registros programados para fines de semana en la selección actual.")


# ---------------------------------------------------------
# PESTAÑA 5: ANÁLISIS POR PROCESO Y ESTADÍSTICAS
# ---------------------------------------------------------
with t_charts:
    st.markdown("### 📊 Indicadores Gráficos de Gestión y Distribución por Proceso")

    g1, g2 = st.columns(2)

    with g1:
        st.subheader("Distribución por PROCESO y Estado")
        df_g_proc = df_base.groupby(['PROCESO', 'Estado']).size().reset_index(name='Cantidad')
        fig_proc = px.bar(
            df_g_proc, x='PROCESO', y='Cantidad', color='Estado',
            barmode='stack', text='Cantidad',
            color_discrete_sequence=px.colors.qualitative.Bold
        )
        fig_proc.update_layout(xaxis_title="Proceso", yaxis_title="N° de Actividades", legend_title="Estado")
        st.plotly_chart(fig_proc, use_container_width=True)

    with g2:
        st.subheader("Estatus Sisproing por Jefe de Cuadrilla")
        df_g_jefe = df_base.groupby(['JEFE DE CUADRILLA', 'DILIGENCIADA EN SISPROING']).size().reset_index(name='Cantidad')
        fig_jefe = px.bar(
            df_g_jefe, y='JEFE DE CUADRILLA', x='Cantidad', color='DILIGENCIADA EN SISPROING',
            orientation='h', text='Cantidad',
            color_discrete_sequence=px.colors.qualitative.Safe
        )
        fig_jefe.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="Cantidad", yaxis_title="")
        st.plotly_chart(fig_jefe, use_container_width=True)


# ---------------------------------------------------------
# PESTAÑA 6: BASE DE DATOS COMPLETA
# ---------------------------------------------------------
with t_data:
    st.markdown("### 📋 Registro Detallado y Exportación de Datos")

    cols_all = ['PROCESO', 'MES', 'FECHA_STR', 'GRUPO', 'JEFE DE CUADRILLA', 'OM', 'AVISO (VP)', 'Estado', 'DILIGENCIADA EN SISPROING', 'OBSERVACIONES']
    cols_exist_all = [c for c in cols_all if c in df_base.columns]

    st.dataframe(df_base[cols_exist_all], use_container_width=True, height=450, hide_index=True)

    csv_full = df_base[cols_exist_all].to_csv(index=False).encode('utf-8')
    st.download_button(
        "📥 Descargar Reporte Filtrado (CSV)",
        data=csv_full,
        file_name=f"reporte_operativo_linea_viva_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )

# Pie de Página Corporativo
st.markdown("---")
st.markdown("<p style='text-align: center; color: #94a3b8; font-size: 0.78rem;'>Sistema Operativo de Control y Seguimiento — Celsia S.A. E.S.P. / Proyectos de Ingeniería S.A.</p>", unsafe_allow_html=True)
