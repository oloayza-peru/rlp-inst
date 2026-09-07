import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io

# ---------------------------------------------------------
# PAGE CONFIGURATION
# ---------------------------------------------------------
st.set_page_config(
    page_title="Dashboard Gerencial - Inspecciones Hidrocarburos",
    page_icon="🛢️",
    layout="wide"
)

st.title("🛢️ Sistema Integrado de Control de Inspecciones & SST")
st.caption("Plataforma Gerencial de Monitoreo de KPIs de Inspección, Seguridad y Operaciones")

# ---------------------------------------------------------
# DATA LOADING & SESSION STATE INITIALIZATION
# ---------------------------------------------------------
@st.cache_data
def load_default_data():
    try:
        df_sst = pd.read_excel("KPI PLAN ANUAL SST.xlsx", sheet_name="ANUAL SST")
    except Exception:
        df_sst = pd.DataFrame(columns=["Tipo", "Descripción", "Mes", "% Cumplimiento"])

    try:
        df_inst_plan = pd.read_excel("KPI PLAN INSTRUMENTACION.xlsx", sheet_name="Plan")
        df_inst_vaar = pd.read_excel("KPI PLAN INSTRUMENTACION.xlsx", sheet_name="Válvulas VAAR- Sensores")
        df_inst = pd.concat([df_inst_plan, df_inst_vaar], ignore_index=True)
    except Exception:
        df_inst = pd.DataFrame(columns=["PLAN", "TIPO", "MES", "UNIDAD", "TAG", "AVANCE DE CAMPO", "COMENTARIO"])

    try:
        df_acc = pd.read_excel("KPI ACCIDENTES.xlsx", sheet_name="SEG_ACCIDENTES")
    except Exception:
        df_acc = pd.DataFrame(columns=["N°", "FECHA", "HORA INCIDENTE", "UBICACIÓN", "TIPO", "CÓDIGO", "AFECTADO", "DESCRIPCIÓN", "CONSECUENCIAS", "ACCIONES INMEDIATAS", "ACCIÓN CORRECTIVA"])

    try:
        df_ops_gen = pd.read_excel("KPI OPS.xlsx", sheet_name="OPS_GENERADAS")
        df_ops_cuota = pd.read_excel("KPI OPS.xlsx", sheet_name="OPS_CUOTA")
    except Exception:
        df_ops_gen = pd.DataFrame(columns=["N°", "FECHA", "OPSID", "GRAVEDAD", "ÁREA", "UNIDAD", "EMPRESA", "ARESP", "DESCRIPCION DE OBSERVACIÓN", "OBSERVADOR"])
        df_ops_cuota = pd.DataFrame(columns=["MES", "CUOTA"])

    return df_sst, df_inst, df_acc, df_ops_gen, df_ops_cuota

if "df_sst" not in st.session_state:
    df_sst, df_inst, df_acc, df_ops_gen, df_ops_cuota = load_default_data()
    st.session_state["df_sst"] = df_sst
    st.session_state["df_inst"] = df_inst
    st.session_state["df_acc"] = df_acc
    st.session_state["df_ops_gen"] = df_ops_gen
    st.session_state["df_ops_cuota"] = df_ops_cuota

def to_excel_download(df, sheet_name="Datos"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# Month names in Spanish mapping
MONTH_NAMES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

def extract_years_and_months(df, date_col):
    if df.empty or date_col not in df.columns:
        return [], []
    dt_series = pd.to_datetime(df[date_col], errors='coerce').dropna()
    years = sorted(list(dt_series.dt.year.unique().astype(int)))
    months = sorted(list(dt_series.dt.month.unique().astype(int)))
    return years, months

def filter_by_year_month(df, date_col, selected_years, selected_months):
    if df.empty or date_col not in df.columns:
        return df
    dt_series = pd.to_datetime(df[date_col], errors='coerce')
    cond_year = dt_series.dt.year.isin(selected_years) if selected_years else True
    cond_month = dt_series.dt.month.isin(selected_months) if selected_months else True
    return df[cond_year & cond_month]

# ---------------------------------------------------------
# MAIN TABS
# ---------------------------------------------------------
tab_resumen, tab_sst, tab_inst, tab_acc, tab_ops = st.tabs([
    "📊 Resumen Ejecutivo",
    "🛡️ Plan Anual SST",
    "🔧 Instrumentación",
    "🚨 Accidentes e Incidentes",
    "👁️ Control OPS"
])

# =========================================================
# TAB: RESUMEN EJECUTIVO
# =========================================================
with tab_resumen:
    st.header("KPIs Principales del Proyecto")
    
    # Global year and month extraction
    y_sst, m_sst = extract_years_and_months(st.session_state["df_sst"], "Mes")
    y_inst, m_inst = extract_years_and_months(st.session_state["df_inst"], "MES")
    y_acc, m_acc = extract_years_and_months(st.session_state["df_acc"], "FECHA")
    y_ops, m_ops = extract_years_and_months(st.session_state["df_ops_gen"], "FECHA")
    
    all_years = sorted(list(set(y_sst + y_inst + y_acc + y_ops)))
    all_months = sorted(list(set(m_sst + m_inst + m_acc + m_ops)))
    
    col_f_y, col_f_m = st.columns(2)
    with col_f_y:
        selected_years_res = st.multiselect(
            "📅 Filtrar por Año (Resumen):",
            options=all_years,
            default=all_years,
            key="filter_res_year"
        )
    with col_f_m:
        selected_months_res = st.multiselect(
            "🗓️ Filtrar por Mes (Resumen):",
            options=all_months,
            default=all_months,
            format_func=lambda x: f"{x:02d} - {MONTH_NAMES.get(x, '')}",
            key="filter_res_month"
        )
    
    df_sst_r = filter_by_year_month(st.session_state["df_sst"], "Mes", selected_years_res, selected_months_res)
    df_inst_r = filter_by_year_month(st.session_state["df_inst"], "MES", selected_years_res, selected_months_res)
    df_acc_r = filter_by_year_month(st.session_state["df_acc"], "FECHA", selected_years_res, selected_months_res)
    df_ops_r = filter_by_year_month(st.session_state["df_ops_gen"], "FECHA", selected_years_res, selected_months_res)

    col1, col2, col3, col4 = st.columns(4)
    avg_sst = (df_sst_r["% Cumplimiento"].mean() * 100) if not df_sst_r.empty and "% Cumplimiento" in df_sst_r.columns else 0
    col1.metric("Cumplimiento SST", f"{avg_sst:.1f}%")
    
    avg_inst = (df_inst_r["AVANCE DE CAMPO"].mean() * 100) if not df_inst_r.empty and "AVANCE DE CAMPO" in df_inst_r.columns else 0
    col2.metric("Avance Instrumentación", f"{avg_inst:.1f}%")
    
    total_acc = len(df_acc_r) if not df_acc_r.empty else 0
    col3.metric("Total Incidentes/Accidentes", f"{total_acc}")
    
    total_ops = len(df_ops_r) if not df_ops_r.empty else 0
    col4.metric("OPS Generadas", f"{total_ops}")
    
    st.markdown("---")
    st.subheader("Estado de Avance General de Planes de Inspección y Seguridad")
    
    c1, c2 = st.columns(2)
    with c1:
        if not df_inst_r.empty and "TIPO" in df_inst_r.columns and "AVANCE DE CAMPO" in df_inst_r.columns:
            fig_inst_summary = px.histogram(
                df_inst_r, x="UNIDAD", y="AVANCE DE CAMPO", color="TIPO",
                histfunc="avg", barmode="group",
                title="Promedio de Avance por Unidad y Tipo de Equipo",
                labels={"AVANCE DE CAMPO": "Avance Promedio (0-1)", "UNIDAD": "Unidad Operativa"}
            )
            st.plotly_chart(fig_inst_summary, use_container_width=True)
    
    with c2:
        if not df_ops_r.empty and "GRAVEDAD" in df_ops_r.columns:
            fig_ops_pie = px.pie(
                df_ops_r, names="GRAVEDAD",
                title="Distribución de Observaciones de Seguridad (OPS) por Gravedad",
                hole=0.4
            )
            st.plotly_chart(fig_ops_pie, use_container_width=True)

# =========================================================
# TAB: PLAN ANUAL SST
# =========================================================
with tab_sst:
    st.header("Plan Anual de Seguridad y Salud en el Trabajo (SST)")
    
    uploaded_sst = st.file_uploader("Reemplazar/Actualizar Matriz SST (Excel)", type=["xlsx"], key="u_sst")
    if uploaded_sst:
        st.session_state["df_sst"] = pd.read_excel(uploaded_sst)
        st.success("Matriz SST actualizada correctamente.")
        
    df_sst = st.session_state["df_sst"]
    years_sst, months_sst = extract_years_and_months(df_sst, "Mes")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        sel_y_sst = st.multiselect("📅 Filtrar por Año (SST):", options=years_sst, default=years_sst, key="filter_sst_year")
    with col_f2:
        sel_m_sst = st.multiselect(
            "🗓️ Filtrar por Mes (SST):",
            options=months_sst,
            default=months_sst,
            format_func=lambda x: f"{x:02d} - {MONTH_NAMES.get(x, '')}",
            key="filter_sst_month"
        )
    
    df_sst_filtered = filter_by_year_month(df_sst, "Mes", sel_y_sst, sel_m_sst)

    if not df_sst_filtered.empty:
        c1, c2 = st.columns([2, 1])
        with c1:
            fig_sst_bar = px.bar(
                df_sst_filtered, x="Tipo", y="% Cumplimiento", color="Tipo",
                title="Cumplimiento Promedio por Tipo de Actividad SST",
                labels={"% Cumplimiento": "Cumplimiento (0.0 a 1.0)"}
            )
            st.plotly_chart(fig_sst_bar, use_container_width=True)
        with c2:
            avg_val = df_sst_filtered["% Cumplimiento"].mean() * 100 if "% Cumplimiento" in df_sst_filtered.columns else 0
            fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number",
                value=avg_val,
                title={'text': "Cumplimiento Global SST (%)"},
                gauge={'axis': {'range': [0, 100]},
                       'bar': {'color': "#1f77b4"},
                       'steps': [
                           {'range': [0, 70], 'color': "#ff9999"},
                           {'range': [70, 90], 'color': "#ffcc99"},
                           {'range': [90, 100], 'color': "#99ff99"}]}
            ))
            st.plotly_chart(fig_gauge, use_container_width=True)
            
    st.subheader("Edición y Gestión de Registros SST")
    edited_sst = st.data_editor(df_sst_filtered, num_rows="dynamic", key="editor_sst")
    if st.button("Guardar Cambios SST"):
        st.session_state["df_sst"] = edited_sst
        st.success("Cambios guardados en el sistema.")
        
    st.download_button(
        label="📥 Descargar Matriz SST Actualizada",
        data=to_excel_download(st.session_state["df_sst"], "ANUAL SST"),
        file_name="KPI_PLAN_ANUAL_SST_ACTUALIZADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =========================================================
# TAB: PLAN INSTRUMENTACIÓN
# =========================================================
with tab_inst:
    st.header("Plan de Inspección de Instrumentación y Válvulas")
    
    uploaded_inst = st.file_uploader("Reemplazar/Actualizar Matriz Instrumentación (Excel)", type=["xlsx"], key="u_inst")
    if uploaded_inst:
        xls = pd.ExcelFile(uploaded_inst)
        dfs = []
        for sheet in xls.sheet_names:
            if sheet != "Hoja1":
                dfs.append(pd.read_excel(uploaded_inst, sheet_name=sheet))
        if dfs:
            st.session_state["df_inst"] = pd.concat(dfs, ignore_index=True)
            st.success("Matriz de Instrumentación cargada con éxito.")

    df_inst = st.session_state["df_inst"]
    years_inst, months_inst = extract_years_and_months(df_inst, "MES")
    
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    with col_f1:
        sel_y_inst = st.multiselect("📅 Filtrar por Año:", options=years_inst, default=years_inst, key="filter_inst_year")
    with col_f2:
        sel_m_inst = st.multiselect(
            "🗓️ Filtrar por Mes:",
            options=months_inst,
            default=months_inst,
            format_func=lambda x: f"{x:02d} - {MONTH_NAMES.get(x, '')}",
            key="filter_inst_month"
        )
    with col_f3:
        tipos = st.multiselect("Filtrar por Tipo:", options=df_inst["TIPO"].unique() if "TIPO" in df_inst.columns else [], default=df_inst["TIPO"].unique() if "TIPO" in df_inst.columns else [], key="filter_inst_tipo")
    with col_f4:
        unidades = st.multiselect("Filtrar por Unidad:", options=df_inst["UNIDAD"].unique() if "UNIDAD" in df_inst.columns else [], default=df_inst["UNIDAD"].unique() if "UNIDAD" in df_inst.columns else [], key="filter_inst_unidad")
        
    df_filtered_inst = filter_by_year_month(df_inst, "MES", sel_y_inst, sel_m_inst)
    if tipos and "TIPO" in df_filtered_inst.columns:
        df_filtered_inst = df_filtered_inst[df_filtered_inst["TIPO"].isin(tipos)]
    if unidades and "UNIDAD" in df_filtered_inst.columns:
        df_filtered_inst = df_filtered_inst[df_filtered_inst["UNIDAD"].isin(unidades)]
        
    if not df_filtered_inst.empty:
        fig_inst_bar = px.bar(
            df_filtered_inst, x="UNIDAD", y="AVANCE DE CAMPO", color="TIPO",
            title="Avance de Campo por Unidad Operativa e Instrumento",
            barmode="group"
        )
        st.plotly_chart(fig_inst_bar, use_container_width=True)

    st.subheader("Matriz de Inspección de Campo")
    edited_inst = st.data_editor(df_filtered_inst, num_rows="dynamic", key="editor_inst")
    if st.button("Guardar Cambios Instrumentación"):
        st.session_state["df_inst"] = edited_inst
        st.success("Matriz de Instrumentación guardada en sesión.")

    st.download_button(
        label="📥 Descargar Matriz Instrumentación Actualizada",
        data=to_excel_download(st.session_state["df_inst"], "Plan"),
        file_name="KPI_PLAN_INSTRUMENTACION_ACTUALIZADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =========================================================
# TAB: ACCIDENTES E INCIDENTES
# =========================================================
with tab_acc:
    st.header("Registro de Accidentes, Incidentes y Hallazgos")
    
    uploaded_acc = st.file_uploader("Subir Archivo de Accidentes (Excel)", type=["xlsx"], key="u_acc")
    if uploaded_acc:
        st.session_state["df_acc"] = pd.read_excel(uploaded_acc, sheet_name="SEG_ACCIDENTES")
        st.success("Registro de accidentes actualizado.")

    df_acc = st.session_state["df_acc"]
    years_acc, months_acc = extract_years_and_months(df_acc, "FECHA")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        sel_y_acc = st.multiselect("📅 Filtrar por Año:", options=years_acc, default=years_acc, key="filter_acc_year")
    with col_f2:
        sel_m_acc = st.multiselect(
            "🗓️ Filtrar por Mes:",
            options=months_acc,
            default=months_acc,
            format_func=lambda x: f"{x:02d} - {MONTH_NAMES.get(x, '')}",
            key="filter_acc_month"
        )
        
    df_acc_filtered = filter_by_year_month(df_acc, "FECHA", sel_y_acc, sel_m_acc)
    
    if not df_acc_filtered.empty:
        c1, c2 = st.columns(2)
        with c1:
            fig_acc_loc = px.pie(
                df_acc_filtered, names="UBICACIÓN", title="Eventos Registrados por Ubicación / Área Planta",
                hole=0.3
            )
            st.plotly_chart(fig_acc_loc, use_container_width=True)
        with c2:
            fig_acc_tipo = px.bar(
                df_acc_filtered, x="TIPO", color="UBICACIÓN", title="Eventos Clasificados por Tipo"
            )
            st.plotly_chart(fig_acc_tipo, use_container_width=True)

    st.subheader("Bitácora y Detalle de Incidentes")
    edited_acc = st.data_editor(df_acc_filtered, num_rows="dynamic", key="editor_acc")
    if st.button("Guardar Registro de Accidentes"):
        st.session_state["df_acc"] = edited_acc
        st.success("Bitácora de accidentes actualizada.")

    st.download_button(
        label="📥 Descargar Registro de Incidentes Actualizado",
        data=to_excel_download(st.session_state["df_acc"], "SEG_ACCIDENTES"),
        file_name="KPI_ACCIDENTES_ACTUALIZADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =========================================================
# TAB: CONTROL OPS
# =========================================================
with tab_ops:
    st.header("Control de Observaciones Preventivas de Seguridad (OPS)")
    
    uploaded_ops = st.file_uploader("Subir Matriz OPS (Excel)", type=["xlsx"], key="u_ops")
    if uploaded_ops:
        st.session_state["df_ops_gen"] = pd.read_excel(uploaded_ops, sheet_name="OPS_GENERADAS")
        st.session_state["df_ops_cuota"] = pd.read_excel(uploaded_ops, sheet_name="OPS_CUOTA")
        st.success("Control de OPS actualizado.")

    df_ops_gen = st.session_state["df_ops_gen"]
    years_ops, months_ops = extract_years_and_months(df_ops_gen, "FECHA")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        sel_y_ops = st.multiselect("📅 Filtrar por Año:", options=years_ops, default=years_ops, key="filter_ops_year")
    with col_f2:
        sel_m_ops = st.multiselect(
            "🗓️ Filtrar por Mes:",
            options=months_ops,
            default=months_ops,
            format_func=lambda x: f"{x:02d} - {MONTH_NAMES.get(x, '')}",
            key="filter_ops_month"
        )

    df_ops_filtered = filter_by_year_month(df_ops_gen, "FECHA", sel_y_ops, sel_m_ops)

    if not df_ops_filtered.empty:
        c1, c2 = st.columns(2)
        with c1:
            fig_ops_area = px.bar(
                df_ops_filtered, x="ÁREA", color="GRAVEDAD",
                title="OPS Generadas por Área Operativa y Gravedad",
                barmode="stack"
            )
            st.plotly_chart(fig_ops_area, use_container_width=True)
        with c2:
            fig_ops_obs = px.bar(
                df_ops_filtered, x="OBSERVADOR", title="Reporte de OPS por Inspector / Observador",
                color_discrete_sequence=["#2ca02c"]
            )
            st.plotly_chart(fig_ops_obs, use_container_width=True)

    st.subheader("Matriz de Observaciones Generadas (OPS)")
    edited_ops = st.data_editor(df_ops_filtered, num_rows="dynamic", key="editor_ops")
    if st.button("Guardar Cambios OPS"):
        st.session_state["df_ops_gen"] = edited_ops
        st.success("Registros OPS actualizados.")

    st.download_button(
        label="📥 Descargar Matriz OPS Actualizada",
        data=to_excel_download(st.session_state["df_ops_gen"], "OPS_GENERADAS"),
        file_name="KPI_OPS_ACTUALIZADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
```