import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io

# ---------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA
# ---------------------------------------------------------
st.set_page_config(
    page_title="Dashboard Gerencial - Inspecciones Hidrocarburos",
    page_icon="🛢️",
    layout="wide"
)

st.title("🛢️ Sistema Integrado de Control de Inspecciones & SST")
st.caption("Plataforma Gerencial de Monitoreo de KPIs de Inspección, Seguridad y Operaciones")

# ---------------------------------------------------------
# CARGA DE DATOS E INICIALIZACIÓN DE ESTADO (SESSION STATE)
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

# Cargar a st.session_state si no existe
if "df_sst" not in st.session_state:
    df_sst, df_inst, df_acc, df_ops_gen, df_ops_cuota = load_default_data()
    st.session_state["df_sst"] = df_sst
    st.session_state["df_inst"] = df_inst
    st.session_state["df_acc"] = df_acc
    st.session_state["df_ops_gen"] = df_ops_gen
    st.session_state["df_ops_cuota"] = df_ops_cuota

# Función auxiliar para descargar DataFrames en formato Excel
def to_excel_download(df, sheet_name="Datos"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    processed_data = output.getvalue()
    return processed_data

# ---------------------------------------------------------
# PESTAÑAS PRINCIPALES
# ---------------------------------------------------------
tab_resumen, tab_sst, tab_inst, tab_acc, tab_ops = st.tabs([
    "📊 Resumen Ejecutivo",
    "🛡️ Plan Anual SST",
    "🔧 Instrumentación",
    "🚨 Accidentes e Incidentes",
    "👁️ Control OPS"
])

# =========================================================
# PESTAÑA: RESUMEN EJECUTIVO
# =========================================================
with tab_resumen:
    st.header("KPIs Principales del Proyecto")
    
    col1, col2, col3, col4 = st.columns(4)
    
    # KPI SST
    df_sst = st.session_state["df_sst"]
    avg_sst = (df_sst["% Cumplimiento"].mean() * 100) if not df_sst.empty and "% Cumplimiento" in df_sst.columns else 0
    col1.metric("Cumplimiento SST", f"{avg_sst:.1f}%")
    
    # KPI Instrumentación
    df_inst = st.session_state["df_inst"]
    avg_inst = (df_inst["AVANCE DE CAMPO"].mean() * 100) if not df_inst.empty and "AVANCE DE CAMPO" in df_inst.columns else 0
    col2.metric("Avance Instrumentación", f"{avg_inst:.1f}%")
    
    # KPI Incidentes
    df_acc = st.session_state["df_acc"]
    total_acc = len(df_acc) if not df_acc.empty else 0
    col3.metric("Total Incidentes/Accidentes", f"{total_acc}")
    
    # KPI OPS
    df_ops_gen = st.session_state["df_ops_gen"]
    total_ops = len(df_ops_gen) if not df_ops_gen.empty else 0
    col4.metric("OPS Generadas", f"{total_ops}")
    
    st.markdown("---")
    st.subheader("Estado de Avance General de Planes de Inspección y Seguridad")
    
    c1, c2 = st.columns(2)
    with c1:
        if not df_inst.empty and "TIPO" in df_inst.columns and "AVANCE DE CAMPO" in df_inst.columns:
            fig_inst_summary = px.histogram(
                df_inst, x="UNIDAD", y="AVANCE DE CAMPO", color="TIPO",
                histfunc="avg", barmode="group",
                title="Promedio de Avance por Unidad y Tipo de Equipo",
                labels={"AVANCE DE CAMPO": "Avance Promedio (0-1)", "UNIDAD": "Unidad Operativa"}
            )
            st.plotly_chart(fig_inst_summary, use_container_width=True)
    
    with c2:
        if not df_ops_gen.empty and "GRAVEDAD" in df_ops_gen.columns:
            fig_ops_pie = px.pie(
                df_ops_gen, names="GRAVEDAD",
                title="Distribución de Observaciones de Seguridad (OPS) por Gravedad",
                hole=0.4
            )
            st.plotly_chart(fig_ops_pie, use_container_width=True)

# =========================================================
# PESTAÑA: PLAN ANUAL SST
# =========================================================
with tab_sst:
    st.header("Plan Anual de Seguridad y Salud en el Trabajo (SST)")
    
    # Subir Matriz
    uploaded_sst = st.file_uploader("Reemplazar/Actualizar Matriz SST (Excel)", type=["xlsx"], key="u_sst")
    if uploaded_sst:
        st.session_state["df_sst"] = pd.read_excel(uploaded_sst)
        st.success("Matriz SST actualizada correctamente.")
        
    df_sst = st.session_state["df_sst"]
    
    # Visualizaciones sugeridas
    if not df_sst.empty:
        c1, c2 = st.columns([2, 1])
        with c1:
            fig_sst_bar = px.bar(
                df_sst, x="Tipo", y="% Cumplimiento", color="Tipo",
                title="Cumplimiento Promedio por Tipo de Actividad SST",
                labels={"% Cumplimiento": "Cumplimiento (0.0 a 1.0)"}
            )
            st.plotly_chart(fig_sst_bar, use_container_width=True)
        with c2:
            avg_val = df_sst["% Cumplimiento"].mean() * 100
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
            
    # Edición e Inserción de Registros
    st.subheader("Edición y Gestión de Registros SST")
    edited_sst = st.data_editor(df_sst, num_rows="dynamic", key="editor_sst")
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
# PESTAÑA: PLAN INSTRUMENTACIÓN
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
    
    if not df_inst.empty:
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            tipos = st.multiselect("Filtrar por Tipo:", options=df_inst["TIPO"].unique() if "TIPO" in df_inst.columns else [], default=df_inst["TIPO"].unique() if "TIPO" in df_inst.columns else [])
        with col_f2:
            unidades = st.multiselect("Filtrar por Unidad:", options=df_inst["UNIDAD"].unique() if "UNIDAD" in df_inst.columns else [], default=df_inst["UNIDAD"].unique() if "UNIDAD" in df_inst.columns else [])
            
        df_filtered_inst = df_inst.copy()
        if tipos:
            df_filtered_inst = df_filtered_inst[df_filtered_inst["TIPO"].isin(tipos)]
        if unidades:
            df_filtered_inst = df_filtered_inst[df_filtered_inst["UNIDAD"].isin(unidades)]
            
        # Gráfica
        fig_inst_bar = px.bar(
            df_filtered_inst, x="UNIDAD", y="AVANCE DE CAMPO", color="TIPO",
            title="Avance de Campo por Unidad Operativa e Instrumento",
            barmode="group"
        )
        st.plotly_chart(fig_inst_bar, use_container_width=True)

    st.subheader("Matriz de Inspección de Campo")
    edited_inst = st.data_editor(df_inst, num_rows="dynamic", key="editor_inst")
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
# PESTAÑA: ACCIDENTES E INCIDENTES
# =========================================================
with tab_acc:
    st.header("Registro de Accidentes, Incidentes y Hallazgos")
    
    uploaded_acc = st.file_uploader("Subir Archivo de Accidentes (Excel)", type=["xlsx"], key="u_acc")
    if uploaded_acc:
        st.session_state["df_acc"] = pd.read_excel(uploaded_acc, sheet_name="SEG_ACCIDENTES")
        st.success("Registro de accidentes actualizado.")

    df_acc = st.session_state["df_acc"]
    
    if not df_acc.empty:
        c1, c2 = st.columns(2)
        with c1:
            fig_acc_loc = px.pie(
                df_acc, names="UBICACIÓN", title="Eventos Registrados por Ubicación / Área Planta",
                hole=0.3
            )
            st.plotly_chart(fig_acc_loc, use_container_width=True)
        with c2:
            fig_acc_tipo = px.bar(
                df_acc, x="TIPO", color="UBICACIÓN", title="Eventos Clasificados por Tipo"
            )
            st.plotly_chart(fig_acc_tipo, use_container_width=True)

    st.subheader("Bitácora y Detalle de Incidentes")
    edited_acc = st.data_editor(df_acc, num_rows="dynamic", key="editor_acc")
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
# PESTAÑA: CONTROL OPS
# =========================================================
with tab_ops:
    st.header("Control de Observaciones Preventivas de Seguridad (OPS)")
    
    uploaded_ops = st.file_uploader("Subir Matriz OPS (Excel)", type=["xlsx"], key="u_ops")
    if uploaded_ops:
        st.session_state["df_ops_gen"] = pd.read_excel(uploaded_ops, sheet_name="OPS_GENERADAS")
        st.session_state["df_ops_cuota"] = pd.read_excel(uploaded_ops, sheet_name="OPS_CUOTA")
        st.success("Control de OPS actualizado.")

    df_ops_gen = st.session_state["df_ops_gen"]
    
    if not df_ops_gen.empty:
        c1, c2 = st.columns(2)
        with c1:
            fig_ops_area = px.bar(
                df_ops_gen, x="ÁREA", color="GRAVEDAD",
                title="OPS Generadas por Área Operativa y Gravedad",
                barmode="stack"
            )
            st.plotly_chart(fig_ops_area, use_container_width=True)
        with c2:
            fig_ops_obs = px.bar(
                df_ops_gen, x="OBSERVADOR", title="Reporte de OPS por Inspector / Observador",
                color_discrete_sequence=["#2ca02c"]
            )
            st.plotly_chart(fig_ops_obs, use_container_width=True)

    st.subheader("Matriz de Observaciones Generadas (OPS)")
    edited_ops = st.data_editor(df_ops_gen, num_rows="dynamic", key="editor_ops")
    if st.button("Guardar Cambios OPS"):
        st.session_state["df_ops_gen"] = edited_ops
        st.success("Registros OPS actualizados.")

    st.download_button(
        label="📥 Descargar Matriz OPS Actualizada",
        data=to_excel_download(st.session_state["df_ops_gen"], "OPS_GENERADAS"),
        file_name="KPI_OPS_ACTUALIZADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
