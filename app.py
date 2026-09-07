import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
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
    except Exception:
        df_inst_plan = pd.DataFrame(columns=["PLAN", "TIPO", "MES", "UNIDAD", "TAG", "AVANCE DE CAMPO", "COMENTARIO"])

    try:
        df_inst_vaar = pd.read_excel("KPI PLAN INSTRUMENTACION.xlsx", sheet_name="Válvulas VAAR- Sensores")
    except Exception:
        df_inst_vaar = pd.DataFrame(columns=["PLAN", "TIPO", "MES", "UNIDAD", "TAG", "AVANCE DE CAMPO", "COMENTARIO"])

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

    return df_sst, df_inst_plan, df_inst_vaar, df_acc, df_ops_gen, df_ops_cuota

if "df_sst" not in st.session_state:
    df_sst, df_inst_plan, df_inst_vaar, df_acc, df_ops_gen, df_ops_cuota = load_default_data()
    st.session_state["df_sst"] = df_sst
    st.session_state["df_inst_plan"] = df_inst_plan
    st.session_state["df_inst_vaar"] = df_inst_vaar
    st.session_state["df_acc"] = df_acc
    st.session_state["df_ops_gen"] = df_ops_gen
    st.session_state["df_ops_cuota"] = df_ops_cuota

def to_excel_download(df, sheet_name="Datos"):
    output = io.BytesIO()
    cols_to_drop = [c for c in ["Year_Temp", "Month_Temp", "Month_Num_Temp"] if c in df.columns]
    df_clean = df.drop(columns=cols_to_drop)
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_clean.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

def extract_year_month(df, date_col):
    df_copy = df.copy()
    if df_copy.empty or date_col not in df_copy.columns:
        df_copy["Year_Temp"] = None
        df_copy["Month_Temp"] = None
        df_copy["Month_Num_Temp"] = None
        return df_copy, [], []
    
    dates = pd.to_datetime(df_copy[date_col], errors='coerce')
    df_copy["Year_Temp"] = dates.dt.year.astype("Int64")
    df_copy["Month_Temp"] = dates.dt.month_name()
    df_copy["Month_Num_Temp"] = dates.dt.month
    
    years = sorted([int(y) for y in df_copy["Year_Temp"].dropna().unique()])
    
    month_order = ["January", "February", "March", "April", "May", "June", 
                   "July", "August", "September", "October", "November", "December"]
    found_months = df_copy["Month_Temp"].dropna().unique()
    months = [m for m in month_order if m in found_months]
    
    return df_copy, years, months

def filter_df(df_in, years, months):
    df_out = df_in.copy()
    if years and "Year_Temp" in df_out.columns:
        df_out = df_out[df_out["Year_Temp"].isin(years)]
    if months and "Month_Temp" in df_out.columns:
        df_out = df_out[df_out["Month_Temp"].isin(months)]
    return df_out

def plot_monthly_trend(df_input, date_col="MES", title="Tendencia Mensual: Programados vs. Inspeccionados (100%)"):
    if df_input.empty or date_col not in df_input.columns:
        return None
    
    df_temp = df_input.copy()
    df_temp["Fecha_DT"] = pd.to_datetime(df_temp[date_col], errors='coerce')
    df_temp = df_temp.dropna(subset=["Fecha_DT"])
    
    if df_temp.empty:
        return None

    df_temp["Mes_Periodo"] = df_temp["Fecha_DT"].dt.to_period("M").astype(str)
    
    df_grouped = df_temp.groupby("Mes_Periodo").agg(
        Programados=('TAG', 'count'),
        Inspeccionados=('AVANCE DE CAMPO', lambda x: (x == 1).sum())
    ).reset_index()

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # Traza de Barras (Programados) con etiquetas de datos
    fig.add_trace(
        go.Bar(
            x=df_grouped["Mes_Periodo"],
            y=df_grouped["Programados"],
            name="Programados",
            marker_color="#1f77b4",
            opacity=0.7,
            text=df_grouped["Programados"],
            textposition="outside"
        ),
        secondary_y=False
    )

    # Traza de Líneas (Inspeccionados) con etiquetas de datos superiores
    fig.add_trace(
        go.Scatter(
            x=df_grouped["Mes_Periodo"],
            y=df_grouped["Inspeccionados"],
            name="Inspeccionados (100%)",
            mode="lines+markers+text",
            text=df_grouped["Inspeccionados"],
            textposition="top center",
            textfont=dict(size=12, color="black"),
            line=dict(color="#2ca02c", width=3),
            marker=dict(size=8)
        ),
        secondary_y=True
    )

    fig.update_layout(
        title_text=title,
        xaxis_title="Mes",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hovermode="x unified"
    )
    
    fig.update_yaxes(title_text="Cant. Instrumentos Programados", secondary_y=False)
    fig.update_yaxes(title_text="Cant. Instrumentos al 100%", secondary_y=True)

    return fig

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
# PESTAÑA 1: RESUMEN EJECUTIVO
# =========================================================
with tab_resumen:
    st.header("KPIs Principales del Proyecto")
    
    df_inst_full = pd.concat([st.session_state["df_inst_plan"], st.session_state["df_inst_vaar"]], ignore_index=True)
    
    df_sst_proc, y_sst, m_sst = extract_year_month(st.session_state["df_sst"], "Mes")
    df_inst_proc, y_inst, m_inst = extract_year_month(df_inst_full, "MES")
    df_acc_proc, y_acc, m_acc = extract_year_month(st.session_state["df_acc"], "FECHA")
    df_ops_proc, y_ops, m_ops = extract_year_month(st.session_state["df_ops_gen"], "FECHA")
    
    all_years = sorted(list(set(y_sst + y_inst + y_acc + y_ops)))
    all_months = list(dict.fromkeys(m_sst + m_inst + m_acc + m_ops))
    
    col_y, col_m = st.columns(2)
    with col_y:
        sel_years_res = st.multiselect("📅 Filtrar por Año (Resumen General):", options=all_years, default=all_years, key="filter_res_year")
    with col_m:
        sel_months_res = st.multiselect("🗓️ Filtrar por Mes (Resumen General):", options=all_months, default=all_months, key="filter_res_month")

    df_sst_r = filter_df(df_sst_proc, sel_years_res, sel_months_res)
    df_inst_r = filter_df(df_inst_proc, sel_years_res, sel_months_res)
    df_acc_r = filter_df(df_acc_proc, sel_years_res, sel_months_res)
    df_ops_r = filter_df(df_ops_proc, sel_years_res, sel_months_res)

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
                labels={"AVANCE DE CAMPO": "Avance Promedio (0-1)", "UNIDAD": "Unidad Operativa"},
                text_auto='.2f'
            )
            fig_inst_summary.update_traces(textposition='outside')
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
# PESTAÑA 2: PLAN ANUAL SST
# =========================================================
with tab_sst:
    st.header("Plan Anual de Seguridad y Salud en el Trabajo (SST)")
    
    uploaded_sst = st.file_uploader("Reemplazar/Actualizar Matriz SST (Excel)", type=["xlsx"], key="u_sst")
    if uploaded_sst:
        st.session_state["df_sst"] = pd.read_excel(uploaded_sst)
        st.success("Matriz SST actualizada correctamente.")
        
    df_sst_proc, years_sst, months_sst = extract_year_month(st.session_state["df_sst"], "Mes")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        sel_years_sst = st.multiselect("📅 Filtrar por Año (SST):", options=years_sst, default=years_sst, key="filter_sst_year")
    with col_f2:
        sel_months_sst = st.multiselect("🗓️ Filtrar por Mes (SST):", options=months_sst, default=months_sst, key="filter_sst_month")
    
    df_sst_filtered = filter_df(df_sst_proc, sel_years_sst, sel_months_sst)

    if not df_sst_filtered.empty:
        c1, c2 = st.columns([2, 1])
        with c1:
            fig_sst_bar = px.bar(
                df_sst_filtered, x="Tipo", y="% Cumplimiento", color="Tipo",
                title="Cumplimiento Promedio por Tipo de Actividad SST",
                labels={"% Cumplimiento": "Cumplimiento (0.0 a 1.0)"},
                text_auto='.2f'
            )
            fig_sst_bar.update_traces(textposition='outside')
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
    cols_to_show = [c for c in df_sst_filtered.columns if c not in ["Year_Temp", "Month_Temp", "Month_Num_Temp"]]
    edited_sst = st.data_editor(df_sst_filtered[cols_to_show], num_rows="dynamic", key="editor_sst")
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
# PESTAÑA 3: PLAN INSTRUMENTACIÓN
# =========================================================
with tab_inst:
    st.header("Plan de Inspección de Instrumentación, Válvulas y Sensores")
    
    uploaded_inst = st.file_uploader("Reemplazar/Actualizar Matriz Instrumentación (Excel)", type=["xlsx"], key="u_inst")
    if uploaded_inst:
        xls = pd.ExcelFile(uploaded_inst)
        if "Plan" in xls.sheet_names:
            st.session_state["df_inst_plan"] = pd.read_excel(uploaded_inst, sheet_name="Plan")
        if "Válvulas VAAR- Sensores" in xls.sheet_names:
            st.session_state["df_inst_vaar"] = pd.read_excel(uploaded_inst, sheet_name="Válvulas VAAR- Sensores")
        st.success("Matriz de Instrumentación cargada con éxito.")

    subtab_plan, subtab_vaar, subtab_sensores = st.tabs([
        "📋 Plan General",
        "🚰 Válvulas VAAR",
        "📡 Sensores de Vibración"
    ])

    # --- SUBTAB 1: PLAN GENERAL ---
    with subtab_plan:
        st.subheader("Plan General de Instrumentación")
        df_p_proc, years_p, months_p = extract_year_month(st.session_state["df_inst_plan"], "MES")
        
        # Filtros
        c1, c2, c3 = st.columns(3)
        with c1:
            s_yp = st.multiselect("📅 Año:", options=years_p, default=years_p, key="f_p_y")
        with c2:
            s_mp = st.multiselect("🗓️ Mes:", options=months_p, default=months_p, key="f_p_m")
        with c3:
            units_plan = df_p_proc["UNIDAD"].dropna().astype(str).unique() if "UNIDAD" in df_p_proc.columns else []
            units_plan_clean = [u for u in units_plan if u.strip().upper() != "MAX U 63"]
            s_up = st.multiselect("Unidad:", options=units_plan_clean, key="f_p_u")

        df_p_filt = filter_df(df_p_proc, s_yp, s_mp)
        if s_up and "UNIDAD" in df_p_filt.columns:
            df_p_filt = df_p_filt[df_p_filt["UNIDAD"].isin(s_up)]

        # Tarjetas de Indicadores KPI
        total_p_prog = len(df_p_filt)
        total_p_ejec = int((df_p_filt["AVANCE DE CAMPO"] == 1).sum()) if "AVANCE DE CAMPO" in df_p_filt.columns else 0
        pct_p_avance = (total_p_ejec / total_p_prog * 100) if total_p_prog > 0 else 0.0

        m1, m2, m3 = st.columns(3)
        m1.metric("📋 Total Programados", f"{total_p_prog:,}")
        m2.metric("✅ Total Inspeccionados (100%)", f"{total_p_ejec:,}")
        m3.metric("📊 % Avance (Ejecutado vs Programado)", f"{pct_p_avance:.1f}%")
        st.markdown("---")

        # Gráfica de Tendencia Mensual (Doble Eje Y) con Etiquetas de Datos
        fig_p_monthly = plot_monthly_trend(df_p_filt, title="Indicador Mensual: Programados vs. Inspeccionados al 100% (Plan General)")
        if fig_p_monthly:
            st.plotly_chart(fig_p_monthly, use_container_width=True)

        # Gráfica de Avance por Unidad con Etiquetas de Datos visibles
        if not df_p_filt.empty and "UNIDAD" in df_p_filt.columns:
            df_p_chart = df_p_filt.copy()
            df_p_chart["UNIDAD_NUM"] = pd.to_numeric(df_p_chart["UNIDAD"], errors="coerce")
            
            df_p_chart = df_p_chart[(df_p_chart["UNIDAD_NUM"] >= 0) & (df_p_chart["UNIDAD_NUM"] <= 70)]
            
            if not df_p_chart.empty:
                fig_p = px.bar(
                    df_p_chart, 
                    x="UNIDAD_NUM", 
                    y="AVANCE DE CAMPO", 
                    color="TIPO", 
                    title="Avance Plan General por Unidad", 
                    barmode="group",
                    labels={"UNIDAD_NUM": "UNIDAD"},
                    text_auto=True
                )
                fig_p.update_traces(textposition="outside")
                fig_p.update_xaxes(range=[0, 70], dtick=10)
                st.plotly_chart(fig_p, use_container_width=True)

        cols_show_p = [c for c in df_p_filt.columns if c not in ["Year_Temp", "Month_Temp", "Month_Num_Temp"]]
        edited_p = st.data_editor(df_p_filt[cols_show_p], num_rows="dynamic", key="ed_p")
        if st.button("Guardar Plan General"):
            st.session_state["df_inst_plan"] = edited_p
            st.success("Plan General guardado.")

    # --- SUBTAB 2: VÁLVULAS VAAR ---
    with subtab_vaar:
        st.subheader("Plan de Válvulas VAAR")
        df_vaar_base = st.session_state["df_inst_vaar"]
        if not df_vaar_base.empty and "TIPO" in df_vaar_base.columns:
            df_vaar_base = df_vaar_base[df_vaar_base["TIPO"].str.contains("VAAR", case=False, na=False)]

        df_v_proc, years_v, months_v = extract_year_month(df_vaar_base, "MES")
        
        c1, c2, c3 = st.columns(3)
        with c1:
            s_yv = st.multiselect("📅 Año:", options=years_v, default=years_v, key="f_v_y")
        with c2:
            s_mv = st.multiselect("🗓️ Mes:", options=months_v, default=months_v, key="f_v_m")
        with c3:
            s_uv = st.multiselect("Unidad:", options=df_v_proc["UNIDAD"].unique() if "UNIDAD" in df_v_proc.columns else [], key="f_v_u")

        df_v_filt = filter_df(df_v_proc, s_yv, s_mv)
        if s_uv and "UNIDAD" in df_v_filt.columns:
            df_v_filt = df_v_filt[df_v_filt["UNIDAD"].isin(s_uv)]

        total_v_prog = len(df_v_filt)
        total_v_ejec = int((df_v_filt["AVANCE DE CAMPO"] == 1).sum()) if "AVANCE DE CAMPO" in df_v_filt.columns else 0
        pct_v_avance = (total_v_ejec / total_v_prog * 100) if total_v_prog > 0 else 0.0

        m1, m2, m3 = st.columns(3)
        m1.metric("🚰 Válvulas Programadas", f"{total_v_prog:,}")
        m2.metric("✅ Válvulas Inspeccionadas (100%)", f"{total_v_ejec:,}")
        m3.metric("📊 % Avance (Ejecutado vs Programado)", f"{pct_v_avance:.1f}%")
        st.markdown("---")

        fig_v_monthly = plot_monthly_trend(df_v_filt, title="Indicador Mensual: Programados vs. Inspeccionados al 100% (Válvulas VAAR)")
        if fig_v_monthly:
            st.plotly_chart(fig_v_monthly, use_container_width=True)

        if not df_v_filt.empty:
            fig_v = px.bar(df_v_filt, x="UNIDAD", y="AVANCE DE CAMPO", color="TAG", title="Avance Válvulas VAAR por Unidad", barmode="group", text_auto=True)
            fig_v.update_traces(textposition="outside")
            st.plotly_chart(fig_v, use_container_width=True)

        cols_show_v = [c for c in df_v_filt.columns if c not in ["Year_Temp", "Month_Temp", "Month_Num_Temp"]]
        edited_v = st.data_editor(df_v_filt[cols_show_v], num_rows="dynamic", key="ed_v")
        if st.button("Guardar Válvulas VAAR"):
            df_other = st.session_state["df_inst_vaar"][~st.session_state["df_inst_vaar"]["TIPO"].str.contains("VAAR", case=False, na=False)]
            st.session_state["df_inst_vaar"] = pd.concat([df_other, edited_v], ignore_index=True)
            st.success("Plan Válvulas VAAR guardado.")

    # --- SUBTAB 3: SENSORES DE VIBRACIÓN ---
    with subtab_sensores:
        st.subheader("Plan de Sensores de Vibración")
        df_sens_base = st.session_state["df_inst_vaar"]
        if not df_sens_base.empty and "TIPO" in df_sens_base.columns:
            df_sens_base = df_sens_base[df_sens_base["TIPO"].str.contains("SENSOR", case=False, na=False)]

        df_s_proc, years_s, months_s = extract_year_month(df_sens_base, "MES")
        
        c1, c2, c3 = st.columns(3)
        with c1:
            s_ys = st.multiselect("📅 Año:", options=years_s, default=years_s, key="f_s_y")
        with c2:
            s_ms = st.multiselect("🗓️ Mes:", options=months_s, default=months_s, key="f_s_m")
        with c3:
            s_us = st.multiselect("Unidad:", options=df_s_proc["UNIDAD"].unique() if "UNIDAD" in df_s_proc.columns else [], key="f_s_u")

        df_s_filt = filter_df(df_s_proc, s_ys, s_ms)
        if s_us and "UNIDAD" in df_s_filt.columns:
            df_s_filt = df_s_filt[df_s_filt["UNIDAD"].isin(s_us)]

        total_s_prog = len(df_s_filt)
        total_s_ejec = int((df_s_filt["AVANCE DE CAMPO"] == 1).sum()) if "AVANCE DE CAMPO" in df_s_filt.columns else 0
        pct_s_avance = (total_s_ejec / total_s_prog * 100) if total_s_prog > 0 else 0.0

        m1, m2, m3 = st.columns(3)
        m1.metric("📡 Sensores Programados", f"{total_s_prog:,}")
        m2.metric("✅ Sensores Inspeccionados (100%)", f"{total_s_ejec:,}")
        m3.metric("📊 % Avance (Ejecutado vs Programado)", f"{pct_s_avance:.1f}%")
        st.markdown("---")

        fig_s_monthly = plot_monthly_trend(df_s_filt, title="Indicador Mensual: Programados vs. Inspeccionados al 100% (Sensores de Vibración)")
        if fig_s_monthly:
            st.plotly_chart(fig_s_monthly, use_container_width=True)

        if not df_s_filt.empty:
            fig_s = px.bar(df_s_filt, x="UNIDAD", y="AVANCE DE CAMPO", color="TAG", title="Avance Sensores de Vibración por Unidad", barmode="group", text_auto=True)
            fig_s.update_traces(textposition="outside")
            st.plotly_chart(fig_s, use_container_width=True)

        cols_show_s = [c for c in df_s_filt.columns if c not in ["Year_Temp", "Month_Temp", "Month_Num_Temp"]]
        edited_s = st.data_editor(df_s_filt[cols_show_s], num_rows="dynamic", key="ed_s")
        if st.button("Guardar Sensores"):
            df_other = st.session_state["df_inst_vaar"][~st.session_state["df_inst_vaar"]["TIPO"].str.contains("SENSOR", case=False, na=False)]
            st.session_state["df_inst_vaar"] = pd.concat([df_other, edited_s], ignore_index=True)
            st.success("Plan Sensores guardado.")

    st.markdown("---")
    df_inst_download = pd.concat([st.session_state["df_inst_plan"], st.session_state["df_inst_vaar"]], ignore_index=True)
    st.download_button(
        label="📥 Descargar Matriz Consolidada de Instrumentación Actualizada",
        data=to_excel_download(df_inst_download, "Plan"),
        file_name="KPI_PLAN_INSTRUMENTACION_ACTUALIZADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =========================================================
# PESTAÑA 4: ACCIDENTES E INCIDENTES
# =========================================================
with tab_acc:
    st.header("Registro de Accidentes, Incidentes y Hallazgos")
    
    uploaded_acc = st.file_uploader("Subir Archivo de Accidentes (Excel)", type=["xlsx"], key="u_acc")
    if uploaded_acc:
        st.session_state["df_acc"] = pd.read_excel(uploaded_acc, sheet_name="SEG_ACCIDENTES")
        st.success("Registro de accidentes actualizado.")

    df_acc_proc, years_acc, months_acc = extract_year_month(st.session_state["df_acc"], "FECHA")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        sel_years_acc = st.multiselect("📅 Filtrar por Año:", options=years_acc, default=years_acc, key="filter_acc_year")
    with col_f2:
        sel_months_acc = st.multiselect("🗓️ Filtrar por Mes:", options=months_acc, default=months_acc, key="filter_acc_month")
        
    df_acc_filtered = filter_df(df_acc_proc, sel_years_acc, sel_months_acc)
    
    if not df_acc_filtered.empty:
        c1, c2 = st.columns(2)
        with c1:
            if "UBICACIÓN" in df_acc_filtered.columns:
                fig_acc_loc = px.pie(
                    df_acc_filtered, names="UBICACIÓN", title="Eventos Registrados por Ubicación / Área Planta",
                    hole=0.3
                )
                st.plotly_chart(fig_acc_loc, use_container_width=True)
        with c2:
            if "TIPO" in df_acc_filtered.columns and "UBICACIÓN" in df_acc_filtered.columns:
                fig_acc_tipo = px.bar(
                    df_acc_filtered, x="TIPO", color="UBICACIÓN", title="Eventos Clasificados por Tipo", text_auto=True
                )
                fig_acc_tipo.update_traces(textposition="outside")
                st.plotly_chart(fig_acc_tipo, use_container_width=True)

    st.subheader("Bitácora y Detalle de Incidentes")
    cols_to_show_acc = [c for c in df_acc_filtered.columns if c not in ["Year_Temp", "Month_Temp", "Month_Num_Temp"]]
    edited_acc = st.data_editor(df_acc_filtered[cols_to_show_acc], num_rows="dynamic", key="editor_acc")
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
# PESTAÑA 5: CONTROL OPS
# =========================================================
with tab_ops:
    st.header("Control de Observaciones Preventivas de Seguridad (OPS)")
    
    uploaded_ops = st.file_uploader("Subir Matriz OPS (Excel)", type=["xlsx"], key="u_ops")
    if uploaded_ops:
        st.session_state["df_ops_gen"] = pd.read_excel(uploaded_ops, sheet_name="OPS_GENERADAS")
        st.session_state["df_ops_cuota"] = pd.read_excel(uploaded_ops, sheet_name="OPS_CUOTA")
        st.success("Control de OPS actualizado.")

    df_ops_proc, years_ops, months_ops = extract_year_month(st.session_state["df_ops_gen"], "FECHA")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        sel_years_ops = st.multiselect("📅 Filtrar por Año:", options=years_ops, default=years_ops, key="filter_ops_year")
    with col_f2:
        sel_months_ops = st.multiselect("🗓️ Filtrar por Mes:", options=months_ops, default=months_ops, key="filter_ops_month")

    df_ops_filtered = filter_df(df_ops_proc, sel_years_ops, sel_months_ops)

    if not df_ops_filtered.empty:
        c1, c2 = st.columns(2)
        with c1:
            if "ÁREA" in df_ops_filtered.columns and "GRAVEDAD" in df_ops_filtered.columns:
                fig_ops_area = px.bar(
                    df_ops_filtered, x="ÁREA", color="GRAVEDAD",
                    title="OPS Generadas por Área Operativa y Gravedad",
                    barmode="stack", text_auto=True
                )
                st.plotly_chart(fig_ops_area, use_container_width=True)
        with c2:
            if "OBSERVADOR" in df_ops_filtered.columns:
                fig_ops_obs = px.bar(
                    df_ops_filtered, x="OBSERVADOR", title="Reporte de OPS por Inspector / Observador",
                    color_discrete_sequence=["#2ca02c"], text_auto=True
                )
                fig_ops_obs.update_traces(textposition="outside")
                st.plotly_chart(fig_ops_obs, use_container_width=True)

    st.subheader("Matriz de Observaciones Generadas (OPS)")
    cols_to_show_ops = [c for c in df_ops_filtered.columns if c not in ["Year_Temp", "Month_Temp", "Month_Num_Temp"]]
    edited_ops = st.data_editor(df_ops_filtered[cols_to_show_ops], num_rows="dynamic", key="editor_ops")
    if st.button("Guardar Cambios OPS"):
        st.session_state["df_ops_gen"] = edited_ops
        st.success("Registros OPS actualizados.")

    st.download_button(
        label="📥 Descargar Matriz OPS Actualizada",
        data=to_excel_download(st.session_state["df_ops_gen"], "OPS_GENERADAS"),
        file_name="KPI_OPS_ACTUALIZADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )