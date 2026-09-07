# Gráfica de Avance por Unidad (Eje X limitado a 70 y nuevo título)
if not df_p_filt.empty and "UNIDAD" in df_p_filt.columns:
    # Aseguramos que la columna sea numérica para poder filtrar hasta la unidad 70
    df_p_chart = df_p_filt.copy()
    df_p_chart["UNIDAD_NUM"] = pd.to_numeric(df_p_chart["UNIDAD"], errors="coerce")
    
    # Filtrar solo unidades entre 0 y 70
    df_p_chart = df_p_chart[(df_p_chart["UNIDAD_NUM"] >= 0) & (df_p_chart["UNIDAD_NUM"] <= 70)]
    
    if not df_p_chart.empty:
        fig_p = px.bar(
            df_p_chart, 
            x="UNIDAD_NUM", 
            y="AVANCE DE CAMPO", 
            color="TIPO", 
            title="Avance Plan General por Unidad", 
            barmode="group",
            labels={"UNIDAD_NUM": "UNIDAD"}
        )
        # Limitar visualmente el eje X de 0 a 70
        fig_p.update_xaxes(range=[0, 70], dtick=10)
        
        st.plotly_chart(fig_p, use_container_width=True)