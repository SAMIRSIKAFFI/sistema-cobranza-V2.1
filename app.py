import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

st.set_page_config(page_title="SISTEMA DE COBRANZA - RESULTADOS", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        padding: 1rem;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 10px;
        border-left: 5px solid #1f77b4;
    }
    .success-card {
        border-left-color: #28a745;
    }
    .warning-card {
        border-left-color: #ffc107;
    }
    .danger-card {
        border-left-color: #dc3545;
    }
    </style>
""", unsafe_allow_html=True)

st.sidebar.title("🏢 SISTEMA DE COBRANZA")
st.sidebar.markdown("---")

menu = st.sidebar.radio(
    "📋 MENÚ PRINCIPAL",
    [
        "📊 Dashboard Cruce Deuda vs Pagos",
        "📈 Gráficos Interactivos",
        "📲 GENERADOR DE SMS",
        "🚧 Módulo Histórico (En Desarrollo)"
    ]
)

def modulo_cruce():
    st.markdown('<div class="main-header">⚖️ DASHBOARD EJECUTIVO DE GESTIÓN DE COBRANZA</div>', unsafe_allow_html=True)

    def limpiar_columnas(df):
        df.columns = df.columns.str.strip().str.upper().str.replace(" ", "_")
        return df

    if "df_deuda_base" not in st.session_state:
        st.session_state.df_deuda_base = None

    if st.session_state.df_deuda_base is None:
        st.info("🔹 **Paso 1:** Carga la base de CARTERA/DEUDA (se guardará en memoria)")
        
        archivo_deuda = st.file_uploader(
            "📂 Subir archivo CARTERA / DEUDA",
            type=["xlsx"],
            help="Debe contener las columnas: ID_COBRANZA, PERIODO, DEUDA, TIPO"
        )

        if archivo_deuda:
            with st.spinner("Procesando cartera..."):
                try:
                    df_deuda = pd.read_excel(archivo_deuda)
                    df_deuda = limpiar_columnas(df_deuda)
                    columnas_deuda = {"ID_COBRANZA", "PERIODO", "DEUDA", "TIPO"}

                    if not columnas_deuda.issubset(df_deuda.columns):
                        st.error("❌ El archivo CARTERA no tiene las columnas obligatorias")
                        return

                    df_deuda["ID_COBRANZA"] = df_deuda["ID_COBRANZA"].astype(str)
                    df_deuda["PERIODO"] = df_deuda["PERIODO"].astype(str)
                    df_deuda["DEUDA"] = pd.to_numeric(df_deuda["DEUDA"], errors="coerce").fillna(0)

                    if (df_deuda["DEUDA"] < 0).any():
                        st.warning("⚠️ Se detectaron montos negativos en DEUDA.")
                        df_deuda["DEUDA"] = df_deuda["DEUDA"].abs()

                    st.session_state.df_deuda_base = df_deuda
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("📄 Registros", f"{len(df_deuda):,}")
                    with col2:
                        st.metric("💰 Cartera", f"Bs. {df_deuda['DEUDA'].sum():,.2f}")
                    with col3:
                        st.metric("📅 Periodos", df_deuda["PERIODO"].nunique())

                    st.success("✅ Cartera cargada correctamente")
                    st.balloons()
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
                    return
        return
    else:
        df_deuda = st.session_state.df_deuda_base
        col1, col2 = st.columns([3, 1])
        with col1:
            st.success("✅ **Cartera base cargada**")
        with col2:
            if st.button("🔄 Reemplazar", use_container_width=True):
                st.session_state.df_deuda_base = None
                st.rerun()

    st.markdown("---")
    st.info("🔹 **Paso 2:** Carga el archivo de PAGOS")
    
    archivo_pagos = st.file_uploader(
        "💵 Subir archivo PAGOS",
        type=["xlsx"],
        help="Debe contener: ID_COBRANZA, PERIODO, IMPORTE"
    )

    if not archivo_pagos:
        return

    with st.spinner("Procesando..."):
        try:
            df_deuda = st.session_state.df_deuda_base.copy()
            df_pagos = pd.read_excel(archivo_pagos)
            df_pagos = limpiar_columnas(df_pagos)
            
            columnas_pagos = {"ID_COBRANZA", "PERIODO", "IMPORTE"}
            if not columnas_pagos.issubset(df_pagos.columns):
                st.error("❌ El archivo PAGOS no tiene las columnas obligatorias")
                return

            df_pagos["ID_COBRANZA"] = df_pagos["ID_COBRANZA"].astype(str)
            df_pagos["PERIODO"] = df_pagos["PERIODO"].astype(str)
            df_pagos["IMPORTE"] = pd.to_numeric(df_pagos["IMPORTE"], errors="coerce").fillna(0)

            if (df_pagos["IMPORTE"] < 0).any():
                st.warning("⚠️ Montos negativos en PAGOS")
                df_pagos["IMPORTE"] = df_pagos["IMPORTE"].abs()

            pagos_resumen = df_pagos.groupby(["ID_COBRANZA", "PERIODO"])["IMPORTE"].sum().reset_index()
            pagos_resumen.rename(columns={"IMPORTE": "TOTAL_PAGADO"}, inplace=True)

            resultado = df_deuda.merge(pagos_resumen, on=["ID_COBRANZA", "PERIODO"], how="left")
            resultado["TOTAL_PAGADO"] = resultado["TOTAL_PAGADO"].fillna(0)
            resultado["SALDO_PENDIENTE"] = resultado["DEUDA"] - resultado["TOTAL_PAGADO"]
            resultado["SALDO_PENDIENTE"] = resultado["SALDO_PENDIENTE"].apply(lambda x: max(0, x))
            resultado["ESTADO"] = resultado.apply(
                lambda row: "✅ PAGADO" if row["TOTAL_PAGADO"] >= row["DEUDA"] else "⏳ PENDIENTE",
                axis=1
            )
            resultado["PORCENTAJE_PAGADO"] = (resultado["TOTAL_PAGADO"] / resultado["DEUDA"] * 100).round(2)
            resultado["PORCENTAJE_PAGADO"] = resultado["PORCENTAJE_PAGADO"].apply(lambda x: min(100, x))

            # Guardar resultado en session_state para gráficos
            st.session_state.resultado_cruce = resultado

            st.success("✅ Cruce realizado")
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            return

    st.markdown("---")
    st.markdown("## 📈 MÉTRICAS EJECUTIVAS")

    total_cartera = resultado["DEUDA"].sum()
    total_recuperado = resultado["TOTAL_PAGADO"].sum()
    saldo_pendiente = resultado["SALDO_PENDIENTE"].sum()
    porcentaje_recuperacion = (total_recuperado / total_cartera * 100) if total_cartera > 0 else 0
    total_casos = len(resultado)
    casos_pagados = len(resultado[resultado["ESTADO"] == "✅ PAGADO"])
    casos_pendientes = len(resultado[resultado["ESTADO"] == "⏳ PENDIENTE"])

    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("💼 CARTERA TOTAL", f"Bs. {total_cartera:,.2f}", f"{total_casos:,} casos")
    with col2:
        st.metric("✅ RECUPERADO", f"Bs. {total_recuperado:,.2f}", f"{porcentaje_recuperacion:.1f}%")
    with col3:
        st.metric("⏳ PENDIENTE", f"Bs. {saldo_pendiente:,.2f}", f"{casos_pendientes:,} casos")
    with col4:
        st.metric("📊 EFECTIVIDAD", f"{porcentaje_recuperacion:.1f}%", f"{casos_pagados:,} pagados")

    st.markdown("---")
    
    with st.expander("🔍 FILTROS", expanded=False):
        col1, col2, col3 = st.columns(3)
        with col1:
            periodos = ["Todos"] + sorted(resultado["PERIODO"].unique().tolist())
            filtro_periodo = st.selectbox("📅 Periodo", periodos)
        with col2:
            tipos = ["Todos"] + sorted(resultado["TIPO"].unique().tolist())
            filtro_tipo = st.selectbox("🏷️ Tipo", tipos)
        with col3:
            estados = ["Todos", "✅ PAGADO", "⏳ PENDIENTE"]
            filtro_estado = st.selectbox("📊 Estado", estados)

    resultado_filtrado = resultado.copy()
    if filtro_periodo != "Todos":
        resultado_filtrado = resultado_filtrado[resultado_filtrado["PERIODO"] == filtro_periodo]
    if filtro_tipo != "Todos":
        resultado_filtrado = resultado_filtrado[resultado_filtrado["TIPO"] == filtro_tipo]
    if filtro_estado != "Todos":
        resultado_filtrado = resultado_filtrado[resultado_filtrado["ESTADO"] == filtro_estado]

    st.markdown("## 📋 ANÁLISIS DETALLADO")
    
    tab1, tab2, tab3 = st.tabs(["🔝 TOP Deudores", "📊 Por Periodo", "📄 Detalle"])

    with tab1:
        pendientes = resultado_filtrado[resultado_filtrado["ESTADO"] == "⏳ PENDIENTE"].copy()
        if len(pendientes) > 0:
            top_20 = pendientes.nlargest(20, "SALDO_PENDIENTE")
            st.dataframe(top_20[["ID_COBRANZA", "PERIODO", "TIPO", "DEUDA", "TOTAL_PAGADO", "SALDO_PENDIENTE"]], use_container_width=True)
            st.metric("💰 Saldo TOP 20", f"Bs. {top_20['SALDO_PENDIENTE'].sum():,.2f}")
        else:
            st.info("✅ No hay pendientes")

    with tab2:
        resumen = resultado_filtrado.groupby("PERIODO").agg({
            "ID_COBRANZA": "count",
            "DEUDA": "sum",
            "TOTAL_PAGADO": "sum",
            "SALDO_PENDIENTE": "sum"
        }).reset_index()
        resumen.columns = ["PERIODO", "CASOS", "DEUDA", "PAGADO", "PENDIENTE"]
        resumen["EFECTIVIDAD_%"] = (resumen["PAGADO"] / resumen["DEUDA"] * 100).round(1)
        st.dataframe(resumen, use_container_width=True)

    with tab3:
        st.dataframe(resultado_filtrado[["ID_COBRANZA", "PERIODO", "TIPO", "DEUDA", "TOTAL_PAGADO", "SALDO_PENDIENTE", "ESTADO"]], use_container_width=True)


def modulo_graficos():
    st.markdown('<div class="main-header">📈 GRÁFICOS INTERACTIVOS AVANZADOS</div>', unsafe_allow_html=True)

    # Verificar si hay datos cargados
    if "resultado_cruce" not in st.session_state:
        st.warning("⚠️ **No hay datos cargados**")
        st.info("👉 Ve al módulo **'Dashboard Cruce Deuda vs Pagos'** y carga tus archivos primero.")
        return

    resultado = st.session_state.resultado_cruce

    st.success(f"✅ Analizando {len(resultado):,} casos de cobranza")
    
    # Calcular métricas
    total_cartera = resultado["DEUDA"].sum()
    total_recuperado = resultado["TOTAL_PAGADO"].sum()
    saldo_pendiente = resultado["SALDO_PENDIENTE"].sum()
    porcentaje_recuperacion = (total_recuperado / total_cartera * 100) if total_cartera > 0 else 0
    total_casos = len(resultado)
    casos_pagados = len(resultado[resultado["ESTADO"] == "✅ PAGADO"])
    casos_pendientes = len(resultado[resultado["ESTADO"] == "⏳ PENDIENTE"])

    # Métricas principales
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("💼 Cartera Total", f"Bs. {total_cartera:,.2f}")
    with col2:
        st.metric("✅ Recuperado", f"Bs. {total_recuperado:,.2f}")
    with col3:
        st.metric("⏳ Pendiente", f"Bs. {saldo_pendiente:,.2f}")
    with col4:
        st.metric("📊 Efectividad", f"{porcentaje_recuperacion:.1f}%")

    st.markdown("---")

    # GRÁFICO 1: Comparativa Recuperado vs Pendiente
    st.markdown("## 💰 Comparativa: Recuperado vs Pendiente")
    
    fig_comparativa = go.Figure()
    
    fig_comparativa.add_trace(go.Bar(
        name='Recuperado',
        x=['Monto Total'],
        y=[total_recuperado],
        marker_color='#28a745',
        text=[f'Bs. {total_recuperado:,.2f}'],
        textposition='auto',
        hovertemplate='<b>Recuperado</b><br>Bs. %{y:,.2f}<extra></extra>'
    ))
    
    fig_comparativa.add_trace(go.Bar(
        name='Pendiente',
        x=['Monto Total'],
        y=[saldo_pendiente],
        marker_color='#dc3545',
        text=[f'Bs. {saldo_pendiente:,.2f}'],
        textposition='auto',
        hovertemplate='<b>Pendiente</b><br>Bs. %{y:,.2f}<extra></extra>'
    ))
    
    fig_comparativa.update_layout(
        barmode='group',
        height=400,
        showlegend=True,
        hovermode='x unified',
        title_text="Comparación de Montos",
        title_font_size=16
    )
    
    st.plotly_chart(fig_comparativa, use_container_width=True)

    st.markdown("---")

    # GRÁFICO 2 y 3: Distribución en dos columnas
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 🎯 Distribución de Casos")
        
        fig_pie = go.Figure(data=[go.Pie(
            labels=['Pagado', 'Pendiente'],
            values=[casos_pagados, casos_pendientes],
            marker=dict(colors=['#28a745', '#ffc107']),
            hole=0.4,
            textinfo='label+percent+value',
            hovertemplate='<b>%{label}</b><br>Casos: %{value}<br>Porcentaje: %{percent}<extra></extra>'
        )])
        
        fig_pie.update_layout(
            height=400,
            showlegend=True,
            annotations=[dict(text=f'{total_casos}<br>Total', x=0.5, y=0.5, font_size=20, showarrow=False)]
        )
        
        st.plotly_chart(fig_pie, use_container_width=True)

    with col2:
        st.markdown("### 💵 Distribución de Montos")
        
        fig_pie_montos = go.Figure(data=[go.Pie(
            labels=['Recuperado', 'Pendiente'],
            values=[total_recuperado, saldo_pendiente],
            marker=dict(colors=['#28a745', '#dc3545']),
            hole=0.4,
            textinfo='label+percent',
            hovertemplate='<b>%{label}</b><br>Bs. %{value:,.2f}<br>%{percent}<extra></extra>'
        )])
        
        fig_pie_montos.update_layout(
            height=400,
            showlegend=True,
            annotations=[dict(text=f'Bs. {total_cartera:,.0f}<br>Total', x=0.5, y=0.5, font_size=16, showarrow=False)]
        )
        
        st.plotly_chart(fig_pie_montos, use_container_width=True)

    st.markdown("---")

    # GRÁFICO 4: Evolución por Periodo
    st.markdown("## 📅 Evolución por Periodo")
    
    periodo_analisis = resultado.groupby("PERIODO").agg({
        "DEUDA": "sum",
        "TOTAL_PAGADO": "sum",
        "SALDO_PENDIENTE": "sum"
    }).reset_index()
    
    fig_periodo = go.Figure()
    
    fig_periodo.add_trace(go.Bar(
        name='Deuda Total',
        x=periodo_analisis['PERIODO'],
        y=periodo_analisis['DEUDA'],
        marker_color='#667eea',
        hovertemplate='<b>Deuda</b><br>Bs. %{y:,.2f}<extra></extra>'
    ))
    
    fig_periodo.add_trace(go.Bar(
        name='Pagado',
        x=periodo_analisis['PERIODO'],
        y=periodo_analisis['TOTAL_PAGADO'],
        marker_color='#28a745',
        hovertemplate='<b>Pagado</b><br>Bs. %{y:,.2f}<extra></extra>'
    ))
    
    fig_periodo.add_trace(go.Bar(
        name='Pendiente',
        x=periodo_analisis['PERIODO'],
        y=periodo_analisis['SALDO_PENDIENTE'],
        marker_color='#ffc107',
        hovertemplate='<b>Pendiente</b><br>Bs. %{y:,.2f}<extra></extra>'
    ))
    
    fig_periodo.update_layout(
        barmode='group',
        height=450,
        xaxis_title="Periodo",
        yaxis_title="Monto (Bs.)",
        hovermode='x unified',
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    
    st.plotly_chart(fig_periodo, use_container_width=True)

    st.markdown("---")

    # GRÁFICO 5: Distribución por Tipo de Deuda
    st.markdown("## 🏷️ Distribución por Tipo de Deuda")
    
    tipo_analisis = resultado.groupby("TIPO").agg({
        "DEUDA": "sum",
        "TOTAL_PAGADO": "sum"
    }).reset_index()
    
    tipo_analisis["Pendiente"] = tipo_analisis["DEUDA"] - tipo_analisis["TOTAL_PAGADO"]
    
    fig_tipo = go.Figure()
    
    fig_tipo.add_trace(go.Bar(
        name='Recuperado',
        x=tipo_analisis['TIPO'],
        y=tipo_analisis['TOTAL_PAGADO'],
        marker_color='#28a745',
        hovertemplate='<b>Recuperado</b><br>Bs. %{y:,.2f}<extra></extra>'
    ))
    
    fig_tipo.add_trace(go.Bar(
        name='Pendiente',
        x=tipo_analisis['TIPO'],
        y=tipo_analisis['Pendiente'],
        marker_color='#ffc107',
        hovertemplate='<b>Pendiente</b><br>Bs. %{y:,.2f}<extra></extra>'
    ))
    
    fig_tipo.update_layout(
        barmode='stack',
        height=450,
        xaxis_title="Tipo de Deuda",
        yaxis_title="Monto (Bs.)",
        hovermode='x unified',
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    
    st.plotly_chart(fig_tipo, use_container_width=True)

    st.markdown("---")

    # GRÁFICO 6: Efectividad por Periodo
    st.markdown("## 🎯 Efectividad de Recuperación por Periodo")
    
    efectividad_periodo = resultado.groupby("PERIODO").apply(
        lambda x: (x["TOTAL_PAGADO"].sum() / x["DEUDA"].sum() * 100) if x["DEUDA"].sum() > 0 else 0
    ).reset_index()
    efectividad_periodo.columns = ["PERIODO", "EFECTIVIDAD"]
    
    fig_efectividad = go.Figure()
    
    fig_efectividad.add_trace(go.Scatter(
        x=efectividad_periodo['PERIODO'],
        y=efectividad_periodo['EFECTIVIDAD'],
        mode='lines+markers+text',
        line=dict(color='#667eea', width=3),
        marker=dict(size=12, color='#764ba2'),
        text=[f'{val:.1f}%' for val in efectividad_periodo['EFECTIVIDAD']],
        textposition='top center',
        hovertemplate='<b>%{x}</b><br>Efectividad: %{y:.1f}%<extra></extra>'
    ))
    
    fig_efectividad.add_hline(
        y=70, 
        line_dash="dash", 
        line_color="green",
        annotation_text="Meta: 70%",
        annotation_position="right"
    )
    
    fig_efectividad.add_hline(
        y=50, 
        line_dash="dot", 
        line_color="orange",
        annotation_text="Umbral: 50%",
        annotation_position="right"
    )
    
    fig_efectividad.update_layout(
        height=400,
        xaxis_title="Periodo",
        yaxis_title="Efectividad (%)",
        yaxis_range=[0, 100],
        hovermode='x unified'
    )
    
    st.plotly_chart(fig_efectividad, use_container_width=True)

    st.markdown("---")

    # GRÁFICO 7: TOP 10 Deudores (Barras Horizontales)
    st.markdown("## 🔝 TOP 10 Deudores con Mayor Saldo Pendiente")
    
    pendientes = resultado[resultado["ESTADO"] == "⏳ PENDIENTE"].copy()
    
    if len(pendientes) > 0:
        top_10 = pendientes.nlargest(10, "SALDO_PENDIENTE")
        
        fig_top = go.Figure(go.Bar(
            x=top_10['SALDO_PENDIENTE'],
            y=top_10['ID_COBRANZA'],
            orientation='h',
            marker=dict(
                color=top_10['SALDO_PENDIENTE'],
                colorscale='Reds',
                showscale=True,
                colorbar=dict(title="Saldo (Bs.)")
            ),
            text=[f'Bs. {val:,.2f}' for val in top_10['SALDO_PENDIENTE']],
            textposition='auto',
            hovertemplate='<b>%{y}</b><br>Saldo: Bs. %{x:,.2f}<extra></extra>'
        ))
        
        fig_top.update_layout(
            height=500,
            xaxis_title="Saldo Pendiente (Bs.)",
            yaxis_title="ID Cobranza",
            yaxis=dict(autorange="reversed")
        )
        
        st.plotly_chart(fig_top, use_container_width=True)
        
        st.metric("💰 Saldo Total TOP 10", f"Bs. {top_10['SALDO_PENDIENTE'].sum():,.2f}")
    else:
        st.info("✅ ¡Excelente! No hay casos pendientes.")

    st.markdown("---")

    # GRÁFICO 8: Heatmap de Efectividad por Periodo y Tipo
    st.markdown("## 🔥 Mapa de Calor: Efectividad por Periodo y Tipo")
    
    pivot_data = resultado.pivot_table(
        values='TOTAL_PAGADO',
        index='TIPO',
        columns='PERIODO',
        aggfunc='sum',
        fill_value=0
    )
    
    pivot_deuda = resultado.pivot_table(
        values='DEUDA',
        index='TIPO',
        columns='PERIODO',
        aggfunc='sum',
        fill_value=0
    )
    
    efectividad_pivot = (pivot_data / pivot_deuda * 100).fillna(0).round(1)
    
    fig_heatmap = go.Figure(data=go.Heatmap(
        z=efectividad_pivot.values,
        x=efectividad_pivot.columns,
        y=efectividad_pivot.index,
        colorscale='RdYlGn',
        text=efectividad_pivot.values,
        texttemplate='%{text:.1f}%',
        textfont={"size": 10},
        colorbar=dict(title="Efectividad %"),
        hovertemplate='<b>Tipo:</b> %{y}<br><b>Periodo:</b> %{x}<br><b>Efectividad:</b> %{z:.1f}%<extra></extra>'
    ))
    
    fig_heatmap.update_layout(
        height=400,
        xaxis_title="Periodo",
        yaxis_title="Tipo de Deuda"
    )
    
    st.plotly_chart(fig_heatmap, use_container_width=True)

    st.markdown("---")
    
    st.info("💡 **Tip:** Todos los gráficos son interactivos. Pasa el mouse para ver detalles, haz zoom, descarga imágenes haciendo clic en el ícono de cámara.")


def modulo_sms():
    st.title("📲 GENERADOR DE SMS")

    def limpiar_columnas(df):
        df.columns = df.columns.str.strip().str.upper().str.replace(" ", "_")
        return df

    archivo_suscriptor = st.file_uploader("📂 BASE SUSCRIPTOR", type=["xlsx"])
    archivo_pagos = st.file_uploader("💵 BASE PAGOS", type=["xlsx"])

    if not archivo_suscriptor or not archivo_pagos:
        retur
