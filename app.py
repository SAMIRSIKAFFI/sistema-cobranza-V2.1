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
    
    if "resultado_cruce" not in st.session_state:
        st.session_state.resultado_cruce = None

    if st.session_state.df_deuda_base is None:
        st.info("🔹 **Paso 1:** Carga la base de CARTERA/DEUDA")
        
        archivo_deuda = st.file_uploader(
            "📂 Subir archivo CARTERA / DEUDA",
            type=["xlsx"],
            help="Debe contener: ID_COBRANZA, PERIODO, DEUDA, TIPO",
            key="uploader_cartera"
        )

        if archivo_deuda:
            with st.spinner("Procesando cartera..."):
                try:
                    df_deuda = pd.read_excel(archivo_deuda)
                    df_deuda = limpiar_columnas(df_deuda)
                    columnas_deuda = {"ID_COBRANZA", "PERIODO", "DEUDA", "TIPO"}

                    if not columnas_deuda.issubset(df_deuda.columns):
                        st.error("❌ El archivo CARTERA no tiene las columnas obligatorias")
                        st.error(f"**Columnas requeridas:** ID_COBRANZA, PERIODO, DEUDA, TIPO")
                        st.error(f"**Columnas encontradas:** {', '.join(df_deuda.columns)}")
                        return

                    df_deuda["ID_COBRANZA"] = df_deuda["ID_COBRANZA"].astype(str)
                    df_deuda["PERIODO"] = df_deuda["PERIODO"].astype(str)
                    df_deuda["DEUDA"] = pd.to_numeric(df_deuda["DEUDA"], errors="coerce").fillna(0)

                    if (df_deuda["DEUDA"] < 0).any():
                        st.warning("⚠️ Montos negativos detectados y corregidos")
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

    df_deuda = st.session_state.df_deuda_base
    
    col1, col2 = st.columns([3, 1])
    with col1:
        st.success("✅ **Cartera base cargada en memoria**")
    with col2:
        if st.button("🔄 Reemplazar", use_container_width=True):
            st.session_state.df_deuda_base = None
            st.session_state.resultado_cruce = None
            st.rerun()

    with st.expander("📊 Ver resumen de Cartera Base"):
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📄 Registros", f"{len(df_deuda):,}")
        with col2:
            st.metric("💰 Cartera Total", f"Bs. {df_deuda['DEUDA'].sum():,.2f}")
        with col3:
            st.metric("📅 Periodos", df_deuda["PERIODO"].nunique())

    st.markdown("---")

    st.info("🔹 **Paso 2:** Carga el archivo de PAGOS para realizar el cruce")
    
    archivo_pagos = st.file_uploader(
        "💵 Subir archivo PAGOS",
        type=["xlsx"],
        help="Debe contener: ID_COBRANZA, PERIODO, IMPORTE",
        key="uploader_pagos"
    )

    if archivo_pagos:
        with st.spinner("Procesando cruce..."):
            try:
                df_pagos = pd.read_excel(archivo_pagos)
                df_pagos = limpiar_columnas(df_pagos)
                
                columnas_pagos = {"ID_COBRANZA", "PERIODO", "IMPORTE"}
                if not columnas_pagos.issubset(df_pagos.columns):
                    st.error("❌ El archivo PAGOS no tiene las columnas obligatorias")
                    st.error(f"**Columnas requeridas:** ID_COBRANZA, PERIODO, IMPORTE")
                    st.error(f"**Columnas encontradas:** {', '.join(df_pagos.columns)}")
                    return

                df_pagos["ID_COBRANZA"] = df_pagos["ID_COBRANZA"].astype(str)
                df_pagos["PERIODO"] = df_pagos["PERIODO"].astype(str)
                df_pagos["IMPORTE"] = pd.to_numeric(df_pagos["IMPORTE"], errors="coerce").fillna(0)

                if (df_pagos["IMPORTE"] < 0).any():
                    st.warning("⚠️ Montos negativos detectados y corregidos")
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

                st.session_state.resultado_cruce = resultado

                st.success("✅ Cruce realizado correctamente")
                
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
                
                with st.expander("🔍 FILTROS Y BÚSQUEDA", expanded=False):
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
                        st.dataframe(top_20[["ID_COBRANZA", "PERIODO", "TIPO", "DEUDA", "TOTAL_PAGADO", "SALDO_PENDIENTE"]], use_container_width=True, height=400)
                        st.metric("💰 Saldo TOP 20", f"Bs. {top_20['SALDO_PENDIENTE'].sum():,.2f}")
                    else:
                        st.info("✅ No hay casos pendientes")

                with tab2:
                    resumen = resultado_filtrado.groupby("PERIODO").agg({
                        "ID_COBRANZA": "count",
                        "DEUDA": "sum",
                        "TOTAL_PAGADO": "sum",
                        "SALDO_PENDIENTE": "sum"
                    }).reset_index()
                    resumen.columns = ["PERIODO", "CASOS", "DEUDA", "PAGADO", "PENDIENTE"]
                    resumen["EFECTIVIDAD_%"] = (resumen["PAGADO"] / resumen["DEUDA"] * 100).round(1)
                    st.dataframe(resumen, use_container_width=True, height=400)

                with tab3:
                    st.dataframe(resultado_filtrado[["ID_COBRANZA", "PERIODO", "TIPO", "DEUDA", "TOTAL_PAGADO", "SALDO_PENDIENTE", "ESTADO"]], use_container_width=True, height=400)
                    st.info(f"📊 Mostrando {len(resultado_filtrado):,} de {len(resultado):,} casos")

            except Exception as e:
                st.error(f"❌ Error: {str(e)}")


def modulo_graficos():
    st.markdown('<div class="main-header">📈 GRÁFICOS INTERACTIVOS AVANZADOS</div>', unsafe_allow_html=True)

    if "resultado_cruce" not in st.session_state or st.session_state.resultado_cruce is None:
        st.warning("⚠️ **No hay datos cargados**")
        st.info("👉 Ve al módulo **'📊 Dashboard Cruce Deuda vs Pagos'** y carga tus archivos primero.")
        
        st.markdown("---")
        st.markdown("### 📋 Pasos para ver los gráficos:")
        st.markdown("""
        1. Haz clic en **'📊 Dashboard Cruce Deuda vs Pagos'** en el menú lateral
        2. Sube tu archivo de **CARTERA**
        3. Sube tu archivo de **PAGOS**
        4. Regresa a este módulo para ver los gráficos interactivos
        """)
        return

    resultado = st.session_state.resultado_cruce

    st.success(f"✅ Analizando {len(resultado):,} casos de cobranza")
    
    total_cartera = resultado["DEUDA"].sum()
    total_recuperado = resultado["TOTAL_PAGADO"].sum()
    saldo_pendiente = resultado["SALDO_PENDIENTE"].sum()
    porcentaje_recuperacion = (total_recuperado / total_cartera * 100) if total_cartera > 0 else 0
    total_casos = len(resultado)
    casos_pagados = len(resultado[resultado["ESTADO"] == "✅ PAGADO"])
    casos_pendientes = len(resultado[resultado["ESTADO"] == "⏳ PENDIENTE"])

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
    fig_comparativa.update_layout(barmode='group', height=400, showlegend=True, hovermode='x unified')
    st.plotly_chart(fig_comparativa, use_container_width=True)

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 🎯 Distribución de Casos")
        fig_pie = go.Figure(data=[go.Pie(
            labels=['Pagado', 'Pendiente'],
            values=[casos_pagados, casos_pendientes],
            marker=dict(colors=['#28a745', '#ffc107']),
            hole=0.4,
            textinfo='label+percent+value',
            hovertemplate='<b>%{label}</b><br>Casos: %{value}<br>%{percent}<extra></extra>'
        )])
        fig_pie.update_layout(height=400, annotations=[dict(text=f'{total_casos}<br>Total', x=0.5, y=0.5, font_size=20, showarrow=False)])
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
        fig_pie_montos.update_layout(height=400, annotations=[dict(text=f'Bs. {total_cartera:,.0f}<br>Total', x=0.5, y=0.5, font_size=16, showarrow=False)])
        st.plotly_chart(fig_pie_montos, use_container_width=True)

    st.markdown("---")

    st.markdown("## 📅 Evolución por Periodo")
    periodo_analisis = resultado.groupby("PERIODO").agg({
        "DEUDA": "sum",
        "TOTAL_PAGADO": "sum",
        "SALDO_PENDIENTE": "sum"
    }).reset_index()
    
    fig_periodo = go.Figure()
    fig_periodo.add_trace(go.Bar(name='Deuda Total', x=periodo_analisis['PERIODO'], y=periodo_analisis['DEUDA'], marker_color='#667eea'))
    fig_periodo.add_trace(go.Bar(name='Pagado', x=periodo_analisis['PERIODO'], y=periodo_analisis['TOTAL_PAGADO'], marker_color='#28a745'))
    fig_periodo.add_trace(go.Bar(name='Pendiente', x=periodo_analisis['PERIODO'], y=periodo_analisis['SALDO_PENDIENTE'], marker_color='#ffc107'))
    fig_periodo.update_layout(barmode='group', height=450, xaxis_title="Periodo", yaxis_title="Monto (Bs.)", hovermode='x unified')
    st.plotly_chart(fig_periodo, use_container_width=True)

    st.markdown("---")

    st.markdown("## 🏷️ Distribución por Tipo de Deuda")
    tipo_analisis = resultado.groupby("TIPO").agg({"DEUDA": "sum", "TOTAL_PAGADO": "sum"}).reset_index()
    tipo_analisis["Pendiente"] = tipo_analisis["DEUDA"] - tipo_analisis["TOTAL_PAGADO"]
    
    fig_tipo = go.Figure()
    fig_tipo.add_trace(go.Bar(name='Recuperado', x=tipo_analisis['TIPO'], y=tipo_analisis['TOTAL_PAGADO'], marker_color='#28a745'))
    fig_tipo.add_trace(go.Bar(name='Pendiente', x=tipo_analisis['TIPO'], y=tipo_analisis['Pendiente'], marker_color='#ffc107'))
    fig_tipo.update_layout(barmode='stack', height=450, xaxis_title="Tipo de Deuda", yaxis_title="Monto (Bs.)", hovermode='x unified')
    st.plotly_chart(fig_tipo, use_container_width=True)

    st.markdown("---")

    st.markdown("## 🎯 Efectividad por Periodo")
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
        textposition='top center'
    ))
    fig_efectividad.add_hline(y=70, line_dash="dash", line_color="green", annotation_text="Meta: 70%")
    fig_efectividad.add_hline(y=50, line_dash="dot", line_color="orange", annotation_text="Umbral: 50%")
    fig_efectividad.update_layout(height=400, xaxis_title="Periodo", yaxis_title="Efectividad (%)", yaxis_range=[0, 100])
    st.plotly_chart(fig_efectividad, use_container_width=True)

    st.markdown("---")

    st.markdown("## 🔝 TOP 10 Deudores")
    pendientes = resultado[resultado["ESTADO"] == "⏳ PENDIENTE"].copy()
    
    if len(pendientes) > 0:
        top_10 = pendientes.nlargest(10, "SALDO_PENDIENTE")
        fig_top = go.Figure(go.Bar(
            x=top_10['SALDO_PENDIENTE'],
            y=top_10['ID_COBRANZA'],
            orientation='h',
            marker=dict(color=top_10['SALDO_PENDIENTE'], colorscale='Reds', showscale=True),
            text=[f'Bs. {val:,.2f}' for val in top_10['SALDO_PENDIENTE']],
            textposition='auto'
        ))
        fig_top.update_layout(height=500, xaxis_title="Saldo (Bs.)", yaxis_title="ID Cobranza", yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig_top, use_container_width=True)
        st.metric("💰 Saldo Total TOP 10", f"Bs. {top_10['SALDO_PENDIENTE'].sum():,.2f}")
    else:
        st.info("✅ No hay casos pendientes")

    st.markdown("---")
    st.info("💡 **Tip:** Pasa el mouse sobre los gráficos para ver detalles. Haz zoom, descarga imágenes con el ícono de cámara.")


def modulo_sms():
    st.markdown('<div class="main-header">📲 GENERADOR PROFESIONAL DE SMS</div>', unsafe_allow_html=True)
    
    # Verificar que exista cartera cargada
    if "df_deuda_base" not in st.session_state or st.session_state.df_deuda_base is None:
        st.warning("⚠️ **No hay CARTERA cargada en el sistema**")
        st.info("👉 Primero debes ir al módulo **'📊 Dashboard Cruce Deuda vs Pagos'** y cargar la CARTERA base.")
        return
    
    df_cartera = st.session_state.df_deuda_base.copy()
    st.success(f"✅ Cartera disponible: {len(df_cartera):,} registros, {df_cartera['ID_COBRANZA'].nunique()} códigos únicos")
    
    st.markdown("---")
    
    # Información de formato
    with st.expander("📋 Formato de archivos requerido", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
            **📂 Archivo SUSCRIPTOR:**
            - `CODIGO` - ID del cliente
            - `NUMERO` - Teléfono para SMS
            - `NOMBRE` - Nombre del cliente
            - `FECHA` - Fecha de referencia
            - `MONTO` - (Opcional, se calcula de cartera)
            """)
        with col2:
            st.markdown("""
            **💵 Archivo PAGOS:**
            - `ID_COBRANZA` o `CODIGO` - ID del cliente
            - `PERIODO` - Periodo pagado
            - `IMPORTE` - Monto pagado
            """)
    
    st.markdown("---")
    
    # Función para limpiar columnas
    def limpiar_columnas(df):
        df.columns = df.columns.str.strip().str.upper().str.replace(" ", "_")
        return df
    
    # PASO 1: Cargar BASE SUSCRIPTOR
    st.markdown("### 📂 PASO 1: Cargar BASE SUSCRIPTOR")
    archivo_suscriptor = st.file_uploader(
        "Subir archivo SUSCRIPTOR",
        type=["xlsx"],
        key="sms_suscriptor",
        help="Debe contener: CODIGO, NUMERO, NOMBRE, FECHA"
    )
    
    if not archivo_suscriptor:
        st.info("⬆️ Sube el archivo de suscriptores para continuar")
        return
    
    try:
        df_suscriptor = pd.read_excel(archivo_suscriptor)
        df_suscriptor = limpiar_columnas(df_suscriptor)
        
        # Validar columnas SUSCRIPTOR
        columnas_suscriptor = {"CODIGO", "NUMERO", "NOMBRE", "FECHA"}
        if not columnas_suscriptor.issubset(df_suscriptor.columns):
            st.error("❌ El archivo SUSCRIPTOR no tiene las columnas obligatorias")
            st.error(f"**Requeridas:** CODIGO, NUMERO, NOMBRE, FECHA")
            st.error(f"**Encontradas:** {', '.join(df_suscriptor.columns)}")
            return
        
        df_suscriptor["CODIGO"] = df_suscriptor["CODIGO"].astype(str)
        st.success(f"✅ Suscriptores cargados: {len(df_suscriptor):,} registros")
        
    except Exception as e:
        st.error(f"❌ Error al cargar SUSCRIPTOR: {str(e)}")
        return
    
    st.markdown("---")
    
    # PASO 2: Cargar BASE PAGOS
    st.markdown("### 💵 PASO 2: Cargar BASE PAGOS")
    archivo_pagos = st.file_uploader(
        "Subir archivo PAGOS",
        type=["xlsx"],
        key="sms_pagos",
        help="Debe contener: ID_COBRANZA (o CODIGO), PERIODO, IMPORTE"
    )
    
    if not archivo_pagos:
        st.info("⬆️ Sube el archivo de pagos para continuar")
        return
    
    try:
        df_pagos = pd.read_excel(archivo_pagos)
        df_pagos = limpiar_columnas(df_pagos)
        
        # Validar columnas PAGOS (acepta ID_COBRANZA o CODIGO)
        if "ID_COBRANZA" in df_pagos.columns:
            df_pagos = df_pagos.rename(columns={"ID_COBRANZA": "CODIGO"})
        
        columnas_pagos = {"CODIGO", "PERIODO", "IMPORTE"}
        if not columnas_pagos.issubset(df_pagos.columns):
            st.error("❌ El archivo PAGOS no tiene las columnas obligatorias")
            st.error(f"**Requeridas:** ID_COBRANZA o CODIGO, PERIODO, IMPORTE")
            st.error(f"**Encontradas:** {', '.join(df_pagos.columns)}")
            return
        
        df_pagos["CODIGO"] = df_pagos["CODIGO"].astype(str)
        df_pagos["PERIODO"] = df_pagos["PERIODO"].astype(str)
        df_pagos["IMPORTE"] = pd.to_numeric(df_pagos["IMPORTE"], errors="coerce").fillna(0)
        
        st.success(f"✅ Pagos cargados: {len(df_pagos):,} registros")
        
    except Exception as e:
        st.error(f"❌ Error al cargar PAGOS: {str(e)}")
        return
    
    st.markdown("---")
    
    # PASO 3: REALIZAR CRUCE
    st.markdown("### 🔗 PASO 3: Cruce de Datos")
    
    with st.spinner("Procesando cruce con CARTERA..."):
        try:
            # Preparar cartera
            df_cartera["CODIGO"] = df_cartera["ID_COBRANZA"].astype(str)
            df_cartera["PERIODO"] = df_cartera["PERIODO"].astype(str)
            
            # Calcular periodos totales por código
            periodos_totales = df_cartera.groupby("CODIGO").agg({
                "PERIODO": "count",
                "DEUDA": "sum",
                "TIPO": lambda x: x.iloc[0] if len(x) > 0 else "SIN TIPO"
            }).reset_index()
            periodos_totales.columns = ["CODIGO", "PERIODOS_TOTALES", "DEUDA_TOTAL", "TIPO"]
            
            # Calcular periodos pagados por código
            periodos_pagados = df_pagos.groupby("CODIGO").agg({
                "PERIODO": "count",
                "IMPORTE": "sum"
            }).reset_index()
            periodos_pagados.columns = ["CODIGO", "PERIODOS_PAGADOS", "TOTAL_PAGADO"]
            
            # Merge con suscriptor
            df_analisis = df_suscriptor.merge(periodos_totales, on="CODIGO", how="left")
            df_analisis = df_analisis.merge(periodos_pagados, on="CODIGO", how="left")
            
            # Rellenar NaN
            df_analisis["PERIODOS_TOTALES"] = df_analisis["PERIODOS_TOTALES"].fillna(0).astype(int)
            df_analisis["PERIODOS_PAGADOS"] = df_analisis["PERIODOS_PAGADOS"].fillna(0).astype(int)
            df_analisis["DEUDA_TOTAL"] = df_analisis["DEUDA_TOTAL"].fillna(0)
            df_analisis["TOTAL_PAGADO"] = df_analisis["TOTAL_PAGADO"].fillna(0)
            df_analisis["TIPO"] = df_analisis["TIPO"].fillna("SIN TIPO")
            
            # Calcular estado
            df_analisis["PERIODOS_PENDIENTES"] = df_analisis["PERIODOS_TOTALES"] - df_analisis["PERIODOS_PAGADOS"]
            df_analisis["SALDO_PENDIENTE"] = df_analisis["DEUDA_TOTAL"] - df_analisis["TOTAL_PAGADO"]
            df_analisis["SALDO_PENDIENTE"] = df_analisis["SALDO_PENDIENTE"].apply(lambda x: max(0, x))
            
            # Clasificar clientes
            def clasificar_cliente(row):
                if row["PERIODOS_TOTALES"] == 0:
                    return "SIN DEUDA"
                elif row["SALDO_PENDIENTE"] == 0:
                    return "PAGO TOTAL"
                elif row["PERIODOS_PAGADOS"] == 0:
                    return "MOROSO TOTAL"
                else:
                    return "PAGADOR PARCIAL"
            
            df_analisis["CLASIFICACION"] = df_analisis.apply(clasificar_cliente, axis=1)
            
            st.success("✅ Cruce realizado correctamente")
            
            # Mostrar resumen
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("👥 Total Clientes", f"{len(df_analisis):,}")
            with col2:
                morosos = len(df_analisis[df_analisis["CLASIFICACION"] == "MOROSO TOTAL"])
                st.metric("🔴 Morosos Totales", f"{morosos:,}")
            with col3:
                parciales = len(df_analisis[df_analisis["CLASIFICACION"] == "PAGADOR PARCIAL"])
                st.metric("🟡 Pagadores Parciales", f"{parciales:,}")
            with col4:
                sin_deuda = len(df_analisis[df_analisis["CLASIFICACION"].isin(["PAGO TOTAL", "SIN DEUDA"])])
                st.metric("🟢 Al Día", f"{sin_deuda:,}")
            
        except Exception as e:
            st.error(f"❌ Error en el cruce: {str(e)}")
            return
    
    st.markdown("---")
    
    # Vista previa de datos
    with st.expander("👁️ Vista previa de análisis", expanded=False):
        st.dataframe(
            df_analisis[["CODIGO", "NOMBRE", "NUMERO", "TIPO", "PERIODOS_TOTALES", "PERIODOS_PAGADOS", "PERIODOS_PENDIENTES", "SALDO_PENDIENTE", "CLASIFICACION"]],
            use_container_width=True,
            height=300
        )
    
    st.markdown("---")
    
    # PASO 4: OPCIONES DE GENERACIÓN
    st.markdown("### 🎯 PASO 4: Opciones de Generación de SMS")
    
    st.info("💡 **Selecciona UNA opción de las siguientes:**")
    
    opcion = st.radio(
        "Modo de generación:",
        [
            "📊 Opción 1: Generar por TIPO de deuda",
            "🔴 Opción 2: Solo MOROSOS TOTALES (0 pagos)",
            "🟡 Opción 3: Solo PAGADORES PARCIALES (pagaron algo, deben algo)"
        ],
        index=0
    )
    
    st.markdown("---")
    
    # Configuración adicional
    col1, col2 = st.columns(2)
    with col1:
        num_archivos = st.number_input(
            "Dividir en cuántos archivos CSV",
            min_value=1,
            max_value=20,
            value=1,
            help="Si tienes muchos registros, puedes dividirlos en varios archivos"
        )
    with col2:
        prefijo = st.text_input(
            "Prefijo de archivos",
            value="SMS",
            help="Nombre base para los archivos generados"
        )
    
    st.markdown("---")
    
    # Botón para generar
    if st.button("🎯 GENERAR ARCHIVOS CSV", type="primary", use_container_width=True):
        
        # Filtrar según opción seleccionada
        if "Opción 1" in opcion:
            # Generar por TIPO
            st.markdown("### 📥 Archivos generados por TIPO:")
            
            # Filtrar solo los que tienen saldo pendiente
            df_export = df_analisis[df_analisis["SALDO_PENDIENTE"] > 0].copy()
            
            if len(df_export) == 0:
                st.warning("⚠️ No hay clientes con saldo pendiente")
                return
            
            tipos_unicos = df_export["TIPO"].unique()
            
            for tipo in tipos_unicos:
                df_tipo = df_export[df_export["TIPO"] == tipo].copy()
                
                if len(df_tipo) == 0:
                    continue
                
                # Preparar columnas para SMS
                df_csv = df_tipo[["NUMERO", "NOMBRE", "FECHA", "CODIGO", "SALDO_PENDIENTE"]].copy()
                df_csv = df_csv.rename(columns={"SALDO_PENDIENTE": "MONTO"})
                
                # Dividir en partes si es necesario
                tamaño = len(df_csv) // num_archivos + 1
                
                for i in range(num_archivos):
                    inicio = i * tamaño
                    fin = inicio + tamaño
                    df_parte = df_csv.iloc[inicio:fin]
                    
                    if df_parte.empty:
                        continue
                    
                    csv = df_parte.to_csv(index=False, sep=";", encoding="utf-8-sig")
                    
                    nombre_archivo = f"{prefijo}_{tipo}_{i+1}.csv" if num_archivos > 1 else f"{prefijo}_{tipo}.csv"
                    
                    st.download_button(
                        label=f"⬇️ {nombre_archivo} ({len(df_parte)} registros)",
                        data=csv,
                        file_name=nombre_archivo,
                        mime="text/csv",
                        key=f"download_tipo_{tipo}_{i}",
                        use_container_width=True
                    )
                
                st.success(f"✅ Tipo '{tipo}': {len(df_tipo):,} registros")
        
        elif "Opción 2" in opcion:
            # Solo MOROSOS TOTALES
            st.markdown("### 📥 Archivos generados: MOROSOS TOTALES")
            
            df_export = df_analisis[df_analisis["CLASIFICACION"] == "MOROSO TOTAL"].copy()
            
            if len(df_export) == 0:
                st.warning("⚠️ No hay morosos totales (todos han pagado al menos 1 periodo)")
                return
            
            st.info(f"📊 Se generarán archivos con {len(df_export):,} clientes que NO han pagado NINGÚN periodo")
            
            # Preparar columnas para SMS
            df_csv = df_export[["NUMERO", "NOMBRE", "FECHA", "CODIGO", "SALDO_PENDIENTE"]].copy()
            df_csv = df_csv.rename(columns={"SALDO_PENDIENTE": "MONTO"})
            
            # Dividir en partes
            tamaño = len(df_csv) // num_archivos + 1
            
            for i in range(num_archivos):
                inicio = i * tamaño
                fin = inicio + tamaño
                df_parte = df_csv.iloc[inicio:fin]
                
                if df_parte.empty:
                    continue
                
                csv = df_parte.to_csv(index=False, sep=";", encoding="utf-8-sig")
                
                nombre_archivo = f"{prefijo}_MOROSOS_{i+1}.csv" if num_archivos > 1 else f"{prefijo}_MOROSOS.csv"
                
                st.download_button(
                    label=f"⬇️ {nombre_archivo} ({len(df_parte)} registros)",
                    data=csv,
                    file_name=nombre_archivo,
                    mime="text/csv",
                    key=f"download_morosos_{i}",
                    use_container_width=True
                )
            
            st.success(f"✅ {len(df_export):,} morosos totales exportados")
        
        elif "Opción 3" in opcion:
            # Solo PAGADORES PARCIALES
            st.markdown("### 📥 Archivos generados: PAGADORES PARCIALES")
            
            df_export = df_analisis[df_analisis["CLASIFICACION"] == "PAGADOR PARCIAL"].copy()
            
            if len(df_export) == 0:
                st.warning("⚠️ No hay pagadores parciales")
                return
            
            st.info(f"📊 Se generarán archivos con {len(df_export):,} clientes que pagaron AL MENOS 1 periodo pero AÚN deben")
            
            # Preparar columnas para SMS
            df_csv = df_export[["NUMERO", "NOMBRE", "FECHA", "CODIGO", "SALDO_PENDIENTE"]].copy()
            df_csv = df_csv.rename(columns={"SALDO_PENDIENTE": "MONTO"})
            
            # Dividir en partes
            tamaño = len(df_csv) // num_archivos + 1
            
            for i in range(num_archivos):
                inicio = i * tamaño
                fin = inicio + tamaño
                df_parte = df_csv.iloc[inicio:fin]
                
                if df_parte.empty:
                    continue
                
                csv = df_parte.to_csv(index=False, sep=";", encoding="utf-8-sig")
                
                nombre_archivo = f"{prefijo}_PARCIALES_{i+1}.csv" if num_archivos > 1 else f"{prefijo}_PARCIALES.csv"
                
                st.download_button(
                    label=f"⬇️ {nombre_archivo} ({len(df_parte)} registros)",
                    data=csv,
                    file_name=nombre_archivo,
                    mime="text/csv",
                    key=f"download_parciales_{i}",
                    use_container_width=True
                )
            
            st.success(f"✅ {len(df_export):,} pagadores parciales exportados")


if menu == "📊 Dashboard Cruce Deuda vs Pagos":
    modulo_cruce()
elif menu == "📈 Gráficos Interactivos":
    modulo_graficos()
elif menu == "📲 GENERADOR DE SMS":
    modulo_sms()
elif menu == "🚧 Módulo Histórico (En Desarrollo)":
    st.title("📈 Módulo Histórico")
    st.info("🚧 Este módulo está en desarrollo. Próximamente podrás ver análisis históricos acumulados.")
