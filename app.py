import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta

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
    .tipo-box {
        background-color: #e8f4f8;
        padding: 1rem;
        border-radius: 8px;
        border: 2px solid #667eea;
        margin: 0.5rem 0;
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
        "📈 Control Diario y Objetivos"
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
    st.markdown('<div class="main-header">📲 GENERADOR DE SMS - CLIENTE VIVA</div>', unsafe_allow_html=True)
    
    if "df_deuda_base" not in st.session_state or st.session_state.df_deuda_base is None:
        st.warning("⚠️ **No hay CARTERA cargada en el sistema**")
        st.info("👉 Primero debes ir al módulo **'📊 Dashboard Cruce Deuda vs Pagos'** y cargar la CARTERA base.")
        return
    
    df_cartera = st.session_state.df_deuda_base.copy()
    
    st.success(f"✅ Cartera VIVA disponible: {len(df_cartera):,} registros | {df_cartera['ID_COBRANZA'].nunique()} códigos | {df_cartera['TIPO'].nunique()} tipos")
    
    st.markdown("---")
    
    def limpiar_columnas(df):
        df.columns = df.columns.str.strip().str.upper().str.replace(" ", "_")
        return df
    
    st.markdown("### 🎯 PASO 1: Seleccionar TIPOS de Cartera para la Campaña")
    
    tipos_disponibles = sorted(df_cartera["TIPO"].unique().tolist())
    tipo_conteo = df_cartera.groupby("TIPO").size().to_dict()
    
    st.markdown('<div class="tipo-box">', unsafe_allow_html=True)
    seleccionar_todos = st.checkbox(
        "✅ SELECCIONAR TODOS LOS TIPOS",
        value=False,
        help="Marca esta opción para incluir todos los tipos en la campaña"
    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    if seleccionar_todos:
        tipos_seleccionados = tipos_disponibles
        st.success(f"✅ **TODOS LOS TIPOS SELECCIONADOS** ({len(tipos_seleccionados)} tipos)")
        st.markdown("**📊 Resumen de tipos incluidos:**")
        resumen_data = []
        for tipo in tipos_seleccionados:
            conteo = tipo_conteo.get(tipo, 0)
            resumen_data.append({"TIPO": tipo, "REGISTROS": f"{conteo:,}"})
        df_resumen = pd.DataFrame(resumen_data)
        st.dataframe(df_resumen, use_container_width=True, hide_index=True)
    else:
        st.markdown("**📋 Selecciona los tipos que deseas incluir en la campaña:**")
        st.markdown('<div class="tipo-box">', unsafe_allow_html=True)
        tipos_seleccionados = []
        cols = st.columns(2)
        for idx, tipo in enumerate(tipos_disponibles):
            col = cols[idx % 2]
            conteo = tipo_conteo.get(tipo, 0)
            with col:
                if st.checkbox(f"☑️ **{tipo}** ({conteo:,} registros)", value=False, key=f"tipo_{tipo}"):
                    tipos_seleccionados.append(tipo)
        st.markdown('</div>', unsafe_allow_html=True)
        if tipos_seleccionados:
            st.success(f"✅ **{len(tipos_seleccionados)} tipo(s) seleccionado(s):** {', '.join(tipos_seleccionados)}")
        else:
            st.warning("⚠️ **No has seleccionado ningún tipo**")
    
    if not tipos_seleccionados:
        st.error("❌ **Debes seleccionar al menos UN tipo para continuar**")
        st.info("💡 Marca la casilla de un tipo específico o selecciona TODOS")
        return
    
    df_cartera_filtrada = df_cartera[df_cartera["TIPO"].isin(tipos_seleccionados)].copy()
    
    st.markdown("---")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("📄 Registros Filtrados", f"{len(df_cartera_filtrada):,}")
    with col2:
        st.metric("👤 Códigos Únicos", f"{df_cartera_filtrada['ID_COBRANZA'].nunique():,}")
    with col3:
        st.metric("💰 Deuda Total", f"Bs. {df_cartera_filtrada['DEUDA'].sum():,.2f}")
    
    st.markdown("---")
    
    st.markdown("### 📂 PASO 2: Cargar BASE SUSCRIPTOR")
    archivo_suscriptor = st.file_uploader(
        "Subir archivo SUSCRIPTOR (NUMERO, NOMBRE, FECHA, CODIGO)",
        type=["xlsx"],
        key="sms_suscriptor"
    )
    
    if not archivo_suscriptor:
        st.info("⬆️ Sube el archivo de suscriptores para continuar")
        return
    
    try:
        df_suscriptor = pd.read_excel(archivo_suscriptor)
        df_suscriptor = limpiar_columnas(df_suscriptor)
        columnas_suscriptor = {"CODIGO", "NUMERO", "NOMBRE", "FECHA"}
        if not columnas_suscriptor.issubset(df_suscriptor.columns):
            st.error("❌ Columnas faltantes en SUSCRIPTOR")
            st.error(f"**Requeridas:** CODIGO, NUMERO, NOMBRE, FECHA")
            st.error(f"**Encontradas:** {', '.join(df_suscriptor.columns)}")
            return
        df_suscriptor["CODIGO"] = df_suscriptor["CODIGO"].astype(str)
        st.success(f"✅ Suscriptores: {len(df_suscriptor):,} registros")
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        return
    
    st.markdown("---")
    
    st.markdown("### 💵 PASO 3: Cargar BASE PAGOS")
    archivo_pagos = st.file_uploader(
        "Subir archivo PAGOS (CODIGO, PERIODO, IMPORTE)",
        type=["xlsx"],
        key="sms_pagos"
    )
    
    if not archivo_pagos:
        st.info("⬆️ Sube el archivo de pagos para continuar")
        return
    
    try:
        df_pagos = pd.read_excel(archivo_pagos)
        df_pagos = limpiar_columnas(df_pagos)
        if "ID_COBRANZA" in df_pagos.columns:
            df_pagos = df_pagos.rename(columns={"ID_COBRANZA": "CODIGO"})
        columnas_pagos = {"CODIGO", "PERIODO", "IMPORTE"}
        if not columnas_pagos.issubset(df_pagos.columns):
            st.error("❌ Columnas faltantes en PAGOS")
            st.error(f"**Requeridas:** CODIGO, PERIODO, IMPORTE")
            st.error(f"**Encontradas:** {', '.join(df_pagos.columns)}")
            return
        df_pagos["CODIGO"] = df_pagos["CODIGO"].astype(str)
        df_pagos["PERIODO"] = df_pagos["PERIODO"].astype(str)
        df_pagos["IMPORTE"] = pd.to_numeric(df_pagos["IMPORTE"], errors="coerce").fillna(0)
        st.success(f"✅ Pagos: {len(df_pagos):,} registros")
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        return
    
    st.markdown("---")
    
    st.markdown("### 🔗 PASO 4: Cruce y Depuración Automática")
    
    with st.spinner("Procesando cruce con cartera VIVA..."):
        try:
            df_cartera_filtrada["CODIGO"] = df_cartera_filtrada["ID_COBRANZA"].astype(str)
            df_cartera_filtrada["PERIODO"] = df_cartera_filtrada["PERIODO"].astype(str)
            periodos_totales = df_cartera_filtrada.groupby("CODIGO")["PERIODO"].count().reset_index()
            periodos_totales.columns = ["CODIGO", "PERIODOS_TOTALES"]
            deuda_total = df_cartera_filtrada.groupby("CODIGO")["DEUDA"].sum().reset_index()
            deuda_total.columns = ["CODIGO", "DEUDA_TOTAL"]
            periodos_pagados = df_pagos.groupby("CODIGO")["PERIODO"].count().reset_index()
            periodos_pagados.columns = ["CODIGO", "PERIODOS_PAGADOS"]
            total_pagado = df_pagos.groupby("CODIGO")["IMPORTE"].sum().reset_index()
            total_pagado.columns = ["CODIGO", "TOTAL_PAGADO"]
            df_analisis = df_suscriptor.copy()
            df_analisis = df_analisis.merge(periodos_totales, on="CODIGO", how="left")
            df_analisis = df_analisis.merge(deuda_total, on="CODIGO", how="left")
            df_analisis = df_analisis.merge(periodos_pagados, on="CODIGO", how="left")
            df_analisis = df_analisis.merge(total_pagado, on="CODIGO", how="left")
            df_analisis["PERIODOS_TOTALES"] = df_analisis["PERIODOS_TOTALES"].fillna(0).astype(int)
            df_analisis["PERIODOS_PAGADOS"] = df_analisis["PERIODOS_PAGADOS"].fillna(0).astype(int)
            df_analisis["DEUDA_TOTAL"] = df_analisis["DEUDA_TOTAL"].fillna(0)
            df_analisis["TOTAL_PAGADO"] = df_analisis["TOTAL_PAGADO"].fillna(0)
            df_analisis["PERIODOS_PENDIENTES"] = df_analisis["PERIODOS_TOTALES"] - df_analisis["PERIODOS_PAGADOS"]
            df_analisis["SALDO_PENDIENTE"] = df_analisis["DEUDA_TOTAL"] - df_analisis["TOTAL_PAGADO"]
            df_analisis["SALDO_PENDIENTE"] = df_analisis["SALDO_PENDIENTE"].apply(lambda x: max(0, x))
            df_analisis_depurado = df_analisis[df_analisis["PERIODOS_PENDIENTES"] > 0].copy()
            eliminados_pago_total = len(df_analisis) - len(df_analisis_depurado)
            st.success("✅ Cruce realizado y pagos totales depurados automáticamente")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("👥 Total Inicial", f"{len(df_analisis):,}")
            with col2:
                st.metric("❌ Pagos Totales (eliminados)", f"{eliminados_pago_total:,}")
            with col3:
                st.metric("✅ Con Saldo Pendiente", f"{len(df_analisis_depurado):,}")
        except Exception as e:
            st.error(f"❌ Error en cruce: {str(e)}")
            return
    
    if len(df_analisis_depurado) == 0:
        st.warning("⚠️ No hay clientes con saldo pendiente después de depurar pagos totales")
        return
    
    st.markdown("---")
    
    with st.expander("👁️ Vista previa de datos procesados"):
        st.dataframe(
            df_analisis_depurado[["CODIGO", "NOMBRE", "NUMERO", "PERIODOS_TOTALES", "PERIODOS_PAGADOS", "PERIODOS_PENDIENTES", "SALDO_PENDIENTE"]].head(20),
            use_container_width=True
        )
    
    st.markdown("---")
    
    st.markdown("### 🎯 PASO 5: Configurar Campaña SMS")
    
    st.info("💡 Los pagos totales ya fueron depurados automáticamente. Ahora elige el tipo de campaña:")
    
    opcion_campana = st.radio(
        "Tipo de campaña:",
        [
            "🔴 CAMPAÑA AGRESIVA: Solo morosos totales (0 pagos realizados)",
            "🟡 CAMPAÑA GENERAL: Todos con al menos 1 periodo pendiente"
        ],
        index=1,
        help="Agresiva = solo quienes NO pagaron nada | General = todos con al menos 1 pendiente (incluye morosos totales + pagadores parciales)"
    )
    
    if "AGRESIVA" in opcion_campana:
        df_campana = df_analisis_depurado[df_analisis_depurado["PERIODOS_PAGADOS"] == 0].copy()
        tipo_campana = "MOROSOS"
    else:
        df_campana = df_analisis_depurado.copy()
        tipo_campana = "GENERAL"
    
    if len(df_campana) == 0:
        st.warning(f"⚠️ No hay clientes para esta campaña")
        return
    
    st.success(f"✅ Clientes para campaña {tipo_campana}: {len(df_campana):,}")
    
    st.markdown("---")
    
    st.markdown("### ⚙️ Configuración de Archivos")
    
    col1, col2 = st.columns(2)
    with col1:
        num_archivos = st.number_input(
            "Dividir en cuántos archivos CSV",
            min_value=1,
            max_value=50,
            value=1,
            help="Para campañas grandes, dividir en varios archivos"
        )
    with col2:
        prefijo = st.text_input(
            "Prefijo de archivos",
            value=f"SMS_VIVA_{tipo_campana}",
            help="Nombre base de los archivos"
        )
    
    st.markdown("---")
    
    if st.button("🚀 GENERAR ARCHIVOS SMS PARA CAMPAÑA", type="primary", use_container_width=True):
        
        st.markdown("### 📥 ARCHIVOS GENERADOS:")
        
        df_csv = df_campana[["NUMERO", "NOMBRE", "FECHA", "CODIGO", "SALDO_PENDIENTE"]].copy()
        df_csv = df_csv.rename(columns={"SALDO_PENDIENTE": "MONTO"})
        
        st.markdown('<div class="tipo-box">', unsafe_allow_html=True)
        st.markdown(f"""
        **📊 RESUMEN DE CAMPAÑA VIVA:**
        
        - **Tipos incluidos:** {', '.join(tipos_seleccionados)}
        - **Total registros:** {len(df_csv):,}
        - **Tipo de campaña:** {tipo_campana}
        - **Archivos a generar:** {num_archivos}
        - **Saldo total:** Bs. {df_csv['MONTO'].sum():,.2f}
        """)
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown("---")
        
        tamaño = len(df_csv) // num_archivos + 1
        
        for i in range(num_archivos):
            inicio = i * tamaño
            fin = inicio + tamaño
            df_parte = df_csv.iloc[inicio:fin]
            
            if df_parte.empty:
                continue
            
            csv = df_parte.to_csv(index=False, sep=";", encoding="utf-8-sig")
            nombre_archivo = f"{prefijo}_{i+1}.csv" if num_archivos > 1 else f"{prefijo}.csv"
            
            st.download_button(
                label=f"⬇️ {nombre_archivo} ({len(df_parte):,} registros | Bs. {df_parte['MONTO'].sum():,.2f})",
                data=csv,
                file_name=nombre_archivo,
                mime="text/csv",
                key=f"download_{i}",
                use_container_width=True
            )
        
        st.success(f"✅ {num_archivos} archivo(s) generado(s) exitosamente para campaña VIVA")
        st.balloons()


def modulo_control_diario():
    st.markdown('<div class="main-header">📊 CONTROL DIARIO DE RECUPERACIÓN Y OBJETIVOS</div>', unsafe_allow_html=True)
    
    if "objetivos" not in st.session_state:
        st.session_state.objetivos = {"diario": 10000, "semanal": 50000, "mensual": 200000}
    
    if "pagos_acumulados" not in st.session_state:
        st.session_state.pagos_acumulados = pd.DataFrame(columns=["CODIGO", "IMPORTE", "FECHA"])
    
    st.markdown("## 🎯 Configuración de Objetivos")
    
    with st.expander("⚙️ Editar Objetivos", expanded=False):
        st.info("💡 Define tus metas de recuperación. Puedes editarlas en cualquier momento.")
        col1, col2, col3 = st.columns(3)
        with col1:
            objetivo_diario = st.number_input("Meta Diaria (Bs.)", min_value=0, value=st.session_state.objetivos["diario"], step=1000)
        with col2:
            objetivo_semanal = st.number_input("Meta Semanal (Bs.)", min_value=0, value=st.session_state.objetivos["semanal"], step=5000)
        with col3:
            objetivo_mensual = st.number_input("Meta Mensual (Bs.)", min_value=0, value=st.session_state.objetivos["mensual"], step=10000)
        
        if st.button("💾 Guardar Objetivos", use_container_width=True):
            st.session_state.objetivos = {"diario": objetivo_diario, "semanal": objetivo_semanal, "mensual": objetivo_mensual}
            st.success("✅ Objetivos actualizados correctamente")
            st.rerun()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("🎯 Objetivo Diario", f"Bs. {st.session_state.objetivos['diario']:,.2f}")
    with col2:
        st.metric("🎯 Objetivo Semanal", f"Bs. {st.session_state.objetivos['semanal']:,.2f}")
    with col3:
        st.metric("🎯 Objetivo Mensual", f"Bs. {st.session_state.objetivos['mensual']:,.2f}")
    
    st.markdown("---")
    
    st.markdown("## 📥 Registrar Pagos del Día")
    
    col1, col2 = st.columns([2, 1])
    with col1:
        st.info("📂 Sube tu archivo Excel con los pagos. Acepta columnas CODIGO o ID_COBRANZA, y FECHA o FECHA_PAGO. Puedes cargar varias veces al día y se acumularán automáticamente.")
    with col2:
        if st.button("🔄 Limpiar Todo", use_container_width=True):
            st.session_state.pagos_acumulados = pd.DataFrame(columns=["CODIGO", "IMPORTE", "FECHA"])
            st.success("✅ Pagos eliminados")
            st.rerun()
    
    archivo_pagos_diarios = st.file_uploader(
        "Subir archivo PAGOS DIARIOS (acepta CODIGO/ID_COBRANZA, IMPORTE, FECHA/FECHA_PAGO)",
        type=["xlsx"],
        key="pagos_diarios"
    )
    
    if archivo_pagos_diarios:
        try:
            df_nuevos = pd.read_excel(archivo_pagos_diarios)
            df_nuevos.columns = df_nuevos.columns.str.strip().str.upper().str.replace(" ", "_")
            
            # ============================================
            # MAPEO DE COLUMNAS ALTERNATIVAS
            # CODIGO = ID_COBRANZA | FECHA = FECHA_PAGO
            # ============================================
            if "CODIGO" not in df_nuevos.columns and "ID_COBRANZA" in df_nuevos.columns:
                df_nuevos = df_nuevos.rename(columns={"ID_COBRANZA": "CODIGO"})
            
            if "FECHA" not in df_nuevos.columns and "FECHA_PAGO" in df_nuevos.columns:
                df_nuevos = df_nuevos.rename(columns={"FECHA_PAGO": "FECHA"})
            
            if not {"CODIGO", "IMPORTE", "FECHA"}.issubset(df_nuevos.columns):
                st.error("❌ El archivo debe tener: CODIGO (o ID_COBRANZA), IMPORTE, FECHA (o FECHA_PAGO)")
                st.error(f"**Columnas encontradas:** {', '.join(df_nuevos.columns)}")
            else:
                df_nuevos["CODIGO"] = df_nuevos["CODIGO"].astype(str)
                df_nuevos["IMPORTE"] = pd.to_numeric(df_nuevos["IMPORTE"], errors="coerce").fillna(0)
                df_nuevos["FECHA"] = pd.to_datetime(df_nuevos["FECHA"], errors="coerce")
                df_nuevos = df_nuevos.dropna(subset=["FECHA"])
                if len(df_nuevos) == 0:
                    st.error("❌ No hay registros válidos (revisa el formato de fecha)")
                else:
                    st.session_state.pagos_acumulados = pd.concat([st.session_state.pagos_acumulados, df_nuevos[["CODIGO", "IMPORTE", "FECHA"]]], ignore_index=True)
                    st.session_state.pagos_acumulados = st.session_state.pagos_acumulados.drop_duplicates()
                    st.success(f"✅ {len(df_nuevos):,} pagos agregados. Total: {len(st.session_state.pagos_acumulados):,}")
                    st.balloons()
                    st.rerun()
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
    
    if len(st.session_state.pagos_acumulados) == 0:
        st.warning("⚠️ No hay pagos registrados. Sube un archivo para comenzar.")
        return
    
    st.markdown("---")
    
    df_pagos = st.session_state.pagos_acumulados.copy()
    df_pagos["FECHA"] = pd.to_datetime(df_pagos["FECHA"])
    
    fecha_hoy = df_pagos["FECHA"].max().date()
    fecha_inicio_semana = fecha_hoy - timedelta(days=fecha_hoy.weekday())
    fecha_inicio_mes = fecha_hoy.replace(day=1)
    
    pagos_hoy = df_pagos[df_pagos["FECHA"].dt.date == fecha_hoy]
    pagos_semana = df_pagos[df_pagos["FECHA"].dt.date >= fecha_inicio_semana]
    pagos_mes = df_pagos[df_pagos["FECHA"].dt.date >= fecha_inicio_mes]
    
    recuperado_hoy = pagos_hoy["IMPORTE"].sum()
    recuperado_semana = pagos_semana["IMPORTE"].sum()
    recuperado_mes = pagos_mes["IMPORTE"].sum()
    
    porcentaje_dia = (recuperado_hoy / st.session_state.objetivos["diario"] * 100) if st.session_state.objetivos["diario"] > 0 else 0
    porcentaje_semana = (recuperado_semana / st.session_state.objetivos["semanal"] * 100) if st.session_state.objetivos["semanal"] > 0 else 0
    porcentaje_mes = (recuperado_mes / st.session_state.objetivos["mensual"] * 100) if st.session_state.objetivos["mensual"] > 0 else 0
    
    st.markdown(f"## 📅 HOY - {fecha_hoy.strftime('%A, %d de %B de %Y')}")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("🎯 Objetivo", f"Bs. {st.session_state.objetivos['diario']:,.2f}")
    with col2:
        delta = recuperado_hoy - st.session_state.objetivos['diario']
        st.metric("✅ Recuperado", f"Bs. {recuperado_hoy:,.2f}", f"{delta:+,.2f}")
    with col3:
        emoji = "🎉" if porcentaje_dia >= 100 else "⚠️" if porcentaje_dia >= 80 else "❌"
        st.metric("📊 Cumplimiento", f"{porcentaje_dia:.1f}% {emoji}", f"{len(pagos_hoy)} pagos")
    
    st.progress(min(porcentaje_dia / 100, 1.0))
    
    if porcentaje_dia >= 100:
        st.success(f"🎉 ¡Excelente! Superaste el objetivo en {porcentaje_dia - 100:.1f}%")
    elif porcentaje_dia >= 80:
        st.warning(f"⚠️ Vas bien, faltan Bs. {st.session_state.objetivos['diario'] - recuperado_hoy:,.2f}")
    else:
        st.error(f"❌ Acelera: faltan Bs. {st.session_state.objetivos['diario'] - recuperado_hoy:,.2f}")
    
    st.markdown("---")
    st.markdown("## 📅 ESTA SEMANA")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("🎯 Objetivo", f"Bs. {st.session_state.objetivos['semanal']:,.2f}")
    with col2:
        st.metric("✅ Recuperado", f"Bs. {recuperado_semana:,.2f}")
    with col3:
        emoji = "🎉" if porcentaje_semana >= 100 else "⚠️" if porcentaje_semana >= 80 else "❌"
        st.metric("📊 Cumplimiento", f"{porcentaje_semana:.1f}% {emoji}")
    
    st.progress(min(porcentaje_semana / 100, 1.0))
    
    st.markdown("---")
    st.markdown("## 📅 ESTE MES")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("🎯 Objetivo", f"Bs. {st.session_state.objetivos['mensual']:,.2f}")
    with col2:
        st.metric("✅ Recuperado", f"Bs. {recuperado_mes:,.2f}")
    with col3:
        st.metric("⏰ Falta", f"Bs. {st.session_state.objetivos['mensual'] - recuperado_mes:,.2f}")
    with col4:
        emoji = "🎉" if porcentaje_mes >= 100 else "⚠️" if porcentaje_mes >= 80 else "❌"
        st.metric("📊 Cumplimiento", f"{porcentaje_mes:.1f}% {emoji}")
    
    st.progress(min(porcentaje_mes / 100, 1.0))
    
    dias_transcurridos = (fecha_hoy - fecha_inicio_mes).days + 1
    promedio_diario = recuperado_mes / dias_transcurridos
    proyeccion = promedio_diario * 30
    
    st.markdown("### 🔮 Proyección Fin de Mes")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Proyección", f"Bs. {proyeccion:,.2f}", f"{(proyeccion / st.session_state.objetivos['mensual'] * 100):.1f}%")
    with col2:
        necesitas = (st.session_state.objetivos['mensual'] - recuperado_mes) / max(30 - dias_transcurridos, 1)
        st.metric("Necesitas/día", f"Bs. {necesitas:,.2f}")


if menu == "📊 Dashboard Cruce Deuda vs Pagos":
    modulo_cruce()
elif menu == "📈 Gráficos Interactivos":
    modulo_graficos()
elif menu == "📲 GENERADOR DE SMS":
    modulo_sms()
elif menu == "📈 Control Diario y Objetivos":
    modulo_control_diario()
