import streamlit as st
import pandas as pd
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

def modulo_sms():
    st.title("📲 GENERADOR DE SMS")

    def limpiar_columnas(df):
        df.columns = df.columns.str.strip().str.upper().str.replace(" ", "_")
        return df

    archivo_suscriptor = st.file_uploader("📂 BASE SUSCRIPTOR", type=["xlsx"])
    archivo_pagos = st.file_uploader("💵 BASE PAGOS", type=["xlsx"])

    if not archivo_suscriptor or not archivo_pagos:
        return

    df_suscriptor = limpiar_columnas(pd.read_excel(archivo_suscriptor))
    df_pagos = limpiar_columnas(pd.read_excel(archivo_pagos))

    df_suscriptor["CODIGO"] = df_suscriptor["CODIGO"].astype(str)
    df_suscriptor["MONTO"] = pd.to_numeric(df_suscriptor["MONTO"], errors="coerce").fillna(0)
    df_pagos["ID_COBRANZA"] = df_pagos["ID_COBRANZA"].astype(str)
    df_pagos["IMPORTE"] = pd.to_numeric(df_pagos["IMPORTE"], errors="coerce").fillna(0)

    pagos_totales = df_pagos.groupby("ID_COBRANZA")["IMPORTE"].sum().reset_index()
    pagos_totales.rename(columns={"IMPORTE": "TOTAL_PAGADO"}, inplace=True)

    df_final = df_suscriptor.merge(pagos_totales, left_on="CODIGO", right_on="ID_COBRANZA", how="left")
    df_final["TOTAL_PAGADO"] = df_final["TOTAL_PAGADO"].fillna(0)
    df_final = df_final[df_final["TOTAL_PAGADO"] < df_final["MONTO"]]

    columnas_exportar = ["NUMERO", "NOMBRE", "FECHA", "CODIGO", "MONTO"]
    df_export = df_final[columnas_exportar].copy()

    st.dataframe(df_export)

    partes = st.number_input("Archivos CSV", min_value=1, value=1)
    prefijo = st.text_input("Prefijo", value="SMS")

    if st.button("Generar CSV"):
        if df_export.empty:
            st.warning("Sin registros")
            return

        tamaño = len(df_export) // partes + 1
        for i in range(partes):
            inicio = i * tamaño
            fin = inicio + tamaño
            df_parte = df_export.iloc[inicio:fin]
            if df_parte.empty:
                continue

            csv = df_parte.to_csv(index=False, sep=";", encoding="utf-8-sig")
            st.download_button(
                label=f"Descargar {prefijo}_{i+1}.csv",
                data=csv,
                file_name=f"{prefijo}_{i+1}.csv",
                mime="text/csv"
            )

if menu == "📊 Dashboard Cruce Deuda vs Pagos":
    modulo_cruce()
elif menu == "📲 GENERADOR DE SMS":
    modulo_sms()
elif menu == "🚧 Módulo Histórico (En Desarrollo)":
    st.title("📈 Histórico")
    st.info("Módulo en desarrollo")
