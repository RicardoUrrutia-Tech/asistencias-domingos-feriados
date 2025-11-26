import streamlit as st
import pandas as pd
from io import BytesIO

# -------------------------------------------
# CONFIGURACIÓN DE LA APP
# -------------------------------------------
st.set_page_config(page_title="Reporte Domingos y Feriados", layout="wide")
st.title("🟦 Analizador de Asistencias BUK – Domingos y Feriados")

# -------------------------------------------
# SUBIR ARCHIVO
# -------------------------------------------
uploaded = st.file_uploader("Sube el archivo de asistencia (.xlsx)", type=["xlsx"])

if uploaded:
    # Leer el archivo (asumiendo encabezado plano como tu XLSX convertido)
    df = pd.read_excel(uploaded, header=0)

    # Renombrar columnas para evitar errores
    df.columns = [
        "Codigo", "RUT", "Nombre", "PrimerApellido", "SegundoApellido",
        "Especialidad", "Area", "Contrato", "Supervisor", "Turno",
        "EntradaFecha", "EntradaHora", "SalidaFecha", "SalidaHora",
        "RUT_Empleador", "DentroRecintoEntrada", "DentroRecintoSalida"
    ]

    st.subheader("📄 Vista previa del archivo cargado")
    st.dataframe(df.head())

    # -------------------------------------------
    # LIMPIEZA DE FECHAS
    # -------------------------------------------
    df["EntradaFecha"] = pd.to_datetime(df["EntradaFecha"], errors="coerce")
    df["SalidaFecha"] = pd.to_datetime(df["SalidaFecha"], errors="coerce")

    # Día de semana (en español)
    df["DiaSemana"] = df["EntradaFecha"].dt.day_name(locale="es_ES")

    # -------------------------------------------
    # DOMINGOS
    # -------------------------------------------
    domingos = df[df["DiaSemana"] == "domingo"]

    st.subheader("👷‍♂️ Registros trabajados en Domingo")
    st.dataframe(domingos)

    # -------------------------------------------
    # FERIADOS
    # -------------------------------------------
    st.subheader("📅 Seleccione feriados (puede ingresar varios)")

    feriados = st.date_input("Ingrese fechas de feriados", [])

    if len(feriados) > 0:
        feriados = pd.to_datetime(feriados)
        df["EsFeriado"] = df["EntradaFecha"].isin(feriados)
        feriados_df = df[df["EsFeriado"] == True]

        st.subheader("🎉 Registros trabajados en Feriados")
        st.dataframe(feriados_df)
    else:
        feriados_df = pd.DataFrame()

    # -------------------------------------------
    # DESCARGA DE ARCHIVO FINAL
    # -------------------------------------------
    st.subheader("📥 Descargar reporte consolidado")

    def to_excel(df_domingos, df_feriados):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine="xlsxwriter")

        df_domingos.to_excel(writer, index=False, sheet_name="Domingos")
        df_feriados.to_excel(writer, index=False, sheet_name="Feriados")

        writer.close()
        return output.getvalue()

    excel_bytes = to_excel(domingos, feriados_df)

    st.download_button(
        label="Descargar Excel",
        data=excel_bytes,
        file_name="Reporte_Domingos_Feriados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
