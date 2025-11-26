import streamlit as st
import pandas as pd
from io import BytesIO

# -------------------------------------------
# CONFIGURACIÓN DE LA APP
# -------------------------------------------
st.set_page_config(page_title="Reporte Domingos y Feriados", layout="wide")
st.title("🟦 Analizador de Asistencias BUK – Domingos y Feriados")

# -------------------------------------------
# SUBIR ARCHIVO (PASO 1)
# -------------------------------------------
uploaded = st.file_uploader("📤 Sube el archivo de asistencia (.xlsx)", type=["xlsx"])

# Si no hay archivo, no hacemos nada más
if not uploaded:
    st.info("Sube un archivo para continuar.")
    st.stop()

# -------------------------------------------
# LEER ARCHIVO
# -------------------------------------------
df = pd.read_excel(uploaded, header=0)

# -------------------------------------------
# RENOMBRAR COLUMNAS SEGÚN ARCHIVO REAL
# -------------------------------------------
df = df.rename(columns={
    "Entrada": "EntradaFecha",
    "Unnamed: 11": "EntradaHora",
    "Salida": "SalidaFecha",
    "Unnamed: 13": "SalidaHora"
})

# Mostrar preview
st.success("Archivo cargado correctamente.")
st.dataframe(df.head())

# -------------------------------------------
# CONVERTIR FECHAS
# -------------------------------------------
df["EntradaFecha"] = pd.to_datetime(df["EntradaFecha"], format="%d/%m/%Y", errors="coerce")
df["SalidaFecha"] = pd.to_datetime(df["SalidaFecha"], format="%d/%m/%Y", errors="coerce")

# -------------------------------------------
# DÍA DE SEMANA EN ESPAÑOL (SIN locale)
# -------------------------------------------
dias_es = {
    "Monday": "lunes",
    "Tuesday": "martes",
    "Wednesday": "miércoles",
    "Thursday": "jueves",
    "Friday": "viernes",
    "Saturday": "sábado",
    "Sunday": "domingo"
}

df["DiaSemana"] = df["EntradaFecha"].dt.day_name().map(dias_es)

# -------------------------------------------
# INGRESAR FERIADOS (PASO 2)
# -------------------------------------------
st.subheader("📅 Ingrese fechas de feriado (opcional)")
feriados = st.date_input("Seleccione uno o más feriados", [])

# Convertir feriados a date
feriados_date = [f for f in feriados]

# Convertir EntradaFecha a date para comparación exacta
df["EntradaFechaDate"] = df["EntradaFecha"].dt.date

# -------------------------------------------
# BOTÓN DE PROCESAR (PASO 3)
# -------------------------------------------
if st.button("Procesar Reporte"):

    # -------------------------
    # DOMINGOS
    # -------------------------
    domingos = df[df["DiaSemana"] == "domingo"]

    st.subheader("👷‍♂️ Registros trabajados en Domingo")
    st.dataframe(domingos)

    # -------------------------
    # FERIADOS
    # -------------------------
    if len(feriados_date) > 0:
        df["EsFeriado"] = df["EntradaFechaDate"].isin(feriados_date)
        feriados_df = df[df["EsFeriado"] == True]

        st.subheader("🎉 Registros trabajados en Feriados")
        st.dataframe(feriados_df)
    else:
        feriados_df = pd.DataFrame()
        st.info("No ingresaste feriados.")

    # -------------------------
    # DESCARGA DEL REPORTE
    # -------------------------
    st.subheader("📥 Descargar reporte")

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
