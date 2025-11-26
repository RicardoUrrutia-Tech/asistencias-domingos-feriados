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

# Renombrar columnas
df.columns = [
    "Codigo", "RUT", "Nombre", "PrimerApellido", "SegundoApellido",
    "Especialidad", "Area", "Contrato", "Supervisor", "Turno",
    "EntradaFecha", "EntradaHora", "SalidaFecha", "SalidaHora",
    "RUT_Empleador", "DentroRecintoEntrada", "DentroRecintoSalida"
]

st.success("Archivo cargado correctamente.")
st.dataframe(df.head())

# -------------------------------------------
# CONVERTIR FECHAS
# -------------------------------------------
df["EntradaFecha"] = pd.to_datetime(df["EntradaFecha"], errors="coerce")
df["SalidaFecha"] = pd.to_datetime(df["SalidaFecha"], errors="coerce")

# Dia de la semana en español
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

# Convertir feriados a datetime64 para poder comparar
feriados_datetime = pd.to_datetime(feriados)

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
    if len(feriados_datetime) > 0:
        df["EsFeriado"] = df["EntradaFecha"].dt.normalize().isin(feriados_datetime)
        feriados_df = df[df["EsFeriado"] == True]

        st.subheader("🎉 Registros trabajados en Feriados")
        st.dataframe(feriados_df)
    else:
        feriados_df = pd.DataFrame()
        st.info("No ingresaste feriados.")

    # -------------------------
    # DESCARGA
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

