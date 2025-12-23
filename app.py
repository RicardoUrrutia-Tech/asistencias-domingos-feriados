import io
import re
from datetime import datetime, date
from typing import List, Set, Tuple

import pandas as pd
import streamlit as st

# ----------------------------
# Config
# ----------------------------
st.set_page_config(page_title="Reporte Domingos y Feriados", layout="wide")

INVALID_TURNOS = {"L", ""}  # "L" y vacío NO cuentan como trabajado


# ----------------------------
# Helpers
# ----------------------------
def _normalize_turno(x) -> str:
    """Normaliza el turno a string limpio. None/NaN -> ''."""
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    s = str(x).strip()
    # Si viene como "nan" texto
    if s.lower() == "nan" or s.lower() == "none":
        return ""
    return s


def is_turno_valido(turno: str) -> bool:
    """Valida turno: todo cuenta excepto 'L' y vacío."""
    t = _normalize_turno(turno)
    return t not in INVALID_TURNOS


def detect_date_columns(df: pd.DataFrame, meta_cols: List[str]) -> List:
    """
    Detecta columnas de fecha del reporte:
    - En tu archivo suelen venir como datetime en el header (Excel).
    """
    date_cols = []
    for c in df.columns:
        if c in meta_cols:
            continue
        # Caso típico: ya es Timestamp/datetime/date
        if isinstance(c, (pd.Timestamp, datetime, date)):
            date_cols.append(c)
            continue
        # Si viene como string "2025-12-01" o similar
        if isinstance(c, str):
            c2 = c.strip()
            try:
                _ = pd.to_datetime(c2, errors="raise")
                date_cols.append(c)
            except Exception:
                pass
    return date_cols


def parse_holidays(text: str) -> Set[pd.Timestamp]:
    """
    Acepta fechas separadas por coma o salto de línea.
    Formatos soportados típicos:
    - dd-mm-aaaa
    - dd/mm/aaaa
    - aaaa-mm-dd
    - aaaa/mm/dd
    - dd.mm.aaaa
    """
    if not text or not text.strip():
        return set()

    parts = re.split(r"[,\n;]+", text.strip())
    holidays = set()

    for p in parts:
        p = p.strip()
        if not p:
            continue

        # Normaliza separadores
        p_norm = p.replace(".", "-").replace("/", "-")

        dt_val = pd.to_datetime(p_norm, errors="coerce", dayfirst=True)
        if pd.isna(dt_val):
            # Intento alternativo (por si venía yyyy-mm-dd y dayfirst lo confunde)
            dt_val = pd.to_datetime(p_norm, errors="coerce", dayfirst=False)

        if pd.isna(dt_val):
            raise ValueError(f"No pude interpretar esta fecha de feriado: '{p}'")

        holidays.add(pd.Timestamp(dt_val.date()))

    return holidays


def build_summary(
    df: pd.DataFrame,
    meta_cols: List[str],
    date_cols: List,
    holidays: Set[pd.Timestamp],
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, str]:
    """
    Devuelve 3 dataframes:
      - domingos
      - festivos
      - total (domingos + festivos)
    """
    if not date_cols:
        raise ValueError("No se detectaron columnas de fecha en el archivo. Revisa el formato del reporte.")

    # Melt largo
    dfl = df.melt(id_vars=meta_cols, value_vars=date_cols, var_name="Fecha", value_name="Turno")
    dfl["Turno_norm"] = dfl["Turno"].apply(_normalize_turno)
    dfl["Trabajado"] = dfl["Turno_norm"].apply(is_turno_valido)

    # Parse Fecha a fecha (sin hora)
    dfl["Fecha_dt"] = pd.to_datetime(dfl["Fecha"], errors="coerce")
    if dfl["Fecha_dt"].isna().all():
        raise ValueError("No pude convertir las columnas de fecha a formato fecha. Revisa los encabezados.")

    dfl["Fecha_dt"] = dfl["Fecha_dt"].dt.date
    dfl["Fecha_ts"] = dfl["Fecha_dt"].apply(lambda x: pd.Timestamp(x))

    # Periodo detectado
    start = min(dfl["Fecha_dt"])
    end = max(dfl["Fecha_dt"])
    periodo_str = f"{start.strftime('%d-%m-%Y')} a {end.strftime('%d-%m-%Y')}"

    # Domingos
    dfl["Es_domingo"] = pd.to_datetime(dfl["Fecha_dt"]).dt.weekday == 6  # lunes=0 ... domingo=6
    dom = dfl[(dfl["Trabajado"]) & (dfl["Es_domingo"])].copy()

    # Festivos (manuales)
    if holidays:
        fest = dfl[(dfl["Trabajado"]) & (dfl["Fecha_ts"].isin(holidays))].copy()
    else:
        fest = dfl.iloc[0:0].copy()  # vacío con mismas cols

    # Aggregations
    def agg_table(sub: pd.DataFrame, label_count: str, label_dates: str) -> pd.DataFrame:
        if sub.empty:
            # Igual devolvemos colaboradores con 0 si queremos? Preferí devolver solo los que tienen conteo>0,
            # pero aquí lo haremos "completo" con 0 para que sea consistente.
            base = df[meta_cols].copy()
            base[label_count] = 0
            base[label_dates] = ""
            return base

        grp = (
            sub.groupby(meta_cols, dropna=False)["Fecha_dt"]
            .apply(lambda s: sorted(set(s)))
            .reset_index()
        )
        grp[label_count] = grp["Fecha_dt"].apply(len)
        grp[label_dates] = grp["Fecha_dt"].apply(lambda lst: ", ".join([pd.Timestamp(x).strftime("%d-%m-%Y") for x in lst]))
        grp = grp.drop(columns=["Fecha_dt"])
        return grp

    dom_tbl = agg_table(dom, "Domingos trabajados", "Fechas (domingos)")
    fest_tbl = agg_table(fest, "Festivos trabajados", "Fechas (festivos)")

    # Total = unión
    # Unimos fechas por colaborador
    dom_map = dom.groupby(meta_cols, dropna=False)["Fecha_dt"].apply(lambda s: set(s)).to_dict() if not dom.empty else {}
    fest_map = fest.groupby(meta_cols, dropna=False)["Fecha_dt"].apply(lambda s: set(s)).to_dict() if not fest.empty else {}

    total_rows = []
    for _, row in df[meta_cols].drop_duplicates().iterrows():
        key = tuple(row[c] for c in meta_cols)
        s_dom = dom_map.get(key, set())
        s_fest = fest_map.get(key, set())
        s_all = set(s_dom) | set(s_fest)

        total_rows.append(
            {
                **{c: row[c] for c in meta_cols},
                "Domingos trabajados": len(s_dom),
                "Festivos trabajados": len(s_fest),
                "Domingos + Festivos": len(s_all),
                "Fechas (domingos)": ", ".join([pd.Timestamp(x).strftime("%d-%m-%Y") for x in sorted(s_dom)]),
                "Fechas (festivos)": ", ".join([pd.Timestamp(x).strftime("%d-%m-%Y") for x in sorted(s_fest)]),
                "Fechas (total)": ", ".join([pd.Timestamp(x).strftime("%d-%m-%Y") for x in sorted(s_all)]),
            }
        )

    total_tbl = pd.DataFrame(total_rows)

    return dom_tbl, fest_tbl, total_tbl, periodo_str


def export_excel(dom_tbl: pd.DataFrame, fest_tbl: pd.DataFrame, total_tbl: pd.DataFrame, periodo: str, holidays: Set[pd.Timestamp]) -> bytes:
    output = io.BytesIO()

    # Info para encabezado
    feriados_str = ", ".join(sorted([h.strftime("%d-%m-%Y") for h in holidays])) if holidays else "(sin feriados ingresados)"

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Escribimos dejando espacio arriba para notas
        startrow = 4

        dom_sheet = "Domingos trabajados"
        fest_sheet = "Festivos trabajados"
        tot_sheet = "Domingos + Festivos"

        dom_tbl.to_excel(writer, index=False, sheet_name=dom_sheet, startrow=startrow)
        fest_tbl.to_excel(writer, index=False, sheet_name=fest_sheet, startrow=startrow)
        total_tbl.to_excel(writer, index=False, sheet_name=tot_sheet, startrow=startrow)

        wb = writer.book

        def decorate(ws_name: str, title: str):
            ws = wb[ws_name]
            ws["A1"] = title
            ws["A2"] = f"Periodo detectado: {periodo}"
            ws["A3"] = "Regla de conteo: se considera 'trabajado' cuando el turno es distinto de 'L' y distinto de vacío. (COON1 cuenta como válido)."
            if ws_name != dom_sheet:
                ws["A4"] = f"Feriados ingresados manualmente: {feriados_str}"

            # Freeze header
            ws.freeze_panes = ws["A6"]  # 1 título + 4 filas info + 1 header (startrow=4 => header en fila 5)
            # Auto width simple
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col[:1500]:  # límite razonable
                    val = "" if cell.value is None else str(cell.value)
                    if len(val) > max_len:
                        max_len = len(val)
                ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 60)

        decorate(dom_sheet, "Reporte de asistencias: Domingos trabajados")
        decorate(fest_sheet, "Reporte de asistencias: Festivos trabajados")
        decorate(tot_sheet, "Reporte de asistencias: Domingos + Festivos trabajados")

    return output.getvalue()


# ----------------------------
# UI
# ----------------------------
st.title("Reporte de asistencias: Domingos y Feriados")

st.markdown(
    """
Sube tu **reporte de turnos** (Excel).  
Luego ingresa manualmente los **feriados** (uno por línea o separados por coma) y descarga el Excel con 3 hojas:
1) Domingos trabajados, 2) Festivos trabajados, 3) Domingos + Festivos.
"""
)

uploaded = st.file_uploader("📤 Cargar reporte de turnos (.xlsx)", type=["xlsx"])

col1, col2 = st.columns([1, 1])
with col1:
    feriados_text = st.text_area(
        "🗓️ Ingresar feriados (manual)",
        placeholder="Ej:\n01-01-2026\n18-09-2026\n25-12-2026\n\n(o separados por coma)",
        height=140,
    )
with col2:
    st.info("Regla: solo se excluye **L** y **vacío**. Todo lo demás cuenta como trabajado (incluye **COON1**).")

if uploaded:
    try:
        xls = pd.ExcelFile(uploaded)
        sheet = st.selectbox("Hoja del Excel a procesar", xls.sheet_names, index=0)

        df = pd.read_excel(uploaded, sheet_name=sheet)

        meta_cols = ["Nombre del Colaborador", "RUT", "Área", "Supervisor"]
        missing = [c for c in meta_cols if c not in df.columns]
        if missing:
            st.error(f"Faltan columnas esperadas: {missing}. Revisa que el reporte tenga esas columnas.")
            st.stop()

        date_cols = detect_date_columns(df, meta_cols)

        holidays = parse_holidays(feriados_text)

        dom_tbl, fest_tbl, total_tbl, periodo_str = build_summary(df, meta_cols, date_cols, holidays)

        st.success(f"✅ Periodo detectado: {periodo_str}")

        tab1, tab2, tab3 = st.tabs(["Domingos", "Festivos", "Domingos + Festivos"])
        with tab1:
            st.dataframe(dom_tbl, use_container_width=True)
        with tab2:
            st.dataframe(fest_tbl, use_container_width=True)
        with tab3:
            st.dataframe(total_tbl, use_container_width=True)

        file_bytes = export_excel(dom_tbl, fest_tbl, total_tbl, periodo_str, holidays)

        st.download_button(
            "⬇️ Descargar Excel (3 hojas)",
            data=file_bytes,
            file_name="Reporte_Asistencias_Domingos_y_Festivos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except ValueError as ve:
        st.error(str(ve))
    except Exception as e:
        st.exception(e)
else:
    st.warning("Carga un archivo .xlsx para comenzar.")
