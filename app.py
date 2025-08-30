# main_dupont_streamlit.py
import streamlit as st
import pandas as pd
from io import BytesIO

# Cálculo DuPont triple
def duPont_triple(util_neta, ventas, activos_totales, capital_contable):
    # Evitar división por cero
    margen_utilidad = (util_neta / ventas) if ventas != 0 else None
    rotacion_activos = (ventas / activos_totales) if activos_totales != 0 else None
    apalancamiento = (activos_totales / capital_contable) if capital_contable != 0 else None
    roe = (
        margen_utilidad * rotacion_activos * apalancamiento
        if None not in (margen_utilidad, rotacion_activos, apalancamiento)
        else None
    )
    return margen_utilidad, rotacion_activos, apalancamiento, roe

# Cargar Excel subido
def cargar_datos_desde_excel(uploaded_file):
    # Leer con pandas; exige que el archivo sea .xlsx/.xlsm y que openpyxl esté disponible
    try:
        df = pd.read_excel(uploaded_file, header=0, engine="openpyxl")
        return df
    except Exception as e:
        raise e

# Detectar periodos a partir de columnas (asumiendo que hay columna "Concepto" o filas por concepto)
def detectar_periodos(df):
    # Si hay columna llamada "Concepto", asumimos formato ancho donde cada columna distinta es un periodo
    if "Concepto" in df.columns:
        periodos = [col for col in df.columns if col != "Concepto"]
        return periodos
    # Si no hay columna Concepto, asumimos que cada fila es un periodo (menos cabeceras)
    # En este caso, devolveremos las columnas menos las conocidas (por ejemplo "Empresa" si existe)
    periodos = [col for col in df.columns if col not in ["Empresa"]]
    return periodos

# Conceptos esperados (filas)
def conceptos_base():
    return ["Ventas Netas", "Utilidad Neta", "Activo Total", "Capital Contable"]

# Construcción del reporte final
def generar_reporte(df, periodos, empresa_nombre="Empresa_1"):
    resultados = []
    # Si el archivo está en formato ancho con una fila por concepto
    # intentamos extraer por periodo buscando cada concepto en la fila correspondiente.
    # Si hay columna "Concepto", usamos esa columna para filtrar.
    if "Concepto" in df.columns:
        for periodo in periodos:
            try:
                v = df.loc[df["Concepto"] == "Ventas Netas", periodo].values
                u = df.loc[df["Concepto"] == "Utilidad Neta", periodo].values
                a = df.loc[df["Concepto"] == "Activo Total", periodo].values
                c = df.loc[df["Concepto"] == "Capital Contable", periodo].values
                # Tomamos el primer valor si existe
                v = float(v[0]) if len(v) > 0 and pd.notna(v[0]) else None
                u = float(u[0]) if len(u) > 0 and pd.notna(u[0]) else None
                a = float(a[0]) if len(a) > 0 and pd.notna(a[0]) else None
                c = float(c[0]) if len(c) > 0 and pd.notna(c[0]) else None
            except Exception:
                v = u = a = c = None

            if None not in (v, u, a, c) and v != 0 and a != 0 and c != 0:
                margen, rot, apal, roe = duPont_triple(u, v, a, c)
                if None not in (margen, rot, apal, roe):
                    resultados.append({
                        "Empresa": empresa_nombre,
                        "Periodo": periodo,
                        "Ventas Netas": v,
                        "Utilidad Neta": u,
                        "Activo Total": a,
                        "Capital Contable": c,
                        "Margen_utilidad_neta": margen,
                        "Rotacion_activos": rot,
                        "Apalancamiento": apal,
                        "ROE": roe
                    })
        df_result = pd.DataFrame(resultados)
        return df_result
    else:
        # Caso alternativo: filas por concepto y columnas por periodo
        # extraemos por periodo buscando filas y columnas correspondientes
        periodos = [p for p in periodos]
        for periodo in periodos:
            try:
                v = df.loc[df["Concepto"] == "Ventas Netas", periodo].values[0]
                u = df.loc[df["Concepto"] == "Utilidad Neta", periodo].values[0]
                a = df.loc[df["Concepto"] == "Activo Total", periodo].values[0]
                c = df.loc[df["Concepto"] == "Capital Contable", periodo].values[0]
                if pd.notna(v) and pd.notna(u) and pd.notna(a) and pd.notna(c) and v != 0 and a != 0 and c != 0:
                    margen, rot, apal, roe = duPont_triple(float(u), float(v), float(a), float(c))
                    resultados.append({
                        "Empresa": empresa_nombre,
                        "Periodo": periodo,
                        "Ventas Netas": v,
                        "Utilidad Neta": u,
                        "Activo Total": a,
                        "Capital Contable": c,
                        "Margen_utilidad_neta": margen,
                        "Rotacion_activos": rot,
                        "Apalancamiento": apal,
                        "ROE": roe
                    })
            except Exception:
                continue
        df_result = pd.DataFrame(resultados)
        return df_result

# Interfaz principal
def main():
    st.title("DuPont - Cálculo por periodo (renglones = conceptos, columnas = periodos)")
    st.write("Lectura de Excel: filas = conceptos (Ventas Netas, Utilidad Neta, Activo Total, Capital Contable); columnas = periodos.")

    uploaded = st.file_uploader("Carga tu Excel (.xlsx/.xlsm)", type=["xlsx", "xlsm"])
    if uploaded is None:
        st.info("Por favor, carga un archivo de Excel para continuar.")
        return

    try:
        df = cargar_datos_desde_excel(uploaded)
    except Exception as e:
        st.error(f"Error leyendo el Excel: {e}")
        return

    # Detección de periodos
    periodos = detectar_periodos(df)
    # Si se detecta una columna de empresa, se puede adaptar; por ahora usaremos un nombre por defecto
    empresa = "Empresa_1"

    # Generar reporte
    with st.spinner("Calculando DuPont por periodo..."):
        try:
            df_result = generar_reporte(df, periodos, empresa_nombre=empresa)
        except Exception as e:
            st.error(f"Error al generar reporte: {e}")
            return

    if df_result is None or df_result.empty:
        st.warning("No se pudo generar un reporte. Verifique el formato del archivo.")
        return

    st.subheader("Reporte DuPont por periodo")
    st.dataframe(df_result)

    # Descargas
    csv = df_result.to_csv(index=False).encode("utf-8")
    st.download_button(label="Descargar CSV", data=csv, file_name="dupont_reporte.csv", mime="text/csv")

    # Excel en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_result.to_excel(writer, index=False, sheet_name="DuPont")
    excel_data = output.getvalue()
    st.download_button(label="Descargar Excel", data=excel_data,
                       file_name="dupont_reporte.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
