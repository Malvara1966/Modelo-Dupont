# main_dupont_streamlit.py
import streamlit as st
import pandas as pd

def duPont_triple(util_neta, ventas, activos_totales, capital_contable):
    margen_utilidad = util_neta / ventas
    rotacion_activos = ventas / activos_totales
    apalancamiento = activos_totales / capital_contable
    roe = margen_utilidad * rotacion_activos * apalancamiento
    return margen_utilidad, rotacion_activos, apalancamiento, roe

def cargar_datos_desde_excel(uploaded_file):
    df = pd.read_excel(uploaded_file, header=0)
    return df

def extraer_periodos(df):
    # Periodos son las columnas distintas (ej.: 2023, 2024, etc.)
    periodos = sorted([col for col in df.columns if col not in ["Empresa", "Empresa"].__class__.__name__])
    # El filtrado anterior se ajusta si tienes una columna de etiqueta como "Empresa"
    # En este caso asumimos que las columnas que no son renglones/conceptos pueden ser periodos
    # Mejor: tomamos todas las columnas excepto las filas que describen los conceptos por fila
    return periodos

def base_conceptos():
    # Nombres de las filas (renglones) que contienen los conceptos
    return ["Ventas Netas", "Utilidad Neta", "Activo Total", "Capital Contable"]

def formato_resultados(df, periodos, conceptos):
    resultados = []
    # Recorremos cada periodo y cada empresa (fila) para obtener columnas por periodo
    for idx, row in df.iterrows():
        empresa = row.get("Empresa", f"Empresa_{idx+1}")
        for periodo in periodos:
            # Las celdas por periodo para cada concepto están organizadas como filas
            # Buscamos el valor en la fila de concepto para ese periodo
            try:
                v = df.loc[df["Concepto"] == "Ventas Netas", periodo].values[0]
                u = df.loc[df["Concepto"] == "Utilidad Neta", periodo].values[0]
                a = df.loc[df["Concepto"] == "Activo Total", periodo].values[0]
                c = df.loc[df["Concepto"] == "Capital Contable", periodo].values[0]
            except Exception:
                # Si la estructura difiere, saltar ese periodo
                continue

            if pd.notna(v) and pd.notna(u) and pd.notna(a) and pd.notna(c) and v != 0 and a != 0 and c != 0:
                margen, rot, apal, roe = duPont_triple(u, v, a, c)
                resultados.append({
                    "Empresa": empresa,
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
    return pd.DataFrame(resultados)

def main():
    st.title("Análisis DuPont - Cálculo pormenorizado (renglones = conceptos, columnas = periodos)")
    st.write("Lectura de Excel: filas = conceptos (Ventas Netas, Utilidad Neta, Activo Total, Capital Contable); columnas = periodos (años/meses).")

    uploaded = st.file_uploader("Carga tu Excel (.xlsx)", type=["xlsx", "xlsm"])
    if uploaded is None:
        st.info("Por favor, carga un archivo de Excel para continuar.")
        return

    df = cargar_datos_desde_excel(uploaded)

    # Esperamos una columna 'Concepto' para distinguir renglones y una columna 'Empresa' opcional
    # Si tu archivo ya está en formato limpio, adapta las columnas accordingly.
    if "Concepto" in df.columns:
        # ya está en formato largo: fila por concepto, columna por periodo
        # Transforma a formato ancho para calcular por periodo
        periodos = [c for c in df.columns if c not in ["Concepto"]]
    else:
        # Si no hay columna Concepto, asumimos que las primeras filas definen conceptos
        # Este caso requiere estructura específica; para robustez, solicita formato estandarizado.
        st.error("Formato de archivo no reconocido: se espera columna 'Concepto'.")
        return

    conceptos = base_conceptos()
    df_formato = df.copy()
    # Genera el reporte suponiendo que cada fila es un concepto y las columnas son periodos
    # Necesitamos convertir a formato donde cada periodo por fila se calcule; aquí asumimos ya en ancho.
    # Reorganizamos para obtener por periodo usando la fila de cada concepto.
    # Crear DF de resultados
    resultados = []
    periodos = [col for col in df.columns if col != "Concepto"]

    for periodo in periodos:
        try:
            v = df.loc[df["Concepto"] == "Ventas Netas", periodo].iloc[0]
            u = df.loc[df["Concepto"] == "Utilidad Neta", periodo].iloc[0]
            a = df.loc[df["Concepto"] == "Activo Total", periodo].iloc[0]
            c = df.loc[df["Concepto"] == "Capital Contable", periodo].iloc[0]
            if all(pd.notna(x) for x in [v, u, a, c]) and v != 0 and a != 0 and c != 0:
                margen, rot, apal, roe = duPont_triple(u, v, a, c)
                resultados.append({
                    "Empresa": "Empresa_1",  # si tienes nombre por fila, ajusta aquí
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

    st.subheader("Reporte DuPont por periodo")
    st.dataframe(df_result)

    csv = df_result.to_csv(index=False).encode('utf-8')
    st.download_button(label="Descargar CSV", data=csv, file_name="dupont_reporte.csv", mime="text/csv")

    # Nota: para Excel, puede generarse en memoria usando io.BytesIO
    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_result.to_excel(writer, index=False, sheet_name="DuPont")
    excel_data = output.getvalue()
    st.download_button(label="Descargar Excel", data=excel_data,
                       file_name="dupont_reporte.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
