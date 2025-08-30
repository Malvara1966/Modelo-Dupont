DuPont - Cálculo pormenorizado (renglones = conceptos, columnas = periodos)
Aplicación en Python con Streamlit que lee un libro de Excel donde:

Filas: conceptos contables de estados financieros (Ventas Netas, Utilidad Neta, Activo Total, Capital Contable).
Columnas: periodos de tiempo (años o meses).
La app calcula el modelo DuPont por periodo y genera un reporte interactivo, con opciones para descargar los resultados en CSV y Excel. Está diseñada para ser desplegada desde Github (p. ej. Streamlit Cloud / Community) y conectada a un repositorio.

Tabla de contenidos
Características principales
Requisitos
Instalación
Uso
Carga de datos
Ejecución y reporte
Descargas
Formato del archivo de entrada (Excel)
Estructura del código
Despliegue en Github / Streamlit Cloud
Notas y consideraciones
Contribuciones
Licencia
Características principales
Lectura de Excel donde:
Filas = conceptos: Ventas Netas, Utilidad Neta, Activo Total, Capital Contable.
Columnas = periodos (años/meses).
Cálculo pormenorizado del DuPont por periodo:
Margen de utilidad neta
Rotación de activos
Apalancamiento (ratio de activos sobre capital contable)
ROE (combinación de los tres factores)
Soporte para múltiples periodos sin necesidad de reconfiguración.
Informe interactivo en Streamlit y opciones de descarga:
CSV
Excel
Preparado para desplegarse desde Github y/o Streamlit Cloud.
Requisitos
Python 3.8 o superior
Streamlit
pandas
openpyxl (para leer/escribir Excel)
Opcional:

numpy (para cálculos numéricos, si se desea)
Instalación
1) Clona este repositorio o crea un nuevo proyecto y agrega los archivos necesarios.

2) Crea un entorno virtual (recomendado):

Windows:
python -m venv venv
venv\Scripts\activate
macOS/Linux:
python3 -m venv venv
source venv/bin/activate
3) Instala las dependencias:

pip install streamlit pandas openpyxl
4) Verifica la instalación ejecutando la app localmente (ver sección Uso).

Uso
1) Ejecuta la app localmente:

streamlit run main_dupont_streamlit.py
Abre el enlace que aparece en la consola (por defecto http://localhost:8501).
2) Interfaz de usuario:

Carga tu archivo Excel (.xlsx, .xlsm) desde la UI.
La app detecta los periodos a partir de las columnas (años/meses) y genera el reporte por periodo.
El reporte muestra, por cada periodo y para cada empresa (si aplica), las columnas:
Ventas Netas
Utilidad Neta
Activo Total
Capital Contable
Margen_utilidad_neta
Rotacion_activos
Apalancamiento
ROE
3) Descargas:

CSV: descarga el reporte completo en formato CSV.
Excel: descarga el reporte en formato Excel (.xlsx).
Formato del archivo de entrada (Excel)
Estructura esperada:
Filas: conceptos contables (renglones) con exactamente los nombres:
Ventas Netas
Utilidad Neta
Activo Total
Capital Contable
Columnas: periodos (años o meses). Cada periodo debe estar en una columna separada.
El archivo puede contener una fila/columna adicional para identificar la empresa si se incluye; si no, la app generará nombres por defecto (Empresa_1, Empresa_2, etc.) por fila.
Si tu archivo tiene una columna para Empresa, la app intentará usarla como etiqueta de cada fila. Si no, se etiquetará automáticamente.
Notas:

Si necesitas adaptar el parser a un formato ligeramente distinto, envíame un ejemplo de cabecera/estructura y ajusto el código.
Estructura del código
main_dupont_streamlit.py

duPont_triple(util_neta, ventas, activos_totales, capital_contable): cálculo del Tríple DuPont.
cargar_datos_desde_excel(uploaded_file): lee el Excel desde la subida de Streamlit.
extraer_periodos(df): identifica los periodos a partir de las columnas (años/meses).
base_conceptos(): devuelve la lista de conceptos en renglones.
formato_resultados(df, periodos, conceptos): genera el DataFrame final con cálculos por periodo.
main(): flujo principal de la app con interfaz de usuario.
Requisitos de salida:

df_result: DataFrame con resultados por periodo y, si aplica, por empresa.
Despliegue en Github / Streamlit Cloud
Github:
Coloca el código en un repositorio.
Incluye un requirements.txt opcional con:
streamlit
pandas
openpyxl
Streamlit Cloud (Streamlit Community):
Vincula tu repositorio en la plataforma.
Especifica el archivo de entrada: main_dupont_streamlit.py.
Despliega y comparte la URL pública.
Despliegue alternativo (CI/CD):
Configura un flujo que ejecute streamlit run main_dupont_streamlit.py al hacer push en la rama principal.
Opcional: habilita despliegues automáticos con GitHub Actions.
Notas y consideraciones
Robustez: el código asume que el archivo Excel utiliza renglones para conceptos (“Ventas Netas”, “Utilidad Neta”, “Activo Total”, “Capital Contable”) y columnas para periodos (años/meses). Si tu formato varía, ajusta las claves o indica el formato para adaptar el parser.
Validaciones: se omite un periodo si faltan datos para las cuatro columnas necesarias en ese periodo y fila.
Extensiones posibles:
Soporte para múltiples hojas (trabajar con varias pestañas).
Formatos de reporte adicionales (resúmenes por empresa, gráficos, tablas dinámicas).
Exportes a PDF o formatos personalizados.
Contribuciones
Si deseas colaborar, por favor:
Abre un issue para describir el cambio.
Crea una rama nueva, implementa y envía un pull request.
Incluye tests básicos si es posible.
Licencia
Este proyecto se distribuye bajo la licencia MIT (o la que corresponda a tu proyecto). Revisa el archivo LICENSE para detalles.
