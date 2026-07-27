import streamlit as st
import pandas as pd
import io
import requests
import re

st.set_page_config(page_title="Limpieza de Listas de Precios", layout="wide")
st.title("🧹 Limpiador de Listas de Precios desde Google Sheets")
st.markdown("Sube el enlace de un Google Sheets **público** y obtén un Excel unificado y limpio.")

# --- Funciones auxiliares ---
def detectar_fila_encabezado(df_raw, palabras_clave):
    """
    Busca la primera fila en df_raw (sin encabezados) que contenga al menos una
    de las palabras_clave (en cualquier columna). Devuelve el índice de esa fila.
    Si no encuentra, devuelve 0.
    """
    for i, row in df_raw.iterrows():
        # Convertir todas las celdas a string y unirlas
        texto_fila = " ".join([str(celda).lower() for celda in row.values if pd.notna(celda)])
        for palabra in palabras_clave:
            if palabra in texto_fila:
                return i
    return 0  # Si no encuentra, asume que la primera fila es encabezado

def leer_hoja_con_encabezado(xls, sheet_name, header_row=0):
    """
    Lee una hoja usando header_row como índice de la fila de encabezados.
    Devuelve un DataFrame con los datos a partir de header_row+1.
    """
    # Leer todas las filas sin asumir encabezados
    df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    # Eliminar filas completamente vacías al inicio
    df_raw = df_raw.dropna(how='all')
    if df_raw.empty:
        return None
    # Usar header_row como la fila de encabezados
    if header_row >= len(df_raw):
        return None
    # Tomar la fila de encabezados
    headers = df_raw.iloc[header_row].values
    # Tomar los datos a partir de la siguiente fila
    data = df_raw.iloc[header_row+1:].copy()
    # Asignar encabezados
    data.columns = headers
    # Eliminar filas donde todas las celdas sean NaN
    data = data.dropna(how='all')
    # Eliminar columnas donde todas las celdas sean NaN
    data = data.dropna(axis=1, how='all')
    return data

# --- Interfaz de usuario ---
sheet_url = st.text_input(
    "📎 Enlace del Google Sheets (compartido públicamente)",
    placeholder="https://docs.google.com/spreadsheets/d/.../edit?usp=sharing",
    help="El archivo debe estar compartido con 'Cualquiera que tenga el enlace' como visor."
)

# --- Botón de procesamiento ---
if st.button("🚀 Procesar y unificar hojas"):
    if not sheet_url:
        st.warning("Por favor, ingresa un enlace válido.")
    else:
        with st.spinner("Descargando y procesando el archivo..."):
            try:
                # 1. Extraer ID del documento
                match = re.search(r'/d/([a-zA-Z0-9_-]+)', sheet_url)
                if not match:
                    st.error("No se pudo extraer el ID del documento.")
                    st.stop()
                doc_id = match.group(1)
                export_url = f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx"

                # 2. Descargar archivo
                response = requests.get(export_url)
                if response.status_code != 200:
                    st.error(f"Error al descargar: {response.status_code}")
                    st.stop()
                excel_data = io.BytesIO(response.content)

                # 3. Leer todas las hojas y mostrar preview
                xls = pd.ExcelFile(excel_data)
                sheet_names = xls.sheet_names
                st.info(f"📄 Se encontraron {len(sheet_names)} hojas.")

                # Diccionario para guardar configuraciones por hoja
                configs = {}
                palabras_clave = ['codigo', 'modelo', 'precio', 'descripcion', 'iva', 'tipo', 'categoria']

                st.subheader("🔍 Paso 2: Configuración de cada hoja")
                st.markdown("Para cada hoja, selecciona la **fila que contiene los encabezados** (los nombres de las columnas).")
                st.markdown("La app intentará detectar automáticamente la fila correcta. Puedes ajustarla manualmente.")

                for sheet_name in sheet_names:
                    st.write(f"#### 📑 Hoja: **{sheet_name}**")
                    # Leer las primeras 15 filas sin encabezados para mostrar preview
                    df_preview = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=15)
                    # Mostrar preview con índice de filas (0-based)
                    st.dataframe(df_preview, use_container_width=True)

                    # Detección automática de la fila de encabezado
                    auto_header = detectar_fila_encabezado(df_preview, palabras_clave)
                    st.caption(f"🔎 Detección automática: fila **{auto_header}** (0 = primera fila)")

                    # Selector para que el usuario elija la fila de encabezado
                    max_row = min(15, len(df_preview)-1)
                    if max_row < 0:
                        max_row = 0
                    header_row = st.number_input(
                        f"Fila de encabezado para '{sheet_name}' (0-based)",
                        min_value=0,
                        max_value=max_row,
                        value=auto_header,
                        key=f"header_{sheet_name}"
                    )
                    configs[sheet_name] = int(header_row)
                    st.markdown("---")

                # 4. Leer cada hoja con la configuración elegida
                all_data = []
                for sheet_name in sheet_names:
                    header_row = configs[sheet_name]
                    st.write(f"🔄 Leyendo hoja: **{sheet_name}** con header_row={header_row}")
                    df = leer_hoja_con_encabezado(xls, sheet_name, header_row)
                    if df is None or df.empty:
                        st.warning(f"La hoja '{sheet_name}' no tiene datos después de la limpieza. Se omite.")
                        continue
                    # Añadir columna de origen
                    df['hoja_origen'] = sheet_name
                    all_data.append(df)

                if not all_data:
                    st.error("No se pudo leer ninguna hoja con datos.")
                    st.stop()

                # 5. Unificar
                df_unificado = pd.concat(all_data, ignore_index=True)
                st.success(f"✅ Se unificaron {len(all_data)} hojas. Total de filas: {len(df_unificado)}")

                # 6. Mapeo de columnas (igual que antes, con conversión a string)
                columnas_objetivo = [
                    'codigo', 'modelo', 'tipo de herramienta',
                    'descripcion', 'precio de lista', 'iva', 'hoja_origen'
                ]
                mapeo_columnas = {
                    'codigo': ['codigo', 'código', 'code', 'id'],
                    'modelo': ['modelo', 'model'],
                    'tipo de herramienta': ['tipo', 'tipo de herramienta', 'categoria', 'category'],
                    'descripcion': ['descripcion', 'descripción', 'description', 'nombre'],
                    'precio de lista': ['precio de lista', 'precio', 'price', 'precio lista'],
                    'iva': ['iva', 'I.V.A.', 'tax', 'impuesto']
                }

                # Renombrar
                columnas_renombradas = {}
                for objetivo, posibles in mapeo_columnas.items():
                    for col in df_unificado.columns:
                        col_str = str(col).strip().lower()
                        if col_str in [p.lower() for p in posibles]:
                            columnas_renombradas[col] = objetivo
                            break
                df_unificado.rename(columns=columnas_renombradas, inplace=True)

                # Agregar columnas faltantes
                for col in columnas_objetivo:
                    if col not in df_unificado.columns:
                        df_unificado[col] = ""

                # Reordenar
                columnas_finales = [c for c in columnas_objetivo if c in df_unificado.columns]
                df_unificado = df_unificado[columnas_finales]

                # 7. Vista previa
                st.subheader("👀 Vista previa de los datos unificados")
                st.dataframe(df_unificado.head(20), use_container_width=True)

                # 8. Descarga
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_unificado.to_excel(writer, index=False, sheet_name='Unificado')

                st.download_button(
                    label="⬇️ Descargar Excel limpio",
                    data=output.getvalue(),
                    file_name="lista_precios_limpia.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.balloons()

            except Exception as e:
                st.error(f"❌ Error: {e}")
                st.stop()

# --- Ayuda ---
st.markdown("---")
st.markdown("""
### 📌 Instrucciones:
1. Ingresa el enlace de un Google Sheets **público**.
2. Para cada hoja, **elige la fila donde están los nombres de las columnas** (0 = primera fila).
3. La app unificará todas las hojas en un solo Excel, con las columnas: `codigo`, `modelo`, `tipo de herramienta`, `descripcion`, `precio de lista`, `iva` y `hoja_origen`.
4. Descarga el resultado.
""")
