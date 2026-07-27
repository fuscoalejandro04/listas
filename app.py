import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
import io

# Configuración de la página
st.set_page_config(page_title="Limpieza de Listas de Precios", layout="wide")
st.title("🧹 Limpiador de Listas de Precios desde Google Sheets")
st.markdown("Sube el enlace de un Google Sheets y obtén un Excel unificado y limpio.")

# --- Entrada del enlace ---
st.header("📎 Paso 1: Ingresa el enlace de tu Google Sheets")
sheet_url = st.text_input(
    "Enlace del Google Sheets",
    placeholder="https://docs.google.com/spreadsheets/d/.../edit?usp=sharing",
    help="Pega el enlace que obtienes al compartir tu archivo de Google Sheets."
)

# --- Botón de proceso ---
if st.button("🚀 Procesar y unificar hojas"):
    if not sheet_url:
        st.warning("Por favor, ingresa un enlace válido.")
    else:
        with st.spinner("Conectando a Google Sheets y leyendo todas las hojas..."):
            try:
                # 1. Establecer conexión usando st.connection
                conn = st.connection("gsheets", type=GSheetsConnection)

                # 2. Obtener el objeto spreadsheet (compatible con múltiples versiones)
                #    Algunas versiones usan '_spreadsheet', otras 'spreadsheet' o 'client'
                if hasattr(conn, '_spreadsheet'):
                    spreadsheet = conn._spreadsheet
                elif hasattr(conn, 'spreadsheet'):
                    spreadsheet = conn.spreadsheet
                elif hasattr(conn, 'client'):
                    # Alternativa: usar el cliente directamente
                    client = conn.client
                    spreadsheet = client.open_by_url(sheet_url)
                else:
                    raise AttributeError(
                        "No se pudo acceder al objeto spreadsheet. "
                        "Asegúrate de tener configuradas las credenciales correctamente."
                    )

                # 3. Obtener todas las hojas
                worksheets = spreadsheet.worksheets()
                st.info(f"📄 Se encontraron {len(worksheets)} hojas en el archivo.")

                all_data = []

                # 4. Recorrer cada hoja
                for sheet in worksheets:
                    sheet_name = sheet.title
                    st.write(f"🔄 Leyendo hoja: **{sheet_name}**")

                    # Obtener todos los valores (texto/números, las imágenes se ignoran)
                    data = sheet.get_all_values()
                    if not data:
                        st.warning(f"La hoja '{sheet_name}' está vacía. Se omite.")
                        continue

                    # Convertir a DataFrame (primera fila = encabezados)
                    df = pd.DataFrame(data[1:], columns=data[0])

                    # Limpieza básica: eliminar filas y columnas totalmente vacías
                    df = df.dropna(how='all')
                    df = df.dropna(axis=1, how='all')

                    # Agregar columna con el nombre de la hoja de origen
                    df['hoja_origen'] = sheet_name

                    all_data.append(df)

                if not all_data:
                    st.error("No se pudo leer ninguna hoja con datos.")
                    st.stop()

                # 5. Unificar todos los DataFrames
                df_unificado = pd.concat(all_data, ignore_index=True)
                st.success(f"✅ Se unificaron {len(all_data)} hojas. Total de filas: {len(df_unificado)}")

                # 6. Mapeo de columnas (ajusta según tus nombres reales)
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

                # Renombrar columnas que coincidan
                for objetivo, posibles in mapeo_columnas.items():
                    for col in df_unificado.columns:
                        if col.lower() in [p.lower() for p in posibles]:
                            df_unificado.rename(columns={col: objetivo}, inplace=True)
                            break

                # Crear columnas faltantes (vacías)
                for col in columnas_objetivo:
                    if col not in df_unificado.columns:
                        df_unificado[col] = ""

                # Reordenar según el orden deseado
                df_unificado = df_unificado[columnas_objetivo]

                # 7. Vista previa
                st.subheader("👀 Vista previa de los datos unificados")
                st.dataframe(df_unificado.head(20))

                # 8. Generar archivo Excel para descargar
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_unificado.to_excel(writer, index=False, sheet_name='Unificado')

                st.download_button(
                    label="⬇️ Descargar Excel (sin imágenes)",
                    data=output.getvalue(),
                    file_name="lista_precios_limpia.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.balloons()

            except Exception as e:
                st.error(f"❌ Ocurrió un error: {e}")
                st.stop()

# --- Ayuda ---
st.markdown("---")
st.markdown("""
### 📌 Notas importantes:
- El enlace debe ser de un Google Sheets **compartido** (público o con la cuenta de servicio configurada).
- Las imágenes y formatos especiales se **ignoran** automáticamente (solo se leen los valores de texto/números).
- La columna `hoja_origen` te indica de qué pestaña vino cada fila.
- Si tus columnas tienen nombres diferentes, la app intentará mapearlas automáticamente.
""")
