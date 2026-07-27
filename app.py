import streamlit as st
import pandas as pd
import io
import requests
import re

# Configuración de la página
st.set_page_config(page_title="Limpieza de Listas de Precios", layout="wide")
st.title("🧹 Limpiador de Listas de Precios desde Google Sheets")
st.markdown("Sube el enlace de un Google Sheets **público** y obtén un Excel unificado y limpio.")

# --- Entrada del enlace ---
st.header("📎 Paso 1: Ingresa el enlace de tu Google Sheets")
sheet_url = st.text_input(
    "Enlace del Google Sheets (compartido públicamente)",
    placeholder="https://docs.google.com/spreadsheets/d/.../edit?usp=sharing",
    help="El archivo debe estar compartido con 'Cualquiera que tenga el enlace' como visor."
)

# --- Botón de proceso ---
if st.button("🚀 Procesar y unificar hojas"):
    if not sheet_url:
        st.warning("Por favor, ingresa un enlace válido.")
    else:
        with st.spinner("Descargando y procesando el archivo..."):
            try:
                # 1. Extraer el ID del documento
                match = re.search(r'/d/([a-zA-Z0-9_-]+)', sheet_url)
                if not match:
                    st.error("No se pudo extraer el ID del documento del enlace proporcionado.")
                    st.stop()
                doc_id = match.group(1)

                # 2. Construir la URL de exportación a Excel (todas las hojas)
                export_url = f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx"

                # 3. Descargar el archivo
                response = requests.get(export_url)
                if response.status_code != 200:
                    st.error(f"No se pudo descargar el archivo. Código de estado: {response.status_code}")
                    st.stop()

                excel_data = io.BytesIO(response.content)

                # 4. Leer todas las hojas
                xls = pd.ExcelFile(excel_data)
                sheet_names = xls.sheet_names
                st.info(f"📄 Se encontraron {len(sheet_names)} hojas en el archivo.")

                all_data = []
                for sheet_name in sheet_names:
                    st.write(f"🔄 Leyendo hoja: **{sheet_name}**")
                    # Leer con header=0 (primera fila como nombres de columna)
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=0)

                    # Limpieza básica: eliminar filas y columnas totalmente vacías
                    df = df.dropna(how='all')
                    df = df.dropna(axis=1, how='all')

                    if df.empty:
                        st.warning(f"La hoja '{sheet_name}' está vacía después de limpiar. Se omite.")
                        continue

                    # Añadir columna con el nombre de la hoja
                    df['hoja_origen'] = sheet_name
                    all_data.append(df)

                if not all_data:
                    st.error("No se pudo leer ninguna hoja con datos.")
                    st.stop()

                # 5. Unificar todos los DataFrames
                df_unificado = pd.concat(all_data, ignore_index=True)
                st.success(f"✅ Se unificaron {len(all_data)} hojas. Total de filas: {len(df_unificado)}")

                # 6. Mapeo de columnas (ahora con conversión a string)
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

                # Renombrar columnas que coincidan (manejando nombres numéricos)
                columnas_renombradas = {}
                for objetivo, posibles in mapeo_columnas.items():
                    for col in df_unificado.columns:
                        col_str = str(col)  # Convertir a string para evitar error con int
                        if col_str.lower() in [p.lower() for p in posibles]:
                            columnas_renombradas[col] = objetivo
                            break

                # Aplicar renombramiento
                df_unificado.rename(columns=columnas_renombradas, inplace=True)

                # Crear columnas faltantes (vacías)
                for col in columnas_objetivo:
                    if col not in df_unificado.columns:
                        df_unificado[col] = ""

                # Reordenar según el orden deseado (solo las columnas que existen)
                columnas_finales = [c for c in columnas_objetivo if c in df_unificado.columns]
                df_unificado = df_unificado[columnas_finales]

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
- El enlace debe ser de un Google Sheets **compartido públicamente** (cualquiera con el enlace puede verlo).
- Las imágenes y formatos especiales se **ignoran** automáticamente (solo se leen los valores de texto/números).
- La columna `hoja_origen` te indica de qué pestaña vino cada fila.
- Si tus columnas tienen nombres diferentes, la app intentará mapearlas automáticamente.
""")
