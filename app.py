import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
import io

# Título de la app
st.set_page_config(page_title="Limpieza de Listas de Precios", layout="wide")
st.title("🧹 Limpiador de Listas de Precios desde Google Sheets")
st.markdown("Sube el enlace de un Google Sheets y obtén un Excel unificado y limpio.")

# --- 1. Entrada del usuario: enlace del Google Sheets ---
st.header("📎 Paso 1: Ingresa el enlace de tu Google Sheets")

# El usuario pega el enlace
sheet_url = st.text_input(
    "Enlace del Google Sheets",
    placeholder="https://docs.google.com/spreadsheets/d/.../edit?usp=sharing",
    help="Pega el enlace que obtienes al compartir tu archivo de Google Sheets."
)

# --- 2. Botón para procesar ---
if st.button("🚀 Procesar y unificar hojas"):
    if not sheet_url:
        st.warning("Por favor, ingresa un enlace válido.")
    else:
        with st.spinner("Conectando a Google Sheets y leyendo todas las hojas..."):
            try:
                # --- 3. Conectar a Google Sheets ---
                # Usamos st.connection con el tipo GSheetsConnection
                # Esto lee el archivo desde el enlace proporcionado
                conn = st.connection("gsheets", type=GSheetsConnection)
                
                # --- 4. Obtener lista de todas las hojas ---
                # Primero, obtenemos el objeto spreadsheet para listar las hojas
                # (Nota: la librería no tiene un método directo para listar hojas,
                #  pero podemos leer el metadata con el cliente subyacente)
                # Usamos el cliente gspread que está debajo
                client = conn._client  # Acceso al cliente gspread
                spreadsheet = client.open_by_url(sheet_url)
                worksheets = spreadsheet.worksheets()
                
                st.info(f"📄 Se encontraron {len(worksheets)} hojas en el archivo.")
                
                # --- 5. Leer cada hoja y unificar ---
                all_data = []  # Lista para ir guardando cada DataFrame
                
                for sheet in worksheets:
                    sheet_name = sheet.title
                    st.write(f"🔄 Leyendo hoja: **{sheet_name}**")
                    
                    # Leer la hoja actual como DataFrame
                    # Usamos conn.read() especificando el nombre de la hoja
                    # Pero conn.read() no acepta el nombre directamente con este conector.
                    # Alternativa: usar gspread para leer la hoja específica
                    # Vamos a leer con pandas desde gspread
                    data = sheet.get_all_values()  # Lista de listas
                    if not data:
                        st.warning(f"La hoja '{sheet_name}' está vacía. Se omite.")
                        continue
                    
                    # Convertir a DataFrame (la primera fila se toma como encabezados)
                    df = pd.DataFrame(data[1:], columns=data[0])
                    
                    # --- 6. Limpieza básica ---
                    # Eliminar filas completamente vacías
                    df = df.dropna(how='all')
                    # Eliminar columnas completamente vacías
                    df = df.dropna(axis=1, how='all')
                    
                    # --- 7. Añadir columna con el nombre de la hoja ---
                    df['hoja_origen'] = sheet_name
                    
                    # Guardar en la lista
                    all_data.append(df)
                
                # --- 8. Unificar todos los DataFrames ---
                if not all_data:
                    st.error("No se pudo leer ninguna hoja con datos.")
                    st.stop()
                
                df_unificado = pd.concat(all_data, ignore_index=True)
                
                st.success(f"✅ Se unificaron {len(all_data)} hojas. Total de filas: {len(df_unificado)}")
                
                # --- 9. Reorganizar columnas según lo solicitado ---
                # Columnas esperadas: codigo, modelo, tipo de herramienta, descripcion, precio de lista, iva, hoja_origen
                # Vamos a mapear las columnas que existan en el DataFrame
                # Nota: los nombres pueden variar, hacemos un mapeo flexible
                
                # Definimos las columnas objetivo
                columnas_objetivo = [
                    'codigo', 'modelo', 'tipo de herramienta', 
                    'descripcion', 'precio de lista', 'iva', 'hoja_origen'
                ]
                
                # Intentamos renombrar columnas si existen con nombres similares
                # (Puedes ajustar este mapeo según cómo vengan tus datos reales)
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
                
                # Verificar qué columnas faltan
                columnas_faltantes = [col for col in columnas_objetivo if col not in df_unificado.columns]
                if columnas_faltantes:
                    st.warning(f"⚠️ No se encontraron estas columnas en los datos: {', '.join(columnas_faltantes)}. Se crearán vacías.")
                    for col in columnas_faltantes:
                        df_unificado[col] = ""
                
                # Reordenar para que las columnas estén en el orden deseado
                # Solo las que existen
                columnas_existentes = [col for col in columnas_objetivo if col in df_unificado.columns]
                df_unificado = df_unificado[columnas_existentes]
                
                # --- 10. Mostrar vista previa ---
                st.subheader("👀 Vista previa de los datos unificados")
                st.dataframe(df_unificado.head(20))
                
                # --- 11. Generar archivo Excel para descargar ---
                st.subheader("📥 Descargar Excel limpio")
                
                # Crear un buffer en memoria para el archivo Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_unificado.to_excel(writer, index=False, sheet_name='Unificado')
                
                # Botón de descarga
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

# --- 12. Instrucciones adicionales ---
st.markdown("---")
st.markdown("""
### 📌 Notas importantes:
- El enlace debe ser de un Google Sheets **compartido** (público o con la cuenta de servicio).
- Las imágenes y formatos especiales se **ignoran** automáticamente (solo se leen los valores de texto/números).
- La columna `hoja_origen` te indica de qué pestaña vino cada fila.
- Si tus columnas tienen nombres diferentes, la app intentará mapearlas automáticamente.
""")
