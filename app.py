import streamlit as st
import pandas as pd
import io
import requests
import re
import warnings
from difflib import get_close_matches

warnings.filterwarnings('ignore')

st.set_page_config(page_title="Estandarizador Universal de Listas de Precios", layout="wide")
st.title("📊 Estandarizador Universal de Listas de Precios")
st.markdown("Pega el enlace de un Google Sheets **público** y obtén un Excel con columnas estandarizadas.")

# -------------------- Configuración de la taxonomía --------------------
# Diccionario de mapeo: columna_objetivo -> lista de sinónimos (en minúsculas, sin tildes)
# Para simplificar, eliminamos tildes y convertimos a minúsculas en la comparación.
TAXONOMIA = {
    'marca': ['marca', 'fabricante', 'proveedor', 'brand'],
    'codigo': ['codigo', 'código', 'id', 'sku', 'articulo', 'artículo', 'item'],
    'modelo': ['modelo', 'model', 'referencia', 'ref'],
    'categoria': ['categoria', 'categoría', 'tipo', 'rubro', 'seccion', 'sección', 'clasificación'],
    'nombre_articulo': ['nombre', 'descripcion corta', 'descripción corta', 'articulo', 'artículo', 'producto', 'denominación'],
    'descripcion': ['descripcion', 'descripción', 'detalle', 'especificación', 'características'],
    'iva': ['iva', 'i.v.a.', 'alicuota', 'alícuota', 'tasa', 'impuesto', '%iva'],
    'precio_lista': ['precio lista', 'precio de lista', 'costo', 'neto', 'precio neto'],
    'precio_sugerido': ['precio sugerido', 'pvp', 'precio con iva', 'sugerido', 'precio público'],
    'precio_2': ['precio 2', 'precio cuotas', 'otro precio', 'precio promocional', 'precio2'],
    'precio_3': ['precio 3', 'tercer precio', 'precio3'],
    'color': ['color', 'colour', 'tono'],
    'tamaño': ['tamaño', 'presentacion', 'presentación', 'volumen', 'capacidad', 'ml', 'l', 'kg', 'mm', 'cm'],
    'embalaje': ['embalaje', 'empaque', 'envase', 'caja', 'bolsa', 'granel', 'tipo embalaje'],
    'cantidad_caja': ['cantidad por caja', 'unidades por caja', 'cant. caja', 'unidades caja'],
    'unidad_precio': ['unidad de precio', 'unidad', 'base', 'precio por'],
    'ean': ['ean', 'codigo de barras', 'código de barras', 'bar code', 'upc'],
}

# Columnas que se propagarán hacia abajo (ffill) para manejar celdas combinadas
COLUMNAS_FFILL = ['marca', 'categoria', 'nombre_articulo', 'descripcion', 'color', 'tamaño', 'embalaje']

# -------------------- Funciones auxiliares --------------------
def extraer_id_documento(url):
    match = re.search(r'/d/([a-zA-Z0-9_-]+)', url)
    return match.group(1) if match else None

def descargar_excel_desde_url(url):
    doc_id = extraer_id_documento(url)
    if not doc_id:
        return None
    export_url = f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx"
    response = requests.get(export_url)
    if response.status_code != 200:
        return None
    return io.BytesIO(response.content)

def limpiar_numero(valor):
    """Convierte a float, manejando comas, puntos y símbolos."""
    if pd.isna(valor) or valor == '':
        return 0.0
    v = str(valor).strip()
    # Reemplazar coma por punto (para decimales)
    v = v.replace(',', '.')
    # Eliminar todo excepto dígitos, punto y signo menos
    v = re.sub(r'[^\d.\-]', '', v)
    try:
        return float(v)
    except:
        return 0.0

def limpiar_iva(valor):
    """Convierte a decimal: '21%' -> 0.21, '0.21' -> 0.21, etc."""
    if pd.isna(valor) or valor == '':
        return 0.21  # valor por defecto
    v = str(valor).strip().replace(',', '.')
    if '%' in v:
        try:
            return float(v.replace('%', '')) / 100
        except:
            return 0.21
    try:
        f = float(v)
        # Si es > 1, asumimos que es porcentaje (ej: 21 -> 0.21)
        if f > 1:
            return f / 100
        return f
    except:
        return 0.21

def normalizar_texto(texto):
    """Quita tildes, espacios extra y convierte a minúsculas para comparación."""
    if not isinstance(texto, str):
        return ''
    # Reemplazar tildes comunes
    import unicodedata
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    texto = texto.lower().strip()
    texto = re.sub(r'\s+', ' ', texto)  # múltiples espacios a uno
    return texto

def detectar_fila_encabezados(df_raw, umbral=2):
    """
    Devuelve el índice de la fila que probablemente contiene los encabezados.
    Busca coincidencias con las palabras clave de las columnas objetivo.
    """
    # Crear lista de todas las palabras clave (sin repetir)
    todas_palabras = set()
    for sinonimos in TAXONOMIA.values():
        todas_palabras.update(sinonimos)
    # Convertir a minúsculas y sin tildes
    palabras_clave = [normalizar_texto(p) for p in todas_palabras]

    for i in range(min(20, len(df_raw))):
        row = df_raw.iloc[i]
        # Unir todas las celdas de la fila en un solo string
        texto_fila = " ".join([str(cell).lower() for cell in row if pd.notna(cell)])
        # Contar cuántas palabras clave aparecen
        coincidencias = sum(1 for palabra in palabras_clave if palabra in texto_fila)
        if coincidencias >= umbral:
            return i
    return None  # Si no encuentra, usa la fila 0

def mapear_columnas(header_row, df_columns):
    """
    header_row: lista de nombres de columnas (strings)
    df_columns: lista de nombres originales de columnas (para renombrar)
    Devuelve un diccionario: {columna_objetivo: nombre_original_de_columna}
    """
    mapeo = {}
    # Normalizar cada nombre de columna
    for idx, col_name in enumerate(header_row):
        col_norm = normalizar_texto(col_name)
        if col_norm == '':
            continue
        mejor_match = None
        mejor_puntaje = 0
        for objetivo, sinonimos in TAXONOMIA.items():
            for sin in sinonimos:
                sin_norm = normalizar_texto(sin)
                # Coincidencia exacta o si el nombre contiene el sinónimo
                if sin_norm == col_norm or sin_norm in col_norm or col_norm in sin_norm:
                    # Puntaje: priorizar coincidencia exacta o que sea una palabra completa
                    if sin_norm == col_norm:
                        puntaje = 3
                    elif sin_norm in col_norm:
                        puntaje = 2
                    else:
                        puntaje = 1
                    if puntaje > mejor_puntaje:
                        mejor_puntaje = puntaje
                        mejor_match = objetivo
        if mejor_match:
            mapeo[mejor_match] = col_name
    return maeo

# -------------------- Procesamiento principal --------------------
def procesar_excel(excel_bytes):
    xls = pd.ExcelFile(excel_bytes)
    sheet_names = xls.sheet_names
    all_dfs = []

    for sheet_name in sheet_names:
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        if df_raw.empty:
            continue

        # 1. Detectar fila de encabezados
        header_idx = detectar_fila_encabezados(df_raw)
        if header_idx is None:
            # Si no detecta, usar la fila 0
            header_idx = 0

        # 2. Obtener la fila de encabezados y el resto de los datos
        header_row = [str(c).strip() for c in df_raw.iloc[header_idx].values]
        # Limpiar nombres de columna: eliminar saltos de línea, espacios extra
        header_row = [re.sub(r'\s+', ' ', c) for c in header_row]
        # Reemplazar celdas vacías por 'Unnamed'
        header_row = [c if c != '' and c != 'nan' and c != 'None' else f'Unnamed_{i}' for i, c in enumerate(header_row)]

        df_data = df_raw.iloc[header_idx+1:].copy()
        df_data.columns = header_row

        # 3. Mapear columnas
        mapeo = mapear_columnas(header_row, df_data.columns)

        # 4. Seleccionar solo las columnas que se mapearon
        # Para las columnas objetivo que no se mapearon, las creamos vacías
        df_clean = pd.DataFrame()
        for objetivo in TAXONOMIA.keys():
            if objetivo in mapeo:
                col_orig = mapeo[objetivo]
                df_clean[objetivo] = df_data[col_orig]
            else:
                df_clean[objetivo] = pd.NA

        # 5. Limpieza de valores
        # Precios: convertir a número
        for col in ['precio_lista', 'precio_sugerido', 'precio_2', 'precio_3']:
            if col in df_clean.columns:
                df_clean[col] = df_clean[col].apply(limpiar_numero)

        # IVA: convertir a decimal
        if 'iva' in df_clean.columns:
            df_clean['iva'] = df_clean['iva'].apply(limpiar_iva)
        else:
            df_clean['iva'] = 0.21  # valor por defecto

        # 6. Propagar valores hacia abajo (ffill) para columnas que suelen tener celdas combinadas
        for col in COLUMNAS_FFILL:
            if col in df_clean.columns:
                df_clean[col] = df_clean[col].replace(r'^\s*$', pd.NA, regex=True)
                df_clean[col] = df_clean[col].ffill()

        # 7. Eliminar filas donde todas las columnas principales estén vacías (excepto hoja_origen)
        # Definimos las columnas que consideramos "esenciales" para mantener la fila
        columnas_esenciales = ['codigo', 'nombre_articulo', 'descripcion', 'precio_lista']
        # Si al menos una de estas tiene valor, la conservamos
        mask = df_clean[columnas_esenciales].notna().any(axis=1)
        df_clean = df_clean[mask]

        # 8. Añadir hoja de origen
        df_clean['hoja_origen'] = sheet_name

        # 9. Reordenar columnas según el orden de TAXONOMIA + hoja_origen
        columnas_orden = list(TAXONOMIA.keys()) + ['hoja_origen']
        # Asegurar que todas existan
        for col in columnas_orden:
            if col not in df_clean.columns:
                df_clean[col] = pd.NA
        df_clean = df_clean[columnas_orden]

        all_dfs.append(df_clean)

    # Unificar todos los DataFrames
    if all_dfs:
        df_final = pd.concat(all_dfs, ignore_index=True)
    else:
        df_final = pd.DataFrame(columns=list(TAXONOMIA.keys()) + ['hoja_origen'])

    return df_final

# -------------------- Interfaz de Streamlit --------------------
sheet_url = st.text_input(
    "📎 Enlace del Google Sheets (público)",
    placeholder="https://docs.google.com/spreadsheets/d/.../edit?usp=sharing",
    help="El archivo debe estar compartido con 'Cualquiera que tenga el enlace' como visor."
)

if st.button("🚀 Estandarizar y procesar", use_container_width=True):
    if not sheet_url:
        st.warning("Por favor, ingresa un enlace válido.")
    else:
        with st.spinner("Descargando, analizando y estandarizando..."):
            try:
                excel_bytes = descargar_excel_desde_url(sheet_url)
                if excel_bytes is None:
                    st.error("No se pudo descargar el archivo. Verifica el enlace y que sea público.")
                    st.stop()

                df_estandar = procesar_excel(excel_bytes)

                if df_estandar.empty:
                    st.warning("No se pudo extraer información del archivo.")
                else:
                    st.success(f"✅ Procesado. Se generaron {len(df_estandar)} filas estandarizadas.")

                    # Mostrar vista previa
                    st.subheader("👀 Vista previa de los datos estandarizados")
                    st.dataframe(df_estandar.head(20))

                    # Generar archivo Excel para descargar
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_estandar.to_excel(writer, index=False, sheet_name='Estandarizado')

                    st.download_button(
                        label="⬇️ Descargar Excel estandarizado",
                        data=output.getvalue(),
                        file_name="Lista_Estandarizada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"❌ Ocurrió un error: {e}")
                st.stop()

# -------------------- Instrucciones --------------------
st.markdown("---")
st.markdown("""
### 📌 ¿Cómo funciona?

1. Pega el enlace de un Google Sheets **compartido públicamente**.
2. La app **analiza automáticamente** todas las hojas, detecta los encabezados y mapea las columnas a un conjunto estándar.
3. Los datos se **limpian** (precios a números, IVA a decimal, celdas combinadas se propagan).
4. Obtienes un Excel con las siguientes columnas estandarizadas:
   - `marca`, `codigo`, `modelo`, `categoria`, `nombre_articulo`, `descripcion`, `iva`, `precio_lista`, `precio_sugerido`, `precio_2`, `precio_3`, `color`, `tamaño`, `embalaje`, `cantidad_caja`, `unidad_precio`, `ean`, `hoja_origen`.

### 💡 Consejos

- Si alguna columna no se detecta, puedes ampliar el diccionario de sinónimos en el código (variable `TAXONOMIA`).
- La app está diseñada para ser **flexible** y funcionar con la mayoría de las listas de precios comunes.
""")
