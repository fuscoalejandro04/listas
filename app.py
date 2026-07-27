import streamlit as st
import pandas as pd
import io
import requests
import re
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

st.set_page_config(page_title="Limpieza de Listas de Precios", layout="wide")
st.title("🧹 Limpiador de Listas de Precios")
st.markdown("Selecciona el tipo de lista, ingresa el enlace de Google Sheets y obtén los datos limpios.")

# -------------------- FUNCIONES AUXILIARES --------------------
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

def limpiar_iva(valor):
    v = str(valor).strip().replace(',', '.')
    if 'IVA' in v.upper():
        return 0.21
    if '%' in v:
        try:
            return float(v.replace('%', '')) / 100
        except:
            return 0.21
    try:
        f = float(v)
        return f / 100 if f > 1 else f
    except:
        return 0.21

def limpiar_precio(valor):
    """Convierte a número, maneja puntos y comas, y elimina símbolos."""
    if valor is None or pd.isna(valor) or str(valor).strip() == '':
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    v = str(valor).strip()
    v = v.replace(',', '.')
    v = re.sub(r'[^\d.]', '', v)
    try:
        return float(v)
    except:
        return 0.0

def detectar_header(df):
    for i, row in df.head(20).iterrows():
        row_str = " ".join([str(v).upper() for v in row.values])
        if "CÓDIGO" in row_str or "CODIGO" in row_str or "ARTÍCULO" in row_str or "ARTICULO" in row_str:
            return i
    return None

def generar_nombre_corto(descripcion, max_len=50):
    """
    Genera un nombre corto a partir de la descripción.
    Toma los primeros max_len caracteres, o hasta el primer punto, coma o guión.
    Si la descripción está vacía o es solo 'kwb'/'einhell', devuelve ''.
    """
    if not descripcion or pd.isna(descripcion):
        return ''
    texto = str(descripcion).strip()
    # Si solo dice 'kwb' o 'einhell' (insensible), devolver vacío
    if texto.lower() in ['kwb', 'einhell', '']:
        return ''
    # Intentar cortar en el primer punto, coma, guión o paréntesis
    for sep in ['.', ',', ' -', ' (', '  ']:
        if sep in texto:
            parte = texto.split(sep)[0].strip()
            if len(parte) > 5:  # que tenga al menos 5 caracteres
                texto = parte
                break
    # Si sigue siendo muy largo, truncar
    if len(texto) > max_len:
        texto = texto[:max_len].strip()
        # Si termina en medio de una palabra, cortar en el último espacio
        if ' ' in texto:
            texto = texto[:texto.rfind(' ')]
    return texto

# -------------------- FUNCIONES DE PROCESAMIENTO POR TIPO --------------------

def procesar_einhell(excel_bytes):
    xls = pd.ExcelFile(excel_bytes)
    sheet_names = xls.sheet_names
    einhell_sheets = ['EINHELL ', 'BATERÍAS Y CARGADORES', 'COMBOS EN PROMOCIÓN', 'DISCONTINUOS EINHELL']
    kwb_sheets = ['ACCESORIOS KWB y EINHELL', 'DISCONTINUOS KWB']
    df_list_einhell = []
    df_list_kwb = []

    for sheet in sheet_names:
        if sheet not in einhell_sheets and sheet not in kwb_sheets:
            continue
        df_raw = pd.read_excel(xls, sheet_name=sheet, header=None)
        header_idx = detectar_header(df_raw)
        if header_idx is None:
            continue
        df_raw.columns = [str(c).strip().upper().replace("\n", " ") for c in df_raw.iloc[header_idx]]
        df = df_raw.iloc[header_idx+1:].copy()
        col_cod = next((c for c in df.columns if "CÓDIGO" in c or "CODIGO" in c), None)
        if not col_cod:
            continue
        df = df.dropna(subset=[col_cod])
        df = df[~df[col_cod].astype(str).str.upper().isin(['CÓDIGO', 'CODIGO', 'NAN', ''])]
        df = df[df[col_cod].astype(str).str.isnumeric() | (df[col_cod].astype(str).str.len() > 3)]

        if sheet in einhell_sheets:
            col_herramienta = next((c for c in df.columns if "HERRAMIENTA" in c), None)
            col_modelo = next((c for c in df.columns if "MODELO" in c or "COMBO" in c), None)
            col_desc = next((c for c in df.columns if "DESCRIPCIÓN" in c or "DESCRIPCION" in c), None)
            col_precio = next((c for c in df.columns if "PRECIO DE LISTA" in c or "COSTO NETO" in c), None)
            col_iva = next((c for c in df.columns if "IVA" in c and "%" in c), None)
            cols = [col_cod]
            if col_herramienta: cols.append(col_herramienta)
            if col_modelo: cols.append(col_modelo)
            if col_desc: cols.append(col_desc)
            if col_precio: cols.append(col_precio)
            if col_iva: cols.append(col_iva)
            df_clean = df[cols].copy()
            rename_map = {}
            if col_herramienta: rename_map[col_herramienta] = 'Herramienta'
            if col_modelo: rename_map[col_modelo] = 'Modelo'
            if col_desc: rename_map[col_desc] = 'Descripcion'
            if col_precio: rename_map[col_precio] = 'Precio_Lista'
            if col_iva: rename_map[col_iva] = 'IVA'
            df_clean.rename(columns=rename_map, inplace=True)
            df_clean.rename(columns={col_cod: 'Codigo'}, inplace=True)
            df_clean['Hoja_Origen'] = sheet
            if 'IVA' in df_clean.columns:
                df_clean['IVA'] = df_clean['IVA'].apply(limpiar_iva)
            else:
                df_clean['IVA'] = 0.21
            if 'Precio_Lista' in df_clean.columns:
                df_clean['Precio_Lista'] = pd.to_numeric(df_clean['Precio_Lista'], errors='coerce').fillna(0).round(2)
            else:
                df_clean['Precio_Lista'] = 0
            df_clean['Marca'] = 'Einhell'
            df_list_einhell.append(df_clean)

        elif sheet in kwb_sheets:
            col_desc = next((c for c in df.columns if "DESCRIPCION" in c or "DESCRIPCIÓN" in c), None)
            col_precio = next((c for c in df.columns if "PRECIO LISTA" in c or "PRECIO DE LISTA" in c), None)
            col_iva = next((c for c in df.columns if "IVA" in c and "%" in c), None)
            col_categoria = next((c for c in df.columns if "CATEGORIA" in c or "CATEGORÍA" in c), None)
            
            cols = [col_cod]
            if col_desc: cols.append(col_desc)
            if col_precio: cols.append(col_precio)
            if col_iva: cols.append(col_iva)
            if col_categoria: cols.append(col_categoria)
            
            df_clean = df[cols].copy()
            rename_map = {}
            if col_desc: rename_map[col_desc] = 'Descripcion'
            if col_precio: rename_map[col_precio] = 'Precio_Lista'
            if col_iva: rename_map[col_iva] = 'IVA'
            if col_categoria: rename_map[col_categoria] = 'Categoria'
            df_clean.rename(columns=rename_map, inplace=True)
            df_clean.rename(columns={col_cod: 'Codigo'}, inplace=True)
            
            # ---- Generar nombre corto a partir de la descripción ----
            df_clean['Nombre'] = df_clean['Descripcion'].apply(generar_nombre_corto)
            # Si el nombre quedó vacío, usar el código como respaldo
            df_clean['Nombre'] = df_clean['Nombre'].replace('', pd.NA)
            df_clean['Nombre'] = df_clean['Nombre'].fillna(df_clean['Codigo'])
            
            # ---- Eliminar filas donde la descripción sea vacía o solo 'kwb'/'einhell' ----
            # Pero conservar si el código existe y el precio > 0
            descripciones_invalidas = ['kwb', 'einhell', '']
            df_clean = df_clean[~df_clean['Descripcion'].astype(str).str.lower().isin(descripciones_invalidas) | (df_clean['Precio_Lista'] > 0)]
            
            # Limpiar IVA y precio
            if 'IVA' in df_clean.columns:
                df_clean['IVA'] = df_clean['IVA'].apply(limpiar_iva)
            else:
                df_clean['IVA'] = 0.21
            if 'Precio_Lista' in df_clean.columns:
                df_clean['Precio_Lista'] = pd.to_numeric(df_clean['Precio_Lista'], errors='coerce').fillna(0).round(2)
            else:
                df_clean['Precio_Lista'] = 0
                
            df_clean['Marca'] = 'KWB'
            df_clean['Hoja_Origen'] = sheet
            df_list_kwb.append(df_clean)

    df_einhell = pd.concat(df_list_einhell, ignore_index=True) if df_list_einhell else pd.DataFrame()
    df_kwb = pd.concat(df_list_kwb, ignore_index=True) if df_list_kwb else pd.DataFrame()
    
    # Para KWB, reordenar columnas: Codigo, Nombre, Descripcion, Precio_Lista, IVA, Marca, Hoja_Origen
    if not df_kwb.empty:
        columnas_kwb = ['Codigo', 'Nombre', 'Descripcion', 'Precio_Lista', 'IVA', 'Marca', 'Hoja_Origen']
        # Asegurar que existan todas
        for col in columnas_kwb:
            if col not in df_kwb.columns:
                df_kwb[col] = ''
        df_kwb = df_kwb[columnas_kwb]
    
    return {
        'einhell': df_einhell,
        'kwb': df_kwb,
        'combinado': pd.concat([df_einhell, df_kwb], ignore_index=True) if not (df_einhell.empty and df_kwb.empty) else pd.DataFrame()
    }

def procesar_fijaciones(excel_bytes):
    xls = pd.ExcelFile(excel_bytes)
    sheet_names = xls.sheet_names
    df_list = []
    for sheet in sheet_names:
        df_raw = pd.read_excel(xls, sheet_name=sheet, header=None)
        header_idx = detectar_header(df_raw)
        if header_idx is None:
            continue
        df_raw.columns = [str(c).strip().upper().replace("\n", " ") for c in df_raw.iloc[header_idx]]
        df = df_raw.iloc[header_idx+1:].copy()
        
        col_cod = next((c for c in df.columns if "CÓDIGO" in c or "CODIGO" in c), None)
        col_desc = next((c for c in df.columns if "DESCRIPCION" in c or "DESCRIPCIÓN" in c), None)
        col_cant_caja = next((c for c in df.columns if "CANTIDAD POR CAJA" in c or "CANT" in c), None)
        col_embalaje = next((c for c in df.columns if "EMBALAJE" in c), None)
        col_unidad_precio = next((c for c in df.columns if "UNIDAD DE PRECIO" in c), None)
        col_precio = next((c for c in df.columns if "PRECIO LISTA" in c or "PRECIO DE LISTA" in c), None)
        col_iva = next((c for c in df.columns if "IVA" in c), None)
        
        if not col_cod or not col_desc or not col_precio:
            continue
            
        cols = [col_cod, col_desc]
        if col_cant_caja: cols.append(col_cant_caja)
        if col_embalaje: cols.append(col_embalaje)
        if col_unidad_precio: cols.append(col_unidad_precio)
        if col_precio: cols.append(col_precio)
        if col_iva: cols.append(col_iva)
        
        df_clean = df[cols].copy()
        
        rename_map = {
            col_cod: 'Codigo',
            col_desc: 'Descripcion',
        }
        if col_cant_caja: rename_map[col_cant_caja] = 'CantidadPorCaja'
        if col_embalaje: rename_map[col_embalaje] = 'Embalaje'
        if col_unidad_precio: rename_map[col_unidad_precio] = 'UnidadPrecio'
        if col_precio: rename_map[col_precio] = 'PrecioLista'
        if col_iva: rename_map[col_iva] = 'IVA'
        
        df_clean.rename(columns=rename_map, inplace=True)
        
        # ---- ELIMINAR FILAS CON ENCABEZADOS REPETIDOS ----
        df_clean = df_clean[~df_clean['Codigo'].astype(str).str.upper().isin(['CODIGO', 'CÓDIGO'])]
        df_clean = df_clean.dropna(subset=['Codigo'])
        df_clean = df_clean[df_clean['Codigo'].astype(str).str.strip() != '']
        # Filtrar precios > 0 (elimina filas con precio 0 que son encabezados repetidos)
        df_clean['PrecioLista'] = pd.to_numeric(df_clean['PrecioLista'], errors='coerce').fillna(0).round(2)
        df_clean = df_clean[df_clean['PrecioLista'] > 0]
        
        if 'IVA' in df_clean.columns:
            df_clean['IVA'] = df_clean['IVA'].apply(limpiar_iva)
        else:
            df_clean['IVA'] = 0.21
            
        df_clean['Hoja_Origen'] = sheet
        df_list.append(df_clean)
        
    return pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()

def procesar_penosil(excel_bytes):
    """
    Procesa el archivo Excel para el modo Penosil.
    Busca columnas: Artículo (código), Nombre, Descripción, Color, Presentación, Precio de lista.
    Propaga hacia abajo (ffill) los valores de Nombre, Descripción, Color y Presentación.
    """
    xls = pd.ExcelFile(excel_bytes)
    sheet_names = xls.sheet_names
    df_list = []

    for sheet in sheet_names:
        df_raw = pd.read_excel(xls, sheet_name=sheet, header=None)

        # Detectar fila de encabezado
        header_idx = None
        for i, row in df_raw.head(20).iterrows():
            row_str = " ".join([str(v).upper() for v in row.values])
            if "ARTÍCULO" in row_str or "ARTICULO" in row_str or "NOMBRE" in row_str or "PRECIO DE LISTA" in row_str:
                header_idx = i
                break
        if header_idx is None:
            continue

        # Asignar encabezados
        df_raw.columns = [str(c).strip().upper().replace("\n", " ") for c in df_raw.iloc[header_idx]]
        df = df_raw.iloc[header_idx+1:].copy()

        # Buscar columnas clave
        col_articulo = next((c for c in df.columns if "ARTÍCULO" in c or "ARTICULO" in c), None)
        col_nombre = next((c for c in df.columns if "NOMBRE" in c), None)
        col_desc = next((c for c in df.columns if "DESCRIPCIÓN" in c or "DESCRIPCION" in c), None)
        col_color = next((c for c in df.columns if "COLOR" in c), None)
        col_presentacion = next((c for c in df.columns if "PRESENTACIÓN" in c or "PRESENTACION" in c), None)
        col_precio = next((c for c in df.columns if "PRECIO DE LISTA" in c or "PRECIO LISTA" in c), None)

        if not col_articulo or not col_precio:
            continue

        cols = [col_articulo]
        if col_nombre: cols.append(col_nombre)
        if col_desc: cols.append(col_desc)
        if col_color: cols.append(col_color)
        if col_presentacion: cols.append(col_presentacion)
        if col_precio: cols.append(col_precio)

        df_clean = df[cols].copy()

        rename_map = {
            col_articulo: 'Artículo',
            col_precio: 'PrecioLista',
        }
        if col_nombre: rename_map[col_nombre] = 'Nombre'
        if col_desc: rename_map[col_desc] = 'Descripcion'
        if col_color: rename_map[col_color] = 'Color'
        if col_presentacion: rename_map[col_presentacion] = 'Presentacion'
        df_clean.rename(columns=rename_map, inplace=True)

        # Propagar valores hacia abajo (ffill) para columnas combinadas
        columnas_a_propagar = ['Nombre', 'Descripcion', 'Color', 'Presentacion']
        for col in columnas_a_propagar:
            if col in df_clean.columns:
                df_clean[col] = df_clean[col].astype(str)
                df_clean[col] = df_clean[col].replace(r'^\s*$', pd.NA, regex=True)
                df_clean[col] = df_clean[col].ffill()

        # Limpiar precio (no se propaga, se toma el valor de cada fila)
        df_clean['PrecioLista'] = df_clean['PrecioLista'].apply(limpiar_precio)

        # Eliminar filas sin Artículo
        df_clean = df_clean.dropna(subset=['Artículo'])
        df_clean = df_clean[df_clean['Artículo'].astype(str).str.strip() != '']

        df_clean['Hoja_Origen'] = sheet
        df_list.append(df_clean)

    return pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()

# -------------------- INTERFAZ DE USUARIO --------------------
modo = st.radio(
    "Selecciona el tipo de lista de precios:",
    ("Einhell / KWB (Herramientas)", "Fijaciones (CPSA y similares)", "Penosil (Productos químicos)")
)

st.markdown("---")

sheet_url = st.text_input(
    "📎 Enlace del Google Sheets (público)",
    placeholder="https://docs.google.com/spreadsheets/d/.../edit?usp=sharing",
    help="El archivo debe estar compartido con 'Cualquiera que tenga el enlace' como visor."
)

if st.button("🚀 Procesar", use_container_width=True):
    if not sheet_url:
        st.warning("Por favor, ingresa un enlace válido.")
    else:
        with st.spinner("Descargando y procesando..."):
            try:
                excel_bytes = descargar_excel_desde_url(sheet_url)
                if excel_bytes is None:
                    st.error("No se pudo descargar el archivo. Verifica el enlace y que sea público.")
                    st.stop()

                if modo == "Einhell / KWB (Herramientas)":
                    resultados = procesar_einhell(excel_bytes)
                    st.session_state['resultados'] = resultados
                    st.session_state['modo'] = 'einhell'
                    st.success(f"✅ Procesado: {len(resultados['einhell'])} artículos Einhell y {len(resultados['kwb'])} artículos KWB.")
                elif modo == "Fijaciones (CPSA y similares)":
                    df_fijaciones = procesar_fijaciones(excel_bytes)
                    st.session_state['resultados'] = {'fijaciones': df_fijaciones}
                    st.session_state['modo'] = 'fijaciones'
                    st.success(f"✅ Procesado: {len(df_fijaciones)} artículos de fijaciones.")
                else:  # Penosil
                    df_penosil = procesar_penosil(excel_bytes)
                    st.session_state['resultados'] = {'penosil': df_penosil}
                    st.session_state['modo'] = 'penosil'
                    st.success(f"✅ Procesado: {len(df_penosil)} artículos de Penosil.")

            except Exception as e:
                st.error(f"❌ Ocurrió un error: {e}")
                st.stop()

# -------------------- MOSTRAR RESULTADOS Y BOTONES DE DESCARGA --------------------
if 'resultados' in st.session_state:
    resultados = st.session_state['resultados']
    modo_actual = st.session_state['modo']

    if modo_actual == 'einhell':
        df_einhell = resultados.get('einhell', pd.DataFrame())
        df_kwb = resultados.get('kwb', pd.DataFrame())
        df_combinado = resultados.get('combinado', pd.DataFrame())

        if not df_einhell.empty:
            st.subheader("👀 Vista previa - Einhell")
            st.dataframe(df_einhell.head(10))
        if not df_kwb.empty:
            st.subheader("👀 Vista previa - KWB")
            # Mostrar columnas relevantes: Codigo, Nombre, Descripcion, Precio_Lista, IVA, Hoja_Origen
            st.dataframe(df_kwb[['Codigo', 'Nombre', 'Descripcion', 'Precio_Lista', 'IVA', 'Hoja_Origen']].head(10))

        output_einhell = io.BytesIO()
        output_kwb = io.BytesIO()
        output_combinado = io.BytesIO()

        with pd.ExcelWriter(output_einhell, engine='openpyxl') as writer:
            if not df_einhell.empty:
                df_einhell.to_excel(writer, index=False, sheet_name='Einhell')
        with pd.ExcelWriter(output_kwb, engine='openpyxl') as writer:
            if not df_kwb.empty:
                df_kwb.to_excel(writer, index=False, sheet_name='KWB')
        with pd.ExcelWriter(output_combinado, engine='openpyxl') as writer:
            if not df_combinado.empty:
                df_combinado.to_excel(writer, index=False, sheet_name='Combinado')

        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button(
                label="⬇️ Descargar Einhell_Limpia.xlsx",
                data=output_einhell.getvalue(),
                file_name="Einhell_Limpia.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                disabled=df_einhell.empty
            )
        with col2:
            st.download_button(
                label="⬇️ Descargar KWB_Limpia.xlsx",
                data=output_kwb.getvalue(),
                file_name="KWB_Limpia.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                disabled=df_kwb.empty
            )
        with col3:
            st.download_button(
                label="⬇️ Descargar Combinado.xlsx",
                data=output_combinado.getvalue(),
                file_name="Combinado_Einhell_KWB.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                disabled=df_combinado.empty
            )

    elif modo_actual == 'fijaciones':
        df_fijaciones = resultados.get('fijaciones', pd.DataFrame())
        if not df_fijaciones.empty:
            st.subheader("👀 Vista previa - Fijaciones")
            st.dataframe(df_fijaciones.head(10))
            output_fijaciones = io.BytesIO()
            with pd.ExcelWriter(output_fijaciones, engine='openpyxl') as writer:
                df_fijaciones.to_excel(writer, index=False, sheet_name='Fijaciones')
            st.download_button(
                label="⬇️ Descargar Fijaciones_Limpia.xlsx",
                data=output_fijaciones.getvalue(),
                file_name="Fijaciones_Limpia.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    elif modo_actual == 'penosil':
        df_penosil = resultados.get('penosil', pd.DataFrame())
        if not df_penosil.empty:
            st.subheader("👀 Vista previa - Penosil")
            st.dataframe(df_penosil.head(10))
            output_penosil = io.BytesIO()
            with pd.ExcelWriter(output_penosil, engine='openpyxl') as writer:
                df_penosil.to_excel(writer, index=False, sheet_name='Penosil')
            st.download_button(
                label="⬇️ Descargar Penosil_Limpia.xlsx",
                data=output_penosil.getvalue(),
                file_name="Penosil_Limpia.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    # Botón para limpiar resultados
    if st.button("🔄 Limpiar resultados y procesar otro archivo"):
        del st.session_state['resultados']
        del st.session_state['modo']
        st.experimental_rerun()

# -------------------- INSTRUCCIONES --------------------
st.markdown("---")
st.markdown("""
### 📌 Instrucciones por modo:

- **Einhell / KWB**: 
  - **Einhell**: columnas `Codigo`, `Herramienta`, `Modelo`, `Descripcion`, `Precio_Lista`, `IVA`, `Marca`, `Hoja_Origen`.
  - **KWB**: columnas `Codigo`, `Nombre` (nombre corto generado), `Descripcion`, `Precio_Lista`, `IVA`, `Marca`, `Hoja_Origen`.

- **Fijaciones**: columnas `Codigo`, `Descripcion`, `CantidadPorCaja`, `Embalaje`, `UnidadPrecio`, `PrecioLista`, `IVA`, `Hoja_Origen`.

- **Penosil**: columnas `Artículo`, `Nombre`, `Descripcion`, `Color`, `Presentacion`, `PrecioLista`, `Hoja_Origen`.
""")
