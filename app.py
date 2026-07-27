import streamlit as st
import pandas as pd
import io
import requests
import re
import warnings
from openpyxl import load_workbook

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

st.set_page_config(page_title="Limpieza de Listas de Precios (Einhell/KWB)", layout="wide")
st.title("🧹 Limpiador de Listas de Precios desde Google Sheets")
st.markdown("Ingresa el enlace de un Google Sheets **público** y obtén dos archivos limpios: **Einhell** y **KWB**, con columna de origen.")

# Función para limpiar el valor de IVA
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

# Función para detectar la fila de encabezados en un DataFrame
def detectar_header(df):
    for i, row in df.head(20).iterrows():
        row_str = " ".join([str(v).upper() for v in row.values])
        if "CÓDIGO" in row_str or "CODIGO" in row_str:
            return i
    return None

# --- Entrada del enlace ---
sheet_url = st.text_input(
    "📎 Enlace del Google Sheets (compartido públicamente)",
    placeholder="https://docs.google.com/spreadsheets/d/.../edit?usp=sharing",
    help="El archivo debe estar compartido con 'Cualquiera que tenga el enlace' como visor."
)

if st.button("🚀 Procesar y unificar"):
    if not sheet_url:
        st.warning("Por favor, ingresa un enlace válido.")
    else:
        with st.spinner("Descargando y procesando el archivo..."):
            try:
                # 1. Extraer el ID del documento
                match = re.search(r'/d/([a-zA-Z0-9_-]+)', sheet_url)
                if not match:
                    st.error("No se pudo extraer el ID del documento.")
                    st.stop()
                doc_id = match.group(1)
                export_url = f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx"

                # 2. Descargar archivo
                response = requests.get(export_url)
                if response.status_code != 200:
                    st.error(f"No se pudo descargar. Código: {response.status_code}")
                    st.stop()

                excel_data = io.BytesIO(response.content)
                xls = pd.ExcelFile(excel_data)
                sheet_names = xls.sheet_names
                st.info(f"📄 Se encontraron {len(sheet_names)} hojas: {', '.join(sheet_names)}")

                # 3. Definir hojas de interés
                einhell_sheets = ['EINHELL ', 'BATERÍAS Y CARGADORES', 'COMBOS EN PROMOCIÓN', 'DISCONTINUOS EINHELL']
                kwb_sheets = ['ACCESORIOS KWB y EINHELL', 'DISCONTINUOS KWB']
                # También procesamos cualquier hoja que contenga "KWB" o "Einhell" como fallback (opcional)
                # Pero usamos las listas exactas.

                df_list_einhell = []
                df_list_kwb = []

                # 4. Procesar cada hoja
                for sheet in sheet_names:
                    if sheet not in einhell_sheets and sheet not in kwb_sheets:
                        continue  # Saltamos hojas que no nos interesan

                    df_raw = pd.read_excel(xls, sheet_name=sheet, header=None)
                    header_idx = detectar_header(df_raw)
                    if header_idx is None:
                        st.warning(f"No se encontró encabezado en la hoja '{sheet}'. Se omite.")
                        continue

                    # Asignar encabezados
                    df_raw.columns = [str(c).strip().upper().replace("\n", " ") for c in df_raw.iloc[header_idx]]
                    df = df_raw.iloc[header_idx+1:].copy()

                    # Buscar columna de código
                    col_cod = [c for c in df.columns if "CÓDIGO" in c or "CODIGO" in c]
                    if not col_cod:
                        st.warning(f"No se encontró columna 'CÓDIGO' en la hoja '{sheet}'. Se omite.")
                        continue
                    col_cod = col_cod[0]

                    # Limpiar filas vacías en código
                    df = df.dropna(subset=[col_cod])
                    # Descartar subheaders repetidos
                    df = df[~df[col_cod].astype(str).str.upper().isin(['CÓDIGO', 'CODIGO', 'NAN', ''])]
                    # Solo números o códigos largos (para Einhell)
                    df = df[df[col_cod].astype(str).str.isnumeric() | (df[col_cod].astype(str).str.len() > 3)]

                    # Determinar columnas según la hoja
                    if sheet in einhell_sheets:
                        # Para Einhell: necesitamos 'HERRAMIENTA', 'MODELO' (o 'COMBO'), 'DESCRIPCIÓN', 'PRECIO DE LISTA', 'IVA'
                        col_herramienta = [c for c in df.columns if "HERRAMIENTA" in c]
                        col_herramienta = col_herramienta[0] if col_herramienta else None

                        col_modelo = [c for c in df.columns if "MODELO" in c or "COMBO" in c]
                        col_modelo = col_modelo[0] if col_modelo else None

                        col_desc = [c for c in df.columns if "DESCRIPCIÓN" in c or "DESCRIPCION" in c]
                        col_desc = col_desc[0] if col_desc else None

                        col_precio = [c for c in df.columns if "PRECIO DE LISTA" in c or "COSTO NETO" in c]
                        col_precio = col_precio[0] if col_precio else None

                        col_iva = [c for c in df.columns if "IVA" in c and "%" in c]
                        col_iva = col_iva[0] if col_iva else None

                        # Seleccionar columnas
                        cols_to_keep = [col_cod]
                        if col_herramienta: cols_to_keep.append(col_herramienta)
                        if col_modelo: cols_to_keep.append(col_modelo)
                        if col_desc: cols_to_keep.append(col_desc)
                        if col_precio: cols_to_keep.append(col_precio)
                        if col_iva: cols_to_keep.append(col_iva)

                        df_clean = df[cols_to_keep].copy()
                        # Renombrar
                        rename_map = {}
                        if col_herramienta: rename_map[col_herramienta] = 'Herramienta'
                        if col_modelo: rename_map[col_modelo] = 'Modelo'
                        if col_desc: rename_map[col_desc] = 'Descripcion'
                        if col_precio: rename_map[col_precio] = 'Precio_Lista'
                        if col_iva: rename_map[col_iva] = 'IVA'
                        # El código ya lo renombraremos después
                        df_clean.rename(columns=rename_map, inplace=True)
                        df_clean.rename(columns={col_cod: 'Codigo'}, inplace=True)

                        # Añadir origen
                        df_clean['Hoja_Origen'] = sheet

                        # Limpiar IVA y precio
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
                        # Para KWB: columnas 'CODIGO', 'DESCRIPCIÓN', 'PRECIO LISTA', 'IVA'
                        col_desc = [c for c in df.columns if "DESCRIPCION" in c or "DESCRIPCIÓN" in c]
                        col_desc = col_desc[0] if col_desc else None

                        col_precio = [c for c in df.columns if "PRECIO LISTA" in c or "PRECIO DE LISTA" in c]
                        col_precio = col_precio[0] if col_precio else None

                        col_iva = [c for c in df.columns if "IVA" in c and "%" in c]
                        col_iva = col_iva[0] if col_iva else None

                        cols_to_keep = [col_cod]
                        if col_desc: cols_to_keep.append(col_desc)
                        if col_precio: cols_to_keep.append(col_precio)
                        if col_iva: cols_to_keep.append(col_iva)

                        df_clean = df[cols_to_keep].copy()
                        rename_map = {}
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

                        df_clean['Marca'] = 'KWB'
                        df_list_kwb.append(df_clean)

                # 5. Unificar por marca
                df_einhell = pd.concat(df_list_einhell, ignore_index=True) if df_list_einhell else pd.DataFrame()
                df_kwb = pd.concat(df_list_kwb, ignore_index=True) if df_list_kwb else pd.DataFrame()

                st.success(f"✅ Procesado: {len(df_einhell)} artículos Einhell y {len(df_kwb)} artículos KWB.")

                # 6. Mostrar vistas previas
                if not df_einhell.empty:
                    st.subheader("👀 Vista previa - Einhell")
                    st.dataframe(df_einhell.head(10))
                if not df_kwb.empty:
                    st.subheader("👀 Vista previa - KWB")
                    st.dataframe(df_kwb.head(10))

                # 7. Generar archivos Excel para descarga
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
                    if not df_einhell.empty:
                        df_einhell.to_excel(writer, index=False, sheet_name='Einhell')
                    if not df_kwb.empty:
                        df_kwb.to_excel(writer, index=False, sheet_name='KWB')

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
                        label="⬇️ Descargar Combinado (ambas marcas)",
                        data=output_combinado.getvalue(),
                        file_name="Lista_Combinada_Limpia.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        disabled=(df_einhell.empty and df_kwb.empty)
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
- La app procesa las hojas: `EINHELL `, `BATERÍAS Y CARGADORES`, `COMBOS EN PROMOCIÓN`, `DISCONTINUOS EINHELL`, `ACCESORIOS KWB y EINHELL`, `DISCONTINUOS KWB`.
- Se añade la columna **`Hoja_Origen`** para saber de qué pestaña proviene cada fila.
- Los resultados se entregan en dos archivos separados por marca, y también uno combinado.
""")
