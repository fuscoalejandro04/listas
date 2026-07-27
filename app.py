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

        # Propagar valores hacia abajo (ffill) para las columnas que pueden estar combinadas
        columnas_a_propagar = ['Nombre', 'Descripcion', 'Color', 'Presentacion']
        for col in columnas_a_propagar:
            if col in df_clean.columns:
                # Convertir a string para detectar vacíos y espacios
                df_clean[col] = df_clean[col].astype(str)
                # Reemplazar cadenas vacías o con solo espacios por NaN
                df_clean[col] = df_clean[col].replace(r'^\s*$', pd.NA, regex=True)
                # Propagación hacia abajo
                df_clean[col] = df_clean[col].ffill()

        # Limpiar precio (no se propaga, se toma el valor de cada fila)
        df_clean['PrecioLista'] = df_clean['PrecioLista'].apply(limpiar_precio)

        # Eliminar filas sin Artículo
        df_clean = df_clean.dropna(subset=['Artículo'])
        df_clean = df_clean[df_clean['Artículo'].astype(str).str.strip() != '']

        df_clean['Hoja_Origen'] = sheet
        df_list.append(df_clean)

    return pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()
