import streamlit as st
import pandas as pd
import io
import gc
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# --- 1. CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Separador de Actividades", page_icon="🏗️", layout="wide")

# --- 2. TÍTULO E INSTRUCCIONES ---
st.title("🏗️ Separador de Actividades de Obra")

# Panel de instrucciones visible
st.info("""
**📋 INSTRUCCIONES DE USO:**

1.  **Sube tu archivo:** Arrastra tu Excel de control de horas (`.xlsx`).
2.  **Procesamiento:** El sistema detectará automáticamente las columnas `RMMAL`.
3.  **Desglose:** Se separarán las actividades mezcladas en filas individuales.
4.  **Limpieza:** Se eliminarán duplicados de comida (solo 1 comida por persona/día) y se limpiará el formato de fechas.
5.  **Descarga:** Obtendrás un nuevo Excel listo para reportar.
""")

# --- 3. LÓGICA INTELIGENTE (OPTIMIZADA) ---
def procesar_excel_master(file_content):
    try:
        # Usamos BytesIO para no gastar disco
        excel_file = io.BytesIO(file_content)
        
        # A) ESCLARECIMIENTO DE ESTRUCTURA
        # Leemos solo un poco para encontrar dónde empieza la tabla real
        df_raw = pd.read_excel(excel_file, header=None, nrows=50)
        
        fila_encabezado = -1
        for i, fila in df_raw.iterrows():
            texto = fila.astype(str).str.upper()
            if texto.str.contains("CLAVE").any() and texto.str.contains("ASIST").any():
                fila_encabezado = i
                break
                
        if fila_encabezado == -1:
            return None, "❌ No se encontraron los encabezados 'CLAVE' y 'ASIST' en el inicio del archivo."

        # Construcción de nombres de columnas (Fusionando fila 1 y fila 2)
        header_top = df_raw.iloc[fila_encabezado].ffill().astype(str).str.strip()
        header_bottom = df_raw.iloc[fila_encabezado + 1].fillna("").astype(str).str.strip()
        
        nombres_columnas = []
        indices_rmmal = []
        columna_comida = None
        
        for k in range(len(header_top)):
            top = header_top.iloc[k]
            bottom = header_bottom.iloc[k]
            if top == "nan": top = f"Col_{k}"
            if bottom == "nan" or bottom == "x": bottom = ""
            
            nombre_unico = f"{top}|{bottom}"
            nombres_columnas.append(nombre_unico)
            
            # Detectamos columnas clave
            if "RMMAL" in top.upper(): 
                indices_rmmal.append(nombre_unico)
            if "COMIDA" in top.upper() and "TOTAL" not in top.upper(): 
                columna_comida = nombre_unico

        # Limpiamos memoria
        del df_raw
        gc.collect()

        # B) CARGA DE DATOS COMPLETA
        excel_file.seek(0)
        df = pd.read_excel(excel_file, header=None, skiprows=fila_encabezado + 2)
        
        # Ajuste seguro de columnas
        df = df.iloc[:, :len(nombres_columnas)]
        df.columns = nombres_columnas
        
        # Convertir números (float32 para ahorrar memoria RAM)
        for col in indices_rmmal: 
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype('float32')
        if columna_comida: 
            df[columna_comida] = pd.to_numeric(df[columna_comida], errors='coerce').fillna(0).astype('float32')

        # C) EL DESGLOSE (SEPARACIÓN DE FILAS)
        nuevas_filas = []
        records = df.to_dict('records') # Convertir a diccionarios es más rápido
        
        # Buscamos columnas destino
        try:
            col_act_key = [c for c in df.columns if str(c).startswith("Act|")][0]
            col_turno_key = [c for c in df.columns if str(c).startswith("Turno|")][0]
        except:
            return None, "❌ El archivo debe tener columnas vacías con encabezado 'Act' y 'Turno'."

        for row in records:
            # Identificar qué actividades tienen horas en esta fila
            cols_activas = [c for c in indices_rmmal if row[c] > 0]
            
            if not cols_activas: 
                continue # Si no trabajó, no se copia
            
            # Ganadora del renglón (para asignarle la comida inicial)
            columna_ganadora = max(cols_activas, key=lambda k: row[k])
            
            for col_actual in cols_activas:
                fila_nueva = row.copy()
                partes = col_actual.split('|')
                
                # Asignar nombres
                fila_nueva[col_act_key] = partes[0]      # Ej: RMMAL-01.01
                fila_nueva[col_turno_key] = partes[1]    # Ej: 1er
                
                # Borrar las horas de las otras actividades en esta copia
                for c in indices_rmmal:
                    if c != col_actual: fila_nueva[c] = 0
                
                # Lógica Comida (Nivel 1: Por fila)
                if columna_comida and col_actual != columna_ganadora:
                    fila_nueva[columna_comida] = 0
                
                nuevas_filas.append(fila_nueva)
        
        # Liberamos memoria
        del df, records
        gc.collect()
        
        # Creamos el DataFrame desglosado
        df_final = pd.DataFrame(nuevas_filas, columns=nombres_columnas)

        # D) CANDADO FINAL: ELIMINAR DOBLES COMIDAS POR DÍA (NUEVO) 🔒
        if columna_comida:
            # Buscamos columnas de Nombre y Fecha
            c_nombre = next((c for c in df_final.columns if "NOMBRE" in c.upper()), None)
            c_fecha = next((c for c in df_final.columns if "FECHA" in c.upper()), None)
            
            if c_nombre and c_fecha:
                # Ordenamos: Quien tenga más horas queda arriba
                # Calculamos total de horas por fila para ordenar
                df_final['__sum_horas__'] = df_final[indices_rmmal].sum(axis=1)
                df_final = df_final.sort_values(by=[c_nombre, c_fecha, '__sum_horas__'], ascending=[True, True, False])
                
                # Detectamos duplicados de (Nombre + Fecha) y ponemos comida en 0 a los repetidos
                # keep='first' respeta el que tiene más horas (porque ya ordenamos)
                mask_duplicados = df_final.duplicated(subset=[c_nombre, c_fecha], keep='first')
                df_final.loc[mask_duplicados, columna_comida] = 0
                
                # Borramos columna auxiliar
                df_final.drop(columns=['__sum_horas__'], inplace=True)

        # E) LIMPIEZA DE FECHAS
        cols_fecha = [c for c in df_final.columns if "FECHA" in str(c).upper()]
        for col in cols_fecha:
            df_final[col] = pd.to_datetime(df_final[col], errors='coerce').dt.date

        # F) EXPORTACIÓN CON MAQUILLAJE (OPENPYXL)
        output = io.BytesIO()
        
        # Header para Excel
        df_headers = pd.DataFrame([header_top.values, header_bottom.values], columns=nombres_columnas)
        df_export = pd.concat([df_headers, df_final], axis=0)
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, header=False, sheet_name='Reporte')
        
        # Estilos visuales
        output.seek(0)
        wb = load_workbook(output)
        ws = wb.active
        
        fill_gris = PatternFill(start_color="404040", end_color="404040", fill_type="solid")
        font_blanca = Font(color="FFFFFF", bold=True)
        fill_amarillo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        borde = Side(style="thin", color="000000")
        caja = Border(left=borde, right=borde, top=borde, bottom=borde)
        
        # Detectar columna Act para pintar de amarillo
        letra_col_act = None
        for col_idx, cell in enumerate(ws[1], 1):
            if "Act" in str(cell.value): letra_col_act = get_column_letter(col_idx)

        # Aplicar bordes y colores (Optimizado)
        for row in ws.iter_rows():
            for cell in row:
                cell.border = caja
                if cell.row <= 2: # Encabezados
                    cell.fill = fill_gris; cell.font = font_blanca
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif letra_col_act and cell.column_letter == letra_col_act:
                    cell.fill = fill_amarillo # Columna Actividad
                    cell.alignment = Alignment(horizontal="left")
                
                # Centrar fechas
                if cell.value and str(cell.value).startswith("202") and "-" in str(cell.value):
                     cell.alignment = Alignment(horizontal="center")

        # Ajuste de ancho de columnas
        for col in ws.columns:
            try:
                # Medimos solo las primeras 50 filas para no tardar mucho
                max_len = max(len(str(cell.value) or "") for cell in col[:50])
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 50)
            except: pass
            
        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.
