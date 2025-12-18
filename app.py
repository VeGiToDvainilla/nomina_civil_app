import streamlit as st
import pandas as pd
import io
import gc # Librería para limpiar memoria RAM
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Procesador Nómina Pro", page_icon="🏗️")
st.title("🏗️ Procesador de Nómina (Optimizado)")

# --- FUNCIÓN DE LIMPIEZA DE MEMORIA ---
def limpiar_memoria():
    gc.collect()

# --- FUNCIÓN DE PROCESAMIENTO ---
@st.cache_data(show_spinner=False) # ESTO ES CLAVE: Guarda en caché para no reprocesar
def procesar_excel_optimizado(file_content):
    try:
        # Usamos BytesIO para manejar el archivo en memoria
        excel_file = io.BytesIO(file_content)
        
        # 1. Lectura Ligera (Solo encabezados)
        df_raw = pd.read_excel(excel_file, header=None, nrows=50)
        
        fila_encabezado = -1
        for i, fila in df_raw.iterrows():
            texto = fila.astype(str).str.upper()
            if texto.str.contains("CLAVE").any() and texto.str.contains("ASIST").any():
                fila_encabezado = i
                break
                
        if fila_encabezado == -1:
            return None, "No encontré 'CLAVE' y 'ASIST' en las primeras 50 filas."

        # Recuperar nombres de columnas
        header_top = df_raw.iloc[fila_encabezado].ffill().astype(str).str.strip()
        header_bottom = df_raw.iloc[fila_encabezado + 1].fillna("").astype(str).str.strip()
        
        nombres_columnas = []
        indices_rmmal = []
        columna_comida = None
        
        for k in range(len(header_top)):
            top = header_top.iloc[k]
            bottom = header_bottom.iloc[k]
            if top == "nan": top = ""
            if bottom == "nan" or bottom == "x": bottom = ""
            nombre_unico = f"{top}|{bottom}"
            nombres_columnas.append(nombre_unico)
            if "RMMAL" in top.upper(): indices_rmmal.append(nombre_unico)
            if "COMIDA" in top.upper() and "TOTAL" not in top.upper(): columna_comida = nombre_unico

        # Liberamos memoria de la lectura inicial
        del df_raw
        limpiar_memoria()

        # 2. Lectura de Datos (Optimizada)
        excel_file.seek(0)
        # Leemos SOLO las columnas necesarias si fuera posible, pero aqui leemos todo y limpiamos rapido
        df = pd.read_excel(excel_file, header=None, skiprows=fila_encabezado + 2)
        
        # Ajuste de columnas
        df = df.iloc[:, :len(nombres_columnas)]
        df.columns = nombres_columnas
        
        # Convertir a numérico (Optimizando tipos de datos a float32 para ahorrar RAM)
        for col in indices_rmmal: 
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype('float32')
        if columna_comida: 
            df[columna_comida] = pd.to_numeric(df[columna_comida], errors='coerce').fillna(0).astype('float32')

        # 3. Desglose (Lógica Core)
        nuevas_filas = []
        # Convertimos a diccionario para iterar más rápido y con menos memoria que iterrows
        records = df.to_dict('records')
        
        # Indices de columnas Act y Turno
        cols_list = df.columns.tolist()
        try:
            idx_act_key = [c for c in cols_list if str(c).startswith("Act|")][0]
            idx_turno_key = [c for c in cols_list if str(c).startswith("Turno|")][0]
        except:
             return None, "Faltan columnas Act o Turno"

        for row in records:
            # Detectar activas
            cols_activas = [c for c in indices_rmmal if row[c] > 0]
            if not cols_activas: continue
            
            # Ganadora
            columna_ganadora = max(cols_activas, key=lambda k: row[k])
            
            for col_actual in cols_activas:
                # Copia ligera del diccionario
                fila_nueva = row.copy()
                partes = col_actual.split('|')
                
                fila_nueva[idx_act_key] = partes[0]
                fila_nueva[idx_turno_key] = partes[1]
                
                # Limpieza de otras horas
                for c in indices_rmmal:
                    if c != col_actual: fila_nueva[c] = 0
                
                # Comida
                if columna_comida and col_actual != columna_ganadora:
                    fila_nueva[columna_comida] = 0
                
                nuevas_filas.append(fila_nueva)
        
        # Liberamos el DF original gigante
        del df
        del records
        limpiar_memoria()
        
        df_final = pd.DataFrame(nuevas_filas, columns=nombres_columnas)
        
        # Limpieza Fechas
        cols_fecha = [c for c in df_final.columns if "FECHA" in str(c).upper()]
        for col in cols_fecha:
            df_final[col] = pd.to_datetime(df_final[col], errors='coerce').dt.date

        # 4. Exportación (Sin guardar en disco, todo en RAM eficiente)
        output = io.BytesIO()
        df_headers = pd.DataFrame([header_top.values, header_bottom.values], columns=nombres_columnas)
        df_export = pd.concat([df_headers, df_final], axis=0)
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, header=False, sheet_name='Reporte')
        
        # Limpieza final de pandas structures
        del df_final
        del df_export
        limpiar_memoria()

        # 5. Maquillaje (OpenPyXL)
        output.seek(0)
        wb = load_workbook(output)
        ws = wb.active
        
        # Estilos (reducidos a lo esencial para velocidad)
        fill_gris = PatternFill(start_color="404040", end_color="404040", fill_type="solid")
        font_blanca = Font(color="FFFFFF", bold=True)
        fill_amarillo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        borde = Side(style="thin", color="000000")
        caja = Border(left=borde, right=borde, top=borde, bottom=borde)
        
        letra_col_act = None
        for col_idx, cell in enumerate(ws[1], 1):
            if "Act" in str(cell.value): letra_col_act = get_column_letter(col_idx)

        # Aplicamos estilos por rango en lugar de celda por celda (Más rápido)
        # Nota: OpenPyXL no soporta estilos por rango nativo fácil sin iterar, 
        # mantenemos iteración pero protegida.
        for row in ws.iter_rows():
            for cell in row:
                cell.border = caja
                if cell.row <= 2:
                    cell.fill = fill_gris; cell.font = font_blanca
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif letra_col_act and cell.column_letter == letra_col_act:
                    cell.fill = fill_amarillo
                    cell.alignment = Alignment(horizontal="left")
        
        # Ajuste de ancho básico
        for col in ws.columns:
            try:
                # Tomamos solo los primeros 100 valores para calcular ancho y no tardar años
                ws.column_dimensions[col[0].column_letter].width = 15 
            except: pass

        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)
        
        return final_output, None

    except Exception as e:
        return None, f"Error de proceso: {str(e)}"

# --- INTERFAZ ---
archivo = st.file_uploader("Sube tu Excel (.xlsx)", type=["xlsx"])

if archivo:
    if st.button("🚀 Procesar"):
        # Leemos los bytes una sola vez
        bytes_data = archivo.getvalue()
        
        with st.spinner("Procesando... (Esto puede tardar unos segundos)"):
            resultado, error = procesar_excel_optimizado(bytes_data)
            
            if error:
                st.error(f"❌ {error}")
                st.warning("Consejo: Si tu archivo es muy pesado, elimina las filas vacías al final del Excel.")
            else:
                st.success("✅ ¡Listo!")
                st.download_button(
                    label="📥 Descargar Reporte",
                    data=resultado,
                    file_name="Reporte_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
