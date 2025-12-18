import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Separador Debug", page_icon="🐞")
st.title("🐞 Herramienta de Nómina (Modo Diagnóstico)")
st.warning("Si ves un error rojo, por favor toma una captura de pantalla y pásala al chat.")

# --- FUNCIÓN ROBUSTA ---
def procesar_excel_seguro(uploaded_file):
    log_pasos = [] # Para guardar el historial de lo que pasa
    
    try:
        log_pasos.append("1. Iniciando lectura del archivo...")
        # Leemos el archivo
        df_raw = pd.read_excel(uploaded_file, header=None, nrows=30)
        log_pasos.append(f"   - Archivo leído. Filas detectadas preliminarmente: {len(df_raw)}")
        
        # BUSCAR ENCABEZADOS
        fila_encabezado = -1
        for i, fila in df_raw.iterrows():
            texto = fila.astype(str).str.upper()
            if texto.str.contains("CLAVE").any() and texto.str.contains("ASIST").any():
                fila_encabezado = i
                break
                
        if fila_encabezado == -1:
            return None, "No encontré las palabras 'CLAVE' y 'ASIST' en las primeras 30 filas.", log_pasos

        log_pasos.append(f"2. Encabezados encontrados en la fila {fila_encabezado + 1}")

        # RECUPERAR NOMBRES DE COLUMNAS
        # Usamos ffill() para rellenar celdas combinadas (ej: RMMAL-05.01 que abarca 2 columnas)
        header_top = df_raw.iloc[fila_encabezado].ffill().astype(str).str.strip()
        header_bottom = df_raw.iloc[fila_encabezado + 1].fillna("").astype(str).str.strip()
        
        nombres_columnas = []
        indices_rmmal = []
        columna_comida = None
        
        for k in range(len(header_top)):
            top = header_top.iloc[k]
            bottom = header_bottom.iloc[k]
            # Limpieza de valores nulos o basura
            if top == "nan": top = f"Col_{k}"
            if bottom == "nan" or bottom == "x": bottom = ""
            
            nombre_unico = f"{top}|{bottom}"
            nombres_columnas.append(nombre_unico)
            
            if "RMMAL" in top.upper(): indices_rmmal.append(nombre_unico)
            if "COMIDA" in top.upper() and "TOTAL" not in top.upper(): columna_comida = nombre_unico

        log_pasos.append(f"3. Columnas detectadas: {len(nombres_columnas)}. Columnas RMMAL: {len(indices_rmmal)}")

        # CARGAR DATOS REALES
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, header=None, skiprows=fila_encabezado + 2)
        
        # AJUSTE DE COLUMNAS (Evitar error "Length mismatch")
        # Si hay más datos que nombres, cortamos los datos.
        # Si hay más nombres que datos, agregamos columnas vacías.
        if df.shape[1] > len(nombres_columnas):
            df = df.iloc[:, :len(nombres_columnas)]
        elif df.shape[1] < len(nombres_columnas):
            for _ in range(len(nombres_columnas) - df.shape[1]):
                df[f"Extra_{_}"] = 0 # Rellenar hueco
            
        df.columns = nombres_columnas
        log_pasos.append(f"4. Datos cargados correctamente. Filas a procesar: {len(df)}")
        
        # CONVERTIR A NÚMEROS
        for col in indices_rmmal: 
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        if columna_comida: 
            df[columna_comida] = pd.to_numeric(df[columna_comida], errors='coerce').fillna(0)

        # DESGLOSE (EL CORAZÓN DEL CÓDIGO)
        nuevas_filas = []
        for idx, row in df.iterrows():
            cols_activas = [c for c in indices_rmmal if row[c] > 0]
            if not cols_activas: continue
            
            horas_map = {c: row[c] for c in cols_activas}
            columna_ganadora = max(horas_map, key=horas_map.get)
            
            for col_actual in cols_activas:
                fila_nueva = row.copy()
                partes = col_actual.split('|')
                
                # Búsqueda segura de columnas Act y Turno
                idx_act_list = [c for c in df.columns if str(c).startswith("Act|")]
                idx_turno_list = [c for c in df.columns if str(c).startswith("Turno|")]
                
                if not idx_act_list or not idx_turno_list:
                    return None, "Error Crítico: No encontré las columnas 'Act' o 'Turno' en el Excel.", log_pasos
                
                fila_nueva[idx_act_list[0]] = partes[0]
                fila_nueva[idx_turno_list[0]] = partes[1]
                
                for c in indices_rmmal:
                    if c != col_actual: fila_nueva[c] = 0
                
                if columna_comida:
                    if col_actual != columna_ganadora: fila_nueva[columna_comida] = 0
                
                nuevas_filas.append(fila_nueva)
                
        df_final = pd.DataFrame(nuevas_filas, columns=nombres_columnas)
        log_pasos.append(f"5. Desglose terminado. Filas resultantes: {len(df_final)}")
        
        # LIMPIEZA DE FECHAS
        cols_fecha = [c for c in df_final.columns if "FECHA" in str(c).upper()]
        for col in cols_fecha:
            df_final[col] = pd.to_datetime(df_final[col], errors='coerce').dt.date

        # EXPORTACIÓN
        output = io.BytesIO()
        df_headers = pd.DataFrame([header_top.values, header_bottom.values], columns=nombres_columnas)
        df_export = pd.concat([df_headers, df_final], axis=0)
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, header=False, sheet_name='Reporte')
        
        # MAQUILLAJE
        output.seek(0)
        wb = load_workbook(output)
        ws = wb.active
        
        # Estilos simples para evitar errores de importación
        fill_gris = PatternFill(start_color="404040", end_color="404040", fill_type="solid")
        font_blanca = Font(color="FFFFFF", bold=True)
        fill_amarillo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        borde = Side(style="thin", color="000000")
        caja = Border(left=borde, right=borde, top=borde, bottom=borde)
        
        letra_col_act = None
        for col_idx, cell in enumerate(ws[1], 1):
            if "Act" in str(cell.value): letra_col_act = get_column_letter(col_idx)

        for fila in ws.iter_rows():
            for celda in fila:
                celda.border = caja
                if celda.row <= 2:
                    celda.fill = fill_gris; celda.font = font_blanca
                elif letra_col_act and celda.column_letter == letra_col_act:
                    celda.fill = fill_amarillo
        
        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)
        
        log_pasos.append("6. ¡Éxito! Archivo generado.")
        return final_output, None, log_pasos
        
    except Exception as e:
        # Aquí capturamos el error exacto
        import traceback
        error_detallado = traceback.format_exc()
        return None, f"Error Inesperado: {str(e)}", log_pasos

# --- INTERFAZ ---
archivo = st.file_uploader("Sube tu Excel (.xlsx)", type=["xlsx"])

if archivo:
    if st.button("🚀 Procesar (Intento Seguro)"):
        resultado, error, historial = procesar_excel_seguro(archivo)
        
        # Mostrar qué pasó (Diagnóstico)
        with st.expander("Ver detalles del proceso (Debug)"):
            for paso in historial:
                st.write(paso)
        
        if error:
            st.error("❌ OCURRIÓ UN ERROR:")
            st.code(error) # Muestra el error en una cajita de código
        else:
            st.success("✅ ¡Funcionó!")
            st.session_state['resultado_final'] = resultado

if 'resultado_final' in st.session_state:
    st.download_button(
        label="📥 Descargar Reporte",
        data=st.session_state['resultado_final'],
        file_name="Reporte_Debug_Listo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
