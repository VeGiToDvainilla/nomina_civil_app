import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
import io

# --- TÍTULO Y DESCRIPCIÓN ---
st.set_page_config(page_title="Separador de Actividades", page_icon="🏗️")
st.title("🏗️ Herramienta: Separador de actividades por fila")
st.write("""
**Instrucciones:**
1. Sube tu archivo de Excel (`.xlsx`).
2. El sistema desglosará las filas y corregirá el formato automáticamente.
3. Descarga el reporte listo.
""")

# --- LA FUNCIÓN DE PROCESAMIENTO ---
def procesar_excel(uploaded_file):
    try:
        # Leemos el archivo
        df_raw = pd.read_excel(uploaded_file, header=None, nrows=20)
        
        fila_encabezado = -1
        for i, fila in df_raw.iterrows():
            texto = fila.astype(str).str.upper()
            if texto.str.contains("CLAVE").any() and texto.str.contains("ASIST").any():
                fila_encabezado = i
                break
                
        if fila_encabezado == -1:
            return None, "No se encontraron los encabezados 'Clave' y 'Asist'."

        # Recuperar encabezados con método compatible
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

        # Cargar datos
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, header=None, skiprows=fila_encabezado + 2)
        df = df.iloc[:, :len(nombres_columnas)]
        df.columns = nombres_columnas
        
        for col in indices_rmmal: 
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        if columna_comida: 
            df[columna_comida] = pd.to_numeric(df[columna_comida], errors='coerce').fillna(0)

        # Desglose
        nuevas_filas = []
        for idx, row in df.iterrows():
            cols_activas = [c for c in indices_rmmal if row[c] > 0]
            if not cols_activas: continue
            
            horas_map = {c: row[c] for c in cols_activas}
            columna_ganadora = max(horas_map, key=horas_map.get)
            
            for col_actual in cols_activas:
                fila_nueva = row.copy()
                partes = col_actual.split('|')
                
                # Buscar columnas destino con seguridad
                try:
                    idx_act = [c for c in df.columns if c.startswith("Act|")][0]
                    idx_turno = [c for c in df.columns if c.startswith("Turno|")][0]
                except IndexError:
                    return None, "No se encontraron las columnas 'Act' o 'Turno' en el archivo."

                fila_nueva[idx_act] = partes[0]
                fila_nueva[idx_turno] = partes[1]
                
                for c in indices_rmmal:
                    if c != col_actual: fila_nueva[c] = 0
                
                if columna_comida:
                    if col_actual != columna_ganadora: fila_nueva[columna_comida] = 0
                
                nuevas_filas.append(fila_nueva)
                
        df_final = pd.DataFrame(nuevas_filas, columns=nombres_columnas)
        
        # Limpieza de Fechas
        cols_fecha = [c for c in df_final.columns if "FECHA" in c.upper()]
        for col in cols_fecha:
            df_final[col] = pd.to_datetime(df_final[col], errors='coerce').dt.date

        # Exportar a memoria
        output = io.BytesIO()
        df_headers = pd.DataFrame([header_top.values, header_bottom.values], columns=nombres_columnas)
        df_export = pd.concat([df_headers, df_final], axis=0)
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, header=False, sheet_name='Reporte')
        
        # Maquillaje
        output.seek(0)
        wb = load_workbook(output)
        ws = wb.active
        
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
                    celda.alignment = Alignment(horizontal="center", vertical="center")
                elif letra_col_act and celda.column_letter == letra_col_act:
                    celda.fill = fill_amarillo
                    celda.alignment = Alignment(horizontal="left")
                if celda.value and str(celda.value).startswith("202") and "-" in str(celda.value):
                     celda.alignment = Alignment(horizontal="center")

        for col in ws.columns:
            try:
                val_list = [len(str(cell.value) or "") for cell in col]
                max_len = max(val_list) if val_list else 0
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 50)
            except: pass
            
        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)
        
        return final_output, None
        
    except Exception as e:
        return None, f"Error inesperado: {str(e)}"

# --- INTERFAZ DE USUARIO ---
archivo = st.file_uploader("Arrastra tu archivo aquí", type=["xlsx"])

if archivo:
    # Botón para iniciar el proceso
    if st.button("🚀 Procesar Archivo"):
        with st.spinner("Procesando... por favor espera"):
            resultado, error = procesar_excel(archivo)
            
            if error:
                st.error(f"❌ {error}")
            else:
                st.success("✅ ¡Archivo procesado con éxito!")
                # Guardamos el resultado en el estado de la sesión para que no se borre
                st.session_state['resultado_listo'] = resultado

# Mostrar el botón de descarga si el resultado ya existe en memoria
if 'resultado_listo' in st.session_state:
    st.download_button(
        label="📥 Descargar Reporte Final",
        data=st.session_state['resultado_listo'],
        file_name="Reporte_Desglosado_Listo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
