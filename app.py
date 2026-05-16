import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from datetime import datetime

# ------------------------------------------------------------
# 1. CARGA Y PROCESAMIENTO DE DATOS DESDE EL EXCEL
# ------------------------------------------------------------
@st.cache_data
def load_excel_data(file_path, sheet_name="MAYO 2026"):
    """Carga los datos del Excel y extrae los agentes de la zona JC"""
    try:
        # Cargar el workbook y la hoja específica
        wb = load_workbook(file_path, data_only=True)
        sheet = wb[sheet_name]
        
        # Obtener todas las filas como lista de listas
        all_rows = []
        for row in sheet.iter_rows(values_only=True):
            all_rows.append(list(row))
        
        if not all_rows:
            return None, "No se encontraron datos en la hoja"
        
        # Identificar la fila de encabezados (buscamos "CONTRATO" o similar)
        header_row_idx = None
        for i, row in enumerate(all_rows):
            if row and any(cell and "CONTRATO" in str(cell) for cell in row):
                header_row_idx = i
                break
        
        if header_row_idx is None:
            return None, "No se encontró la fila de encabezados"
        
        headers = all_rows[header_row_idx]
        # Encontrar índice de columna de nombres y códigos
        name_col_idx = None
        code_col_idx = None
        est_col_idx = None
        zona_col_idx = None  # Columna A (identificador JC, AEZ6, etc.)
        
        for i, cell in enumerate(headers):
            if cell and "NOMBRE" in str(cell).upper():
                name_col_idx = i
            if cell and "COD." in str(cell).upper() or (cell and i == 2 and "COD." in str(headers[i-1] if i>0 else "")):
                # Buscar columna de código (suele ser columna C)
                if i < len(all_rows[header_row_idx-1]) and all_rows[header_row_idx-1][i] and "COD" in str(all_rows[header_row_idx-1][i]):
                    code_col_idx = i
            if cell and "EST." in str(cell).upper():
                est_col_idx = i
            if i == 0:  # Primera columna suele tener identificador de zona (JC, AV, AA, etc.)
                zona_col_idx = i
        
        # Si no encontramos, usamos valores por defecto basados en la estructura observada
        if name_col_idx is None:
            name_col_idx = 4  # Columna E (índice 4) suele tener los nombres
        if code_col_idx is None:
            code_col_idx = 2  # Columna C (índice 2) suele tener los códigos
        
        # Encontrar la fila donde empiezan los agentes (después del encabezado)
        agents_data = []
        start_row = header_row_idx + 1
        
        # Días del mes (columnas E hasta BN, aproximadamente)
        # Los días están en la fila de encabezados superior
        day_columns = []
        for i in range(start_row, min(start_row + 200, len(all_rows))):
            row = all_rows[i]
            if not row or len(row) < 5:
                continue
            
            # Verificar si es un agente (tiene código y nombre)
            code_cell = row[code_col_idx] if code_col_idx < len(row) else None
            name_cell = row[name_col_idx] if name_col_idx < len(row) else None
            zona_id = row[zona_col_idx] if zona_col_idx < len(row) else None
            
            # Solo procesamos filas que parecen agentes (tienen código numérico y nombre)
            if code_cell and name_cell and str(code_cell).strip() and str(name_cell).strip():
                # Determinar la zona basada en la primera columna o contenido
                zona = "JC"  # Por defecto
                if zona_id and isinstance(zona_id, str):
                    if "AV" in zona_id.upper() or "CID" in zona_id.upper():
                        zona = "AV"
                    elif "AA" in zona_id.upper() or "ALAMEDA" in str(row).upper():
                        zona = "AA"
                    elif "FO" in zona_id.upper() or "FOIOS" in str(row).upper():
                        zona = "FO"
                    elif "MS" in zona_id.upper() or "MASSAMAGRELL" in str(row).upper():
                        zona = "MS"
                    elif "AE" in zona_id.upper() or "AEROPORT" in str(row).upper():
                        zona = "AE"
                    elif "MM" in zona_id.upper() or "MARITIM" in str(row).upper():
                        zona = "MM"
                    elif "TMC" in zona_id.upper() or "TALLER" in str(row).upper():
                        zona = "TMC"
                
                # Extraer turnos (desde columna F hasta antes de CONTRATO)
                shifts = []
                start_shift_col = 5  # Columna F (índice 5) es el primer día
                # Buscar hasta la columna donde está "CONTRATO" o límite
                max_col = min(len(row), 70)  # Hasta columna BN aproximadamente
                
                for col in range(start_shift_col, max_col):
                    shift_val = row[col] if col < len(row) else ""
                    if shift_val is None:
                        shift_val = ""
                    shifts.append(str(shift_val).strip() if shift_val != "" else "")
                
                # Limitar a 31 días (mayo)
                shifts = shifts[:31]
                while len(shifts) < 31:
                    shifts.append("")
                
                agents_data.append({
                    "zona": zona,
                    "codigo": str(code_cell).strip(),
                    "nombre": str(name_cell).strip(),
                    "turnos": shifts
                })
        
        # Filtrar solo zona JC para la vista inicial
        jc_agents = [a for a in agents_data if a["zona"] == "JC"]
        
        if not jc_agents:
            return None, "No se encontraron agentes en la zona JC"
        
        # Crear DataFrame para los turnos
        days = list(range(1, 32))
        df_data = []
        for agent in jc_agents:
            row_data = {"Código": agent["codigo"], "Agente": agent["nombre"]}
            for i, day in enumerate(days):
                row_data[f"Día {day}"] = agent["turnos"][i] if i < len(agent["turnos"]) else ""
            df_data.append(row_data)
        
        df = pd.DataFrame(df_data)
        return df, f"✅ Cargados {len(jc_agents)} agentes de la zona JC"
    
    except Exception as e:
        return None, f"❌ Error al cargar el archivo: {str(e)}"

# ------------------------------------------------------------
# 2. CONFIGURACIÓN DE LA APLICACIÓN STREAMLIT
# ------------------------------------------------------------
st.set_page_config(page_title="Cuadrantes Metrovalencia", layout="wide", page_icon="📅")

st.title("📅 Gestión de Cuadrantes - Metrovalencia")
st.markdown("### Visualización de turnos y administración de reglas")

# Sidebar para carga de archivo
with st.sidebar:
    st.header("📁 Cargar archivo Excel")
    uploaded_file = st.file_uploader("Selecciona el archivo AÑO 2026 ESTACIONES .xlsx", type=["xlsx"])
    
    st.header("🗂️ Seleccionar zona")
    zona_seleccionada = st.selectbox("Zona", ["JC", "AEZ6", "AEZ7", "AEZ8"], index=0)
    
    st.header("⚙️ Modo Administrador")
    admin_mode = st.checkbox("Activar modo administrador", value=False)
    
    if admin_mode:
        st.subheader("📋 Reglas e incompatibilidades")
        st.info("Define reglas para detectar conflictos en los turnos")
        nueva_regla = st.text_input("Nueva regla (ej: D + N = conflicto)")
        if st.button("➕ Añadir regla") and nueva_regla:
            if "reglas" not in st.session_state:
                st.session_state.reglas = []
            st.session_state.reglas.append(nueva_regla)
            st.success(f"Regla añadida: {nueva_regla}")
        
        if "reglas" in st.session_state and st.session_state.reglas:
            st.write("**Reglas activas:**")
            for i, regla in enumerate(st.session_state.reglas):
                col1, col2 = st.columns([4, 1])
                col1.write(f"- {regla}")
                if col2.button("❌", key=f"del_{i}"):
                    st.session_state.reglas.pop(i)
                    st.rerun()

# Estado inicial para reglas
if "reglas" not in st.session_state:
    st.session_state.reglas = []

# Cargar datos
if uploaded_file is not None:
    # Guardar archivo temporalmente
    import tempfile
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_path = tmp_file.name
    
    df, msg = load_excel_data(tmp_path, "MAYO 2026")
    
    if df is not None:
        st.success(msg)
        
        # Mostrar selector de agente para filtro
        col1, col2 = st.columns([1, 3])
        with col1:
            agentes = df["Agente"].tolist()
            agente_filtro = st.selectbox("🔍 Filtrar por agente", ["Todos"] + agentes)
        
        # Filtrar datos
        if agente_filtro != "Todos":
            df_filtrado = df[df["Agente"] == agente_filtro]
        else:
            df_filtrado = df
        
        # Mostrar estadísticas
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total agentes", len(df))
        with col2:
            # Contar días con turnos
            turnos_unicos = set()
            for col in df.columns[2:]:  # Saltar Código y Agente
                turnos_unicos.update(df[col].unique())
            st.metric("Tipos de turno", len([t for t in turnos_unicos if t and t != ""]))
        with col3:
            st.metric("Mes", "Mayo 2026")
        
        # Mostrar tabla de turnos
        st.markdown("### 📊 Cuadrante de turnos")
        
        # Configurar estilo de la tabla
        st.dataframe(
            df_filtrado,
            use_container_width=True,
            height=min(800, len(df_filtrado) * 35 + 38),
            column_config={
                "Código": st.column_config.TextColumn("Código", width="small"),
                "Agente": st.column_config.TextColumn("Agente", width="medium"),
                **{f"Día {i}": st.column_config.TextColumn(f"Día {i}", width="small") for i in range(1, 32)}
            }
        )
        
        # Vista detallada de un agente específico
        st.markdown("---")
        st.markdown("### 👤 Detalle de agente")
        
        agente_detalle = st.selectbox("Seleccionar agente para ver detalles", agentes)
        if agente_detalle:
            agente_data = df[df["Agente"] == agente_detalle].iloc[0]
            st.subheader(f"📋 {agente_detalle} (Código: {agente_data['Código']})")
            
            # Mostrar turnos del mes en una cuadrícula
            cols = st.columns(7)
            for i, day in enumerate(range(1, 32)):
                col_idx = (i % 7)
                turno = agente_data[f"Día {day}"]
                with cols[col_idx]:
                    st.metric(f"Día {day}", turno if turno else "—")
            
            # Mostrar resumen de turnos
            turnos_lista = [agente_data[f"Día {d}"] for d in range(1, 32)]
            turnos_frecuencia = {}
            for t in turnos_lista:
                if t and t != "":
                    turnos_frecuencia[t] = turnos_frecuencia.get(t, 0) + 1
            
            if turnos_frecuencia:
                st.markdown("**Resumen de turnos en el mes:**")
                for turno, count in sorted(turnos_frecuencia.items(), key=lambda x: x[1], reverse=True):
                    st.write(f"- {turno}: {count} día(s)")
        
        # Panel de reglas en modo administrador
        if admin_mode and st.session_state.reglas:
            st.markdown("---")
            st.markdown("### ⚠️ Verificación de reglas")
            st.info("Las siguientes reglas están activas y se aplicarán en futuras versiones para validar los turnos.")
            for regla in st.session_state.reglas:
                st.write(f"- {regla}")
        
    else:
        st.error(msg)
else:
    st.info("👈 Por favor, carga el archivo Excel 'AÑO 2026 ESTACIONES .xlsx' en el panel lateral para comenzar")
    
    # Mostrar ejemplo de estructura esperada
    with st.expander("📖 Instrucciones"):
        st.markdown("""
        ### Formato esperado del archivo Excel:
        
        1. El archivo debe contener una hoja llamada **"MAYO 2026"**
        2. La estructura debe incluir:
           - Columna con identificador de zona (JC, AV, AA, etc.)
           - Columna con código del agente
           - Columna con nombre del agente
           - Columnas de días (1 al 31) con los turnos asignados
        
        ### Zonas disponibles:
        - **JC**: Jefes de Circulación (principal)
        - **AEZ6, AEZ7, AEZ8**: Agentes de estaciones (en desarrollo)
        
        ### Funcionalidades:
        - Visualización de cuadrante completo
        - Filtro por agente
        - Modo administrador para añadir reglas de negocio
        - Vista detallada por agente
        """)

# Footer
st.markdown("---")
st.markdown("📅 **Metrovalencia - Gestión de Cuadrantes** | Desarrollado para visualización de turnos y administración de personal")