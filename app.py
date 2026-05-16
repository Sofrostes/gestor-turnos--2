import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile

st.set_page_config(page_title="Cuadrantes Metrovalencia", layout="wide")

st.title("📅 Cuadrante de Servicios - Metrovalencia")
st.caption("Usando hoja REL para agentes | Hoja MAYO 2026 para turnos")

# ============================================================
# FUNCIONES
# ============================================================

def cargar_agentes_desde_rel(archivo_path):
    """Carga la lista de agentes desde la hoja REL"""
    try:
        wb = load_workbook(archivo_path, data_only=True)
        
        if "REL" not in wb.sheetnames:
            return None, None, "No se encontró la hoja REL"
        
        sheet = wb["REL"]
        
        agentes = []
        
        # Buscar encabezados en REL
        header_row = None
        for i in range(1, 20):
            fila = sheet.cell(row=i, column=1).value
            if fila and ("COD" in str(fila).upper() or "CÓD" in str(fila).upper()):
                header_row = i
                break
        
        if not header_row:
            header_row = 1
        
        # Encontrar columnas
        col_codigo = None
        col_nombre = None
        col_zona = None
        
        for col in range(1, 10):
            celda = sheet.cell(row=header_row, column=col).value
            if celda:
                celda_str = str(celda).upper()
                if "COD" in celda_str or "CÓD" in celda_str:
                    col_codigo = col
                elif "NOMBRE" in celda_str or "AGENTE" in celda_str:
                    col_nombre = col
                elif "ZONA" in celda_str:
                    col_zona = col
        
        # Si no encuentra, usar valores por defecto
        if col_codigo is None:
            col_codigo = 1
        if col_nombre is None:
            col_nombre = 2
        if col_zona is None:
            col_zona = 3
        
        # Leer agentes
        for fila in range(header_row + 1, sheet.max_row + 1):
            codigo = sheet.cell(row=fila, column=col_codigo).value
            nombre = sheet.cell(row=fila, column=col_nombre).value
            zona = sheet.cell(row=fila, column=col_zona).value if col_zona else ""
            
            if not codigo or not nombre:
                continue
            if str(codigo).strip() == "0":
                continue
            
            agentes.append({
                "codigo": str(codigo).strip(),
                "nombre": str(nombre).strip(),
                "zona": str(zona).strip() if zona else "",
                "turnos": []  # Se llenará después
            })
        
        return agentes, wb, f"✅ Cargados {len(agentes)} agentes desde REL"
    
    except Exception as e:
        return None, None, f"❌ Error en REL: {e}"


def cargar_turnos_desde_mayo(agentes, archivo_path):
    """Carga los turnos desde la hoja MAYO 2026"""
    try:
        wb = load_workbook(archivo_path, data_only=True)
        
        if "MAYO 2026" not in wb.sheetnames:
            return None, "No se encontró la hoja MAYO 2026"
        
        sheet = wb["MAYO 2026"]
        
        # Buscar fila de encabezados (donde están los días)
        header_row = None
        for i in range(1, 20):
            fila = sheet.cell(row=i, column=6).value  # Columna F
            if fila and str(fila).strip().isdigit() and 1 <= int(fila) <= 31:
                header_row = i
                break
        
        if not header_row:
            header_row = 11  # Valor por defecto observado
        
        # Crear diccionario de agentes por código
        agentes_dict = {ag["codigo"]: ag for ag in agentes}
        
        # Columnas de turnos (F=6, H=8, J=10...)
        columnas_turnos = [6 + i*2 for i in range(31)]
        
        # Buscar agentes en MAYO
        for fila in range(header_row + 1, sheet.max_row + 1):
            codigo = sheet.cell(row=fila, column=3).value  # Columna C
            nombre = sheet.cell(row=fila, column=5).value  # Columna E
            
            if not codigo:
                continue
            
            codigo_str = str(codigo).strip()
            
            if codigo_str in agentes_dict:
                # Leer turnos
                turnos = []
                for col in columnas_turnos:
                    valor = sheet.cell(row=fila, column=col).value
                    turnos.append(str(valor).strip() if valor else "")
                
                agentes_dict[codigo_str]["turnos"] = turnos
                agentes_dict[codigo_str]["fila_mayo"] = fila
                agentes_dict[codigo_str]["celdas_mayo"] = [(fila, col) for col in columnas_turnos]
        
        # Filtrar agentes que tienen turnos
        agentes_con_turnos = [ag for ag in agentes if ag["turnos"]]
        
        return agentes_con_turnos, f"✅ Cargados turnos para {len(agentes_con_turnos)} agentes"
    
    except Exception as e:
        return None, f"❌ Error en MAYO: {e}"


def guardar_cambios(agentes, archivo_path):
    """Guarda los turnos modificados en el Excel"""
    try:
        wb = load_workbook(archivo_path)
        sheet = wb["MAYO 2026"]
        
        columnas_turnos = [6 + i*2 for i in range(31)]
        
        for agente in agentes:
            if "fila_mayo" in agente and "celdas_mayo" in agente:
                for i, turno in enumerate(agente["turnos"]):
                    fila, col = agente["celdas_mayo"][i]
                    sheet.cell(row=fila, column=col).value = turno if turno else None
        
        wb.save(archivo_path)
        return True, "✅ Cambios guardados"
    except Exception as e:
        return False, f"❌ Error al guardar: {e}"


def intercambiar_turnos(agentes, idx1, idx2, dia):
    """Intercambia turnos entre dos agentes"""
    if idx1 == idx2:
        return False
    
    turno1 = agentes[idx1]["turnos"][dia]
    turno2 = agentes[idx2]["turnos"][dia]
    
    agentes[idx1]["turnos"][dia] = turno2
    agentes[idx2]["turnos"][dia] = turno1
    
    return True


# ============================================================
# INICIALIZACIÓN
# ============================================================

if 'agentes' not in st.session_state:
    st.session_state.agentes = None
if 'archivo_path' not in st.session_state:
    st.session_state.archivo_path = None
if 'cargado' not in st.session_state:
    st.session_state.cargado = False


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("📁 Cargar archivo")
    archivo_subido = st.file_uploader("Selecciona el Excel", type=["xlsx"])
    
    if archivo_subido and not st.session_state.cargado:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(archivo_subido.getvalue())
            archivo_temp = tmp.name
        
        st.session_state.archivo_path = archivo_temp
        
        with st.spinner("Cargando agentes desde REL..."):
            agentes, wb, msg = cargar_agentes_desde_rel(archivo_temp)
        
        if agentes:
            st.success(msg)
            
            with st.spinner("Cargando turnos desde MAYO 2026..."):
                agentes_con_turnos, msg2 = cargar_turnos_desde_mayo(agentes, archivo_temp)
            
            if agentes_con_turnos:
                st.session_state.agentes = agentes_con_turnos
                st.session_state.cargado = True
                st.success(msg2)
                
                # Mostrar distribución por zona
                zonas = {}
                for ag in agentes_con_turnos:
                    zona = ag["zona"] if ag["zona"] else "SIN ZONA"
                    zonas[zona] = zonas.get(zona, 0) + 1
                
                st.write("**Distribución por zona:**")
                for zona, count in zonas.items():
                    st.write(f"- {zona}: {count} agentes")
            else:
                st.error(msg2)
        else:
            st.error(msg)
    
    if st.session_state.cargado:
        st.markdown("---")
        
        # Selector de zona
        zonas = list(set(ag["zona"] for ag in st.session_state.agentes))
        zona_seleccionada = st.selectbox("📍 Zona", zonas)
        
        agentes_filtrados = [ag for ag in st.session_state.agentes if ag["zona"] == zona_seleccionada]
        st.session_state.agentes_filtrados = agentes_filtrados
        
        st.metric("👥 Agentes", len(agentes_filtrados))
        
        st.markdown("---")
        
        if st.button("💾 Guardar cambios en Excel", use_container_width=True):
            ok, msg = guardar_cambios(st.session_state.agentes, st.session_state.archivo_path)
            if ok:
                st.success(msg)
            else:
                st.error(msg)


# ============================================================
# CONTENIDO PRINCIPAL
# ============================================================

if not st.session_state.cargado:
    st.info("👈 Carga el archivo Excel. La aplicación usará la hoja REL para los agentes y MAYO 2026 para los turnos.")
    
    with st.expander("📖 Estructura esperada"):
        st.markdown("""
        ### Hojas utilizadas:
        
        1. **REL** (Relación de agentes):
           - Columna 1: Código del agente
           - Columna 2: Nombre del agente  
           - Columna 3: Zona (JC, AV, AA, etc.)
        
        2. **MAYO 2026** (Turnos):
           - Columna C: Código del agente
           - Columna E: Nombre del agente
           - Columna F, H, J...: Turnos (día 1, 2, 3...)
        """)
else:
    agentes = st.session_state.agentes_filtrados
    zona = st.session_state.get('zona_seleccionada', '')
    
    # Construir DataFrame
    dias = list(range(1, 32))
    data = []
    
    for i, ag in enumerate(agentes):
        fila = {"#": i, "Código": ag["codigo"], "Agente": ag["nombre"]}
        for d in range(31):
            turno = ag["turnos"][d] if d < len(ag["turnos"]) else ""
            fila[f"D{d+1}"] = turno if turno else "—"
        data.append(fila)
    
    df = pd.DataFrame(data)
    
    st.markdown(f"## 📊 Zona {zona}")
    st.caption(f"Total agentes: {len(agentes)}")
    
    column_config = {
        "#": st.column_config.NumberColumn("#", width="small"),
        "Código": st.column_config.TextColumn("Código", width="small"),
        "Agente": st.column_config.TextColumn("Agente", width="medium"),
    }
    for d in range(1, 32):
        column_config[f"D{d}"] = st.column_config.TextColumn(f"D{d}", width="small")
    
    st.dataframe(
        df,
        use_container_width=True,
        height=min(700, len(agentes) * 35 + 50),
        column_config=column_config,
        hide_index=True
    )
    
    # Interfaz de intercambio
    st.markdown("---")
    st.markdown("## 🔄 Intercambiar turnos")
    
    if len(agentes) >= 2:
        col1, col2, col3 = st.columns(3)
        
        with col1:
            nombres = [f"{a['nombre']}" for a in agentes]
            ag1 = st.selectbox("Agente 1", nombres, key="ag1")
            idx1 = nombres.index(ag1)
        
        with col2:
            ag2 = st.selectbox("Agente 2", nombres, key="ag2")
            idx2 = nombres.index(ag2)
        
        with col3:
            dia = st.selectbox("Día", list(range(1, 32)), key="dia")
            dia_idx = dia - 1
        
        # Mostrar turnos actuales
        turno1 = agentes[idx1]["turnos"][dia_idx] if dia_idx < len(agentes[idx1]["turnos"]) else ""
        turno2 = agentes[idx2]["turnos"][dia_idx] if dia_idx < len(agentes[idx2]["turnos"]) else ""
        
        st.info(f"📌 **Actual:** {agentes[idx1]['nombre']} → `{turno1 if turno1 else '—'}` | {agentes[idx2]['nombre']} → `{turno2 if turno2 else '—'}`")
        
        if st.button("🔄 Intercambiar turnos", type="primary", use_container_width=True):
            if intercambiar_turnos(agentes, idx1, idx2, dia_idx):
                st.success("✅ Turnos intercambiados (recuerda guardar en el panel lateral)")
                st.rerun()
            else:
                st.error("Error al intercambiar")
    else:
        st.warning("Se necesitan al menos 2 agentes en esta zona")
