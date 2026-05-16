import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile

st.set_page_config(page_title="Cuadrantes Metrovalencia", layout="wide")

st.title("📅 Cuadrante de Servicios - Metrovalencia")
st.caption("Solo se lee la primera columna de cada día (turno) | Edición con guardado directo")

# ============================================================
# FUNCIONES PRINCIPALES
# ============================================================

def cargar_excel(archivo_path):
    """Carga el Excel extrayendo SOLO la primera columna de cada día (turno)"""
    try:
        wb = load_workbook(archivo_path, data_only=True)
        sheet = wb["MAYO 2026"]
        
        # Buscar fila de encabezados (columna E suele tener "NOMBRE")
        header_row = None
        for i in range(1, 100):
            celda = sheet.cell(row=i, column=5).value
            if celda and "NOMBRE" in str(celda).upper():
                header_row = i
                break
        
        if not header_row:
            return None, None, "No se encontró la fila de encabezados"
        
        agentes = []
        
        # Las columnas de turnos (solo las impares: F, H, J, L, N, P...)
        # Columna F = día 1, H = día 2, J = día 3, etc.
        columnas_turnos = [6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38, 40, 42, 44, 46, 48, 50, 52, 54, 56, 58, 60, 62, 64, 66]
        
        for fila in range(header_row + 1, sheet.max_row + 1):
            # Columna A = zona, Columna C = código, Columna E = nombre
            zona = sheet.cell(row=fila, column=1).value
            codigo = sheet.cell(row=fila, column=3).value
            nombre = sheet.cell(row=fila, column=5).value
            
            # Filtrar agentes válidos
            if not codigo or not nombre:
                continue
            if str(codigo).strip() == "0" or str(codigo).strip() == "":
                continue
            if "DESPLAZADO" in str(nombre).upper() or "VACANTE" in str(nombre).upper():
                continue
            
            # Extraer SOLO la primera columna de cada día (turno)
            turnos = []
            celdas_turnos = []
            
            for col in columnas_turnos:
                celda = sheet.cell(row=fila, column=col)
                valor = celda.value
                turnos.append(str(valor).strip() if valor else "")
                celdas_turnos.append(celda)
            
            agentes.append({
                "zona": str(zona).strip() if zona else "",
                "codigo": str(codigo).strip(),
                "nombre": str(nombre).strip(),
                "turnos": turnos,
                "celdas_turnos": celdas_turnos,
                "fila": fila
            })
        
        if not agentes:
            return None, None, "No se encontraron agentes válidos"
        
        # Organizar por zona
        agentes_por_zona = {}
        for ag in agentes:
            zona = ag["zona"]
            if zona not in agentes_por_zona:
                agentes_por_zona[zona] = []
            agentes_por_zona[zona].append(ag)
        
        return agentes_por_zona, wb, f"✅ Cargados {len(agentes)} agentes"
    
    except Exception as e:
        return None, None, f"❌ Error: {e}"


def intercambiar_turnos(agentes, idx1, idx2, dia):
    """Intercambia SOLO el turno (primera columna) entre dos agentes"""
    if idx1 == idx2:
        return False
    
    ag1 = agentes[idx1]
    ag2 = agentes[idx2]
    
    turno1 = ag1["turnos"][dia]
    turno2 = ag2["turnos"][dia]
    
    # Intercambiar en memoria
    ag1["turnos"][dia] = turno2
    ag2["turnos"][dia] = turno1
    
    # Intercambiar en las celdas de Excel
    celda1 = ag1["celdas_turnos"][dia]
    celda2 = ag2["celdas_turnos"][dia]
    celda1.value = turno2 if turno2 else None
    celda2.value = turno1 if turno1 else None
    
    return True


def guardar_excel(wb, archivo_path):
    """Guarda los cambios en el archivo"""
    try:
        wb.save(archivo_path)
        return True, "✅ Cambios guardados"
    except Exception as e:
        return False, f"❌ Error al guardar: {e}"


# ============================================================
# INTERFAZ
# ============================================================

with st.sidebar:
    st.header("📁 Cargar archivo")
    archivo_subido = st.file_uploader("Selecciona el Excel", type=["xlsx"])
    
    if archivo_subido:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(archivo_subido.getvalue())
            archivo_temp = tmp.name
        
        agentes_por_zona, wb, msg = cargar_excel(archivo_temp)
        
        if agentes_por_zona:
            st.success(msg)
            st.session_state['agentes_por_zona'] = agentes_por_zona
            st.session_state['wb'] = wb
            st.session_state['archivo_path'] = archivo_temp
            st.session_state['cargado'] = True
        else:
            st.error(msg)
    
    if st.session_state.get('cargado', False):
        st.markdown("---")
        zonas = list(st.session_state['agentes_por_zona'].keys())
        zona_seleccionada = st.selectbox("📍 Zona", zonas)
        st.session_state['zona'] = zona_seleccionada
        
        agentes = st.session_state['agentes_por_zona'][zona_seleccionada]
        st.metric("👥 Agentes", len(agentes))

# Contenido principal
if not st.session_state.get('cargado', False):
    st.info("👈 Carga el archivo Excel en el panel lateral")
else:
    zona = st.session_state['zona']
    agentes = st.session_state['agentes_por_zona'][zona]
    
    # Construir tabla
    dias = [f"D{i}" for i in range(1, 32)]
    
    data = []
    for i, ag in enumerate(agentes):
        fila = {"idx": i, "Código": ag["codigo"], "Agente": ag["nombre"]}
        for d in range(31):
            turno = ag["turnos"][d]
            fila[dias[d]] = turno if turno else "—"
        data.append(fila)
    
    df = pd.DataFrame(data)
    
    st.markdown(f"## 📊 Zona {zona}")
    
    column_config = {
        "idx": st.column_config.NumberColumn("ID", width="small"),
        "Código": st.column_config.TextColumn("Código", width="small"),
        "Agente": st.column_config.TextColumn("Agente", width="medium"),
    }
    for d in dias:
        column_config[d] = st.column_config.TextColumn(d, width="small")
    
    st.dataframe(
        df,
        use_container_width=True,
        height=min(600, len(agentes) * 35 + 38),
        column_config=column_config,
        hide_index=True
    )
    
    # Interfaz de edición
    st.markdown("---")
    st.markdown("## 🔄 Intercambiar turnos")
    st.caption("Solo se intercambia el turno (primera columna). La segunda columna (datos adicionales) no se modifica.")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        nombres = [f"{a['nombre']} (Cód: {a['codigo']})" for a in agentes]
        ag1 = st.selectbox("Agente 1", nombres, key="ag1")
        idx1 = nombres.index(ag1)
    
    with col2:
        ag2 = st.selectbox("Agente 2", nombres, key="ag2")
        idx2 = nombres.index(ag2)
    
    with col3:
        dia = st.selectbox("Día", list(range(1, 32)), key="dia")
        dia_idx = dia - 1
    
    # Mostrar turnos actuales
    turno1 = agentes[idx1]["turnos"][dia_idx]
    turno2 = agentes[idx2]["turnos"][dia_idx]
    st.info(f"**Turno actual:** {agentes[idx1]['nombre']} → {turno1 if turno1 else '—'} | {agentes[idx2]['nombre']} → {turno2 if turno2 else '—'}")
    
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        if st.button("🔄 Intercambiar (solo vista)", use_container_width=True):
            if intercambiar_turnos(agentes, idx1, idx2, dia_idx):
                st.success("✅ Turnos intercambiados en memoria")
                st.rerun()
            else:
                st.error("Error")
    
    with col_btn2:
        if st.button("💾 Intercambiar Y GUARDAR", type="primary", use_container_width=True):
            if intercambiar_turnos(agentes, idx1, idx2, dia_idx):
                ok, msg = guardar_excel(st.session_state['wb'], st.session_state['archivo_path'])
                if ok:
                    st.success(msg)
                    st.balloons()
                    st.rerun()
                else:
                    st.error(msg)
            else:
                st.error("Error al intercambiar")
