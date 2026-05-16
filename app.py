import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile

st.set_page_config(page_title="Cuadrantes Metrovalencia", layout="wide")

st.title("📅 Cuadrante de Servicios - Metrovalencia")
st.caption("Datos en memoria | Rápido | Guardado manual en Excel")

# ============================================================
# FUNCIONES
# ============================================================

def cargar_excel(archivo_path):
    """Carga el Excel una sola vez y guarda los datos en memoria"""
    try:
        wb = load_workbook(archivo_path, data_only=True)
        sheet = wb["MAYO 2026"]
        
        # Buscar fila de encabezados
        header_row = None
        for i in range(1, 30):
            celda = sheet.cell(row=i, column=5).value
            if celda and ("NOMBRE" in str(celda).upper() or "AGENTE" in str(celda).upper()):
                header_row = i
                break
        
        if not header_row:
            header_row = 11  # Valor por defecto observado en el archivo
        
        # Columnas de días (F=6, H=8, J=10... hasta día 31)
        columnas_turnos = [6 + i*2 for i in range(31)]  # 6,8,10,12...66
        
        agentes = []
        
        for fila in range(header_row + 1, sheet.max_row + 1):
            zona = sheet.cell(row=fila, column=1).value
            codigo = sheet.cell(row=fila, column=3).value
            nombre = sheet.cell(row=fila, column=5).value
            
            # Filtrar
            if not codigo or not nombre:
                continue
            if str(codigo).strip() == "0":
                continue
            if "DESPLAZADO" in str(nombre).upper() or "VACANTE" in str(nombre).upper():
                continue
            
            # Leer turnos
            turnos = []
            for col in columnas_turnos:
                valor = sheet.cell(row=fila, column=col).value
                turnos.append(str(valor).strip() if valor else "")
            
            agentes.append({
                "zona": str(zona).strip() if zona else "",
                "codigo": str(codigo).strip(),
                "nombre": str(nombre).strip(),
                "turnos": turnos,
                "fila": fila
            })
        
        return agentes, wb, f"✅ Cargados {len(agentes)} agentes"
    
    except Exception as e:
        return None, None, f"❌ Error: {e}"


def guardar_en_excel(agentes, wb, archivo_path):
    """Guarda los turnos actuales en el Excel"""
    try:
        sheet = wb["MAYO 2026"]
        columnas_turnos = [6 + i*2 for i in range(31)]
        
        for agente in agentes:
            fila = agente["fila"]
            for i, turno in enumerate(agente["turnos"]):
                col = columnas_turnos[i]
                sheet.cell(row=fila, column=col).value = turno if turno else None
        
        wb.save(archivo_path)
        return True, "✅ Cambios guardados en Excel"
    except Exception as e:
        return False, f"❌ Error: {e}"


def intercambiar_turnos(agentes, idx1, idx2, dia):
    """Intercambia turnos entre dos agentes en memoria"""
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
if 'wb' not in st.session_state:
    st.session_state.wb = None
if 'archivo_path' not in st.session_state:
    st.session_state.archivo_path = None
if 'cargado' not in st.session_state:
    st.session_state.cargado = False
if 'hay_cambios' not in st.session_state:
    st.session_state.hay_cambios = False


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
        
        with st.spinner("Cargando datos..."):
            agentes, wb, msg = cargar_excel(archivo_temp)
        
        if agentes:
            st.session_state.agentes = agentes
            st.session_state.wb = wb
            st.session_state.archivo_path = archivo_temp
            st.session_state.cargado = True
            st.session_state.hay_cambios = False
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)
    
    if st.session_state.cargado:
        st.markdown("---")
        
        # Organizar por zona para el selector
        zonas = {}
        for ag in st.session_state.agentes:
            zona = ag["zona"]
            if zona not in zonas:
                zonas[zona] = []
            zonas[zona].append(ag)
        
        zona_seleccionada = st.selectbox("📍 Zona", list(zonas.keys()))
        st.session_state.agentes_filtrados = zonas[zona_seleccionada]
        
        st.metric("👥 Agentes", len(zonas[zona_seleccionada]))
        
        st.markdown("---")
        
        # Botón de guardado
        if st.session_state.hay_cambios:
            st.warning("⚠️ Hay cambios sin guardar")
            if st.button("💾 GUARDAR en Excel", type="primary", use_container_width=True):
                ok, msg = guardar_en_excel(
                    st.session_state.agentes,
                    st.session_state.wb,
                    st.session_state.archivo_path
                )
                if ok:
                    st.success(msg)
                    st.session_state.hay_cambios = False
                    st.rerun()
                else:
                    st.error(msg)
        else:
            st.success("✓ Sin cambios pendientes")


# ============================================================
# CONTENIDO PRINCIPAL
# ============================================================

if not st.session_state.cargado:
    st.info("👈 Carga el archivo Excel en el panel lateral")
else:
    agentes = st.session_state.agentes_filtrados
    zona = st.session_state.get('zona_seleccionada', '')
    
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
        "idx": st.column_config.NumberColumn("#", width="small"),
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
        turno1 = agentes[idx1]["turnos"][dia_idx]
        turno2 = agentes[idx2]["turnos"][dia_idx]
        st.info(f"**Turno actual:** {agentes[idx1]['nombre']} → {turno1 if turno1 else '—'} | {agentes[idx2]['nombre']} → {turno2 if turno2 else '—'}")
        
        if st.button("🔄 Intercambiar turnos", type="primary", use_container_width=True):
            if intercambiar_turnos(agentes, idx1, idx2, dia_idx):
                st.session_state.hay_cambios = True
                st.success("✅ Turnos intercambiados en memoria")
                st.rerun()
            else:
                st.error("Error al intercambiar")
    else:
        st.warning("Se necesitan al menos 2 agentes en esta zona para intercambiar turnos")
    
    # Mostrar estado de cambios
    if st.session_state.hay_cambios:
        st.info("💡 Recuerda hacer clic en 'GUARDAR en Excel' en el panel lateral para guardar los cambios permanentemente")
