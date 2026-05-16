import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile
import os

st.set_page_config(page_title="Cuadrantes Metrovalencia", layout="wide")

st.title("📅 Cuadrante de Servicios - Metrovalencia")
st.caption("Edición en tiempo real | Guardado automático en Excel")

# ============================================================
# FUNCIONES
# ============================================================

def cargar_excel(archivo_path):
    """Carga todos los agentes desde el Excel"""
    try:
        wb = load_workbook(archivo_path, data_only=True)
        sheet = wb["MAYO 2026"]
        
        # Según el archivo: fila 9 tiene los días (1,2,3...)
        # fila 11 tiene los encabezados (COD., NOMBRE, etc.)
        # Los agentes empiezan en fila 12
        
        agentes = []
        
        # Recorrer desde fila 12 hasta el final
        for fila in range(12, sheet.max_row + 1):
            # Columna A = zona (JC, AV, AA, FO, etc.)
            zona = sheet.cell(row=fila, column=1).value
            # Columna C = código del agente
            codigo = sheet.cell(row=fila, column=3).value
            # Columna E = nombre del agente
            nombre = sheet.cell(row=fila, column=5).value
            
            # Saltar filas vacías
            if not codigo or not nombre:
                continue
            if str(codigo).strip() == "0":
                continue
            
            nombre_str = str(nombre).strip()
            # Excluir desplazados y vacantes
            if "DESPLAZADO" in nombre_str.upper() or "VACANTE" in nombre_str.upper():
                continue
            
            # Leer turnos: columna F = día 1, G = dato extra, H = día 2, I = dato extra, etc.
            # Solo nos interesa la primera de cada par (F, H, J, L...)
            turnos = []
            celdas_turnos = []
            
            for dia in range(31):
                col_turno = 6 + (dia * 2)  # F=6, H=8, J=10, L=12...
                celda = sheet.cell(row=fila, column=col_turno)
                valor = celda.value
                turnos.append(str(valor).strip() if valor else "")
                celdas_turnos.append(celda)
            
            agentes.append({
                "zona": str(zona).strip() if zona else "",
                "codigo": str(codigo).strip(),
                "nombre": nombre_str,
                "turnos": turnos,
                "celdas_turnos": celdas_turnos,
                "fila": fila
            })
        
        if not agentes:
            return None, None, "No se encontraron agentes"
        
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


def guardar_todos_los_cambios(agentes_por_zona, wb, archivo_path):
    """Guarda todos los turnos modificados en el Excel"""
    try:
        for zona, agentes in agentes_por_zona.items():
            for agente in agentes:
                for i, turno in enumerate(agente["turnos"]):
                    if agente["celdas_turnos"][i]:
                        agente["celdas_turnos"][i].value = turno if turno else None
        
        wb.save(archivo_path)
        return True, "✅ Cambios guardados en Excel"
    except Exception as e:
        return False, f"❌ Error al guardar: {e}"


def intercambiar_turnos(agentes, idx1, idx2, dia):
    """Intercambia turnos entre dos agentes y guarda inmediatamente"""
    if idx1 == idx2:
        return False
    
    ag1 = agentes[idx1]
    ag2 = agentes[idx2]
    
    turno1 = ag1["turnos"][dia]
    turno2 = ag2["turnos"][dia]
    
    # Intercambiar en memoria
    ag1["turnos"][dia] = turno2
    ag2["turnos"][dia] = turno1
    
    # Intercambiar en las celdas de Excel (para guardado rápido)
    if ag1["celdas_turnos"][dia] and ag2["celdas_turnos"][dia]:
        ag1["celdas_turnos"][dia].value = turno2 if turno2 else None
        ag2["celdas_turnos"][dia].value = turno1 if turno1 else None
    
    return True


# ============================================================
# INICIALIZACIÓN
# ============================================================

if 'agentes_por_zona' not in st.session_state:
    st.session_state.agentes_por_zona = None
if 'wb' not in st.session_state:
    st.session_state.wb = None
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
        
        with st.spinner("Cargando agentes..."):
            agentes_por_zona, wb, msg = cargar_excel(archivo_temp)
        
        if agentes_por_zona:
            st.session_state.agentes_por_zona = agentes_por_zona
            st.session_state.wb = wb
            st.session_state.archivo_path = archivo_temp
            st.session_state.cargado = True
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)
    
    if st.session_state.cargado:
        st.markdown("---")
        zonas = list(st.session_state.agentes_por_zona.keys())
        zona_seleccionada = st.selectbox("📍 Zona", zonas)
        st.session_state.zona_actual = zona_seleccionada
        
        agentes = st.session_state.agentes_por_zona[zona_seleccionada]
        st.metric("👥 Agentes", len(agentes))
        
        st.markdown("---")
        st.success("💾 Los cambios se guardan automáticamente en Excel")


# ============================================================
# CONTENIDO PRINCIPAL
# ============================================================

if not st.session_state.cargado:
    st.info("👈 Carga el archivo Excel en el panel lateral")
    
    with st.expander("📖 Estructura esperada"):
        st.markdown("""
        ### El archivo debe contener:
        - Hoja **MAYO 2026**
        - **Fila 9**: Números de día (1, 2, 3...)
        - **Fila 11**: Encabezados (COD., NOMBRE, etc.)
        - **Fila 12 en adelante**: Agentes con sus turnos
        
        ### Columnas:
        - **A**: Zona (JC, AV, AA, FO, etc.)
        - **C**: Código del agente
        - **E**: Nombre del agente
        - **F, H, J, L...**: Turnos (solo las impares)
        - **G, I, K, M...**: Datos adicionales (se ignoran)
        """)
else:
    zona = st.session_state.zona_actual
    agentes = st.session_state.agentes_por_zona[zona]
    
    # Construir DataFrame para mostrar
    dias = list(range(1, 32))
    data = []
    
    for i, ag in enumerate(agentes):
        fila = {"idx": i, "Código": ag["codigo"], "Agente": ag["nombre"]}
        for d in range(31):
            turno = ag["turnos"][d]
            fila[f"D{d+1}"] = turno if turno else "—"
        data.append(fila)
    
    df = pd.DataFrame(data)
    
    st.markdown(f"## 📊 Zona {zona}")
    
    # Configurar columnas para mejor visualización
    column_config = {
        "idx": st.column_config.NumberColumn("#", width="small"),
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
        turno1 = agentes[idx1]["turnos"][dia_idx]
        turno2 = agentes[idx2]["turnos"][dia_idx]
        
        st.info(f"📌 **Actual:** {agentes[idx1]['nombre']} → `{turno1 if turno1 else '—'}` | {agentes[idx2]['nombre']} → `{turno2 if turno2 else '—'}`")
        
        if st.button("🔄 Intercambiar turnos", type="primary", use_container_width=True):
            if intercambiar_turnos(agentes, idx1, idx2, dia_idx):
                # Guardar en Excel inmediatamente
                ok, msg = guardar_todos_los_cambios(
                    st.session_state.agentes_por_zona,
                    st.session_state.wb,
                    st.session_state.archivo_path
                )
                if ok:
                    st.success("✅ Turnos intercambiados y guardados en Excel")
                    st.balloons()
                    st.rerun()
                else:
                    st.error(msg)
            else:
                st.error("Error al intercambiar")
    else:
        st.warning("Se necesitan al menos 2 agentes en esta zona")
    
    # Mostrar estadísticas de la zona
    with st.expander("📊 Estadísticas de la zona"):
        st.markdown(f"**Total agentes:** {len(agentes)}")
        
        # Contar tipos de turno más comunes
        todos_turnos = []
        for ag in agentes:
            todos_turnos.extend([t for t in ag["turnos"] if t])
        
        if todos_turnos:
            from collections import Counter
            counter = Counter(todos_turnos)
            st.markdown("**Turnos más frecuentes:**")
            for turno, count in counter.most_common(10):
                st.write(f"- `{turno}`: {count} veces")
