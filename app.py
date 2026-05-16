import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile
import os

st.set_page_config(page_title="Cuadrantes Metrovalencia", layout="wide")

st.title("📅 Cuadrante de Servicios - Metrovalencia")
st.caption("Detección manual de filas | Guardado en archivo original")

# ============================================================
# FUNCIONES
# ============================================================

def mostrar_primeras_filas(archivo_path):
    """Muestra las primeras filas para identificar la estructura"""
    wb = load_workbook(archivo_path, data_only=True)
    sheet = wb["MAYO 2026"]
    
    datos = []
    for fila in range(1, 25):
        fila_datos = []
        for col in range(1, 10):
            celda = sheet.cell(row=fila, column=col).value
            fila_datos.append(str(celda)[:30] if celda else "")
        datos.append(fila_datos)
    
    return pd.DataFrame(datos, columns=["A", "B", "C", "D", "E", "F", "G", "H", "I"])


def cargar_excel(archivo_path, fila_inicio):
    """Carga agentes desde la fila especificada"""
    try:
        wb = load_workbook(archivo_path, data_only=True)
        sheet = wb["MAYO 2026"]
        
        agentes = []
        
        # Columnas de turnos: F(6), H(8), J(10), L(12), N(14), P(16), R(18), T(20), V(22), X(24), Z(26), AB(28), AD(30), AF(32), AH(34), AJ(36), AL(38), AN(40), AP(42), AR(44), AT(46), AV(48), AX(50), AZ(52), BB(54), BD(56), BF(58), BH(60), BJ(62), BL(64), BN(66)
        columnas_turnos = [6 + i*2 for i in range(31)]
        
        for fila in range(fila_inicio, sheet.max_row + 1):
            zona = sheet.cell(row=fila, column=1).value
            codigo = sheet.cell(row=fila, column=3).value
            nombre = sheet.cell(row=fila, column=5).value
            
            # Saltar filas vacías o inválidas
            if not codigo or not nombre:
                continue
            if str(codigo).strip() == "0" or str(codigo).strip() == "":
                continue
            
            nombre_str = str(nombre).strip()
            # Excluir desplazados y vacantes
            if "DESPLAZADO" in nombre_str.upper() or "VACANTE" in nombre_str.upper():
                continue
            
            # Leer turnos
            turnos = []
            celdas_turnos = []
            
            for col in columnas_turnos:
                if col <= sheet.max_column:
                    celda = sheet.cell(row=fila, column=col)
                    valor = celda.value
                    turnos.append(str(valor).strip() if valor else "")
                    celdas_turnos.append(celda)
                else:
                    turnos.append("")
                    celdas_turnos.append(None)
            
            # Asegurar 31 días
            while len(turnos) < 31:
                turnos.append("")
            
            agente = {
                "zona": str(zona).strip() if zona else "",
                "codigo": str(codigo).strip(),
                "nombre": nombre_str,
                "turnos": turnos[:31],
                "celdas_turnos": celdas_turnos[:31],
                "fila": fila
            }
            agentes.append(agente)
        
        return agentes, wb, f"✅ Cargados {len(agentes)} agentes desde fila {fila_inicio}"
    
    except Exception as e:
        return None, None, f"❌ Error: {e}"


def guardar_todos_los_cambios(agentes, wb, archivo_path):
    """Guarda todos los turnos modificados en el Excel original"""
    try:
        for agente in agentes:
            for i, turno in enumerate(agente["turnos"]):
                if agente["celdas_turnos"][i]:
                    agente["celdas_turnos"][i].value = turno if turno else None
        
        wb.save(archivo_path)
        return True, f"✅ Cambios guardados en {archivo_path}"
    except Exception as e:
        return False, f"❌ Error al guardar: {e}"


def intercambiar_turnos(agentes, idx1, idx2, dia):
    """Intercambia turnos entre dos agentes"""
    if idx1 == idx2:
        return False
    
    ag1 = agentes[idx1]
    ag2 = agentes[idx2]
    
    turno1 = ag1["turnos"][dia]
    turno2 = ag2["turnos"][dia]
    
    ag1["turnos"][dia] = turno2
    ag2["turnos"][dia] = turno1
    
    if ag1["celdas_turnos"][dia]:
        ag1["celdas_turnos"][dia].value = turno2 if turno2 else None
    if ag2["celdas_turnos"][dia]:
        ag2["celdas_turnos"][dia].value = turno1 if turno1 else None
    
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


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("📁 Cargar archivo")
    archivo_subido = st.file_uploader("Selecciona el Excel", type=["xlsx"])
    
    if archivo_subido:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(archivo_subido.getvalue())
            archivo_temp = tmp.name
        
        st.session_state.archivo_path = archivo_temp
        
        # Mostrar primeras filas para identificar estructura
        st.markdown("---")
        st.subheader("🔍 Vista previa (filas 1-24)")
        
        df_vista = mostrar_primeras_filas(archivo_temp)
        st.dataframe(df_vista, use_container_width=True)
        
        st.markdown("---")
        st.subheader("⚙️ Configuración")
        
        # Seleccionar fila donde empiezan los agentes
        fila_inicio = st.number_input("Fila donde empiezan los agentes", min_value=1, max_value=50, value=12, step=1)
        
        if st.button("📊 Cargar agentes", type="primary", use_container_width=True):
            with st.spinner("Cargando..."):
                agentes, wb, msg = cargar_excel(archivo_temp, fila_inicio)
            
            if agentes:
                st.session_state.agentes = agentes
                st.session_state.wb = wb
                st.session_state.cargado = True
                st.success(msg)
                
                # Contar por zona
                zonas = {}
                for ag in agentes:
                    zona = ag["zona"]
                    zonas[zona] = zonas.get(zona, 0) + 1
                
                st.write("**Distribución por zona:**")
                for zona, count in zonas.items():
                    st.write(f"- {zona}: {count} agentes")
                
                st.rerun()
            else:
                st.error(msg)
    
    if st.session_state.cargado:
        st.markdown("---")
        
        # Selector de zona
        zonas = {}
        for ag in st.session_state.agentes:
            zona = ag["zona"]
            if zona not in zonas:
                zonas[zona] = []
            zonas[zona].append(ag)
        
        zona_seleccionada = st.selectbox("📍 Zona", list(zonas.keys()))
        st.session_state.agentes_filtrados = zonas[zona_seleccionada]
        
        st.metric("👥 Agentes en zona", len(zonas[zona_seleccionada]))
        
        st.markdown("---")
        st.info("💾 Los cambios se guardan automáticamente en el archivo original")
        
        # Botón para forzar guardado
        if st.button("💾 Forzar guardado en Excel", use_container_width=True):
            ok, msg = guardar_todos_los_cambios(
                st.session_state.agentes,
                st.session_state.wb,
                st.session_state.archivo_path
            )
            if ok:
                st.success(msg)
            else:
                st.error(msg)


# ============================================================
# CONTENIDO PRINCIPAL
# ============================================================

if not st.session_state.cargado:
    st.info("👈 Carga el archivo Excel y selecciona la fila donde empiezan los agentes")
    
    with st.expander("📖 Instrucciones"):
        st.markdown("""
        ### ¿Cómo identificar la fila correcta?
        
        1. Mira la tabla de **Vista previa** en el panel lateral
        2. Busca la fila donde aparecen los primeros agentes con código y nombre
        3. Según el archivo, los agentes empiezan en **fila 12**
        4. Ajusta el número si ves que faltan agentes o hay filas vacías
        
        ### Columnas:
        - **Columna A**: Zona (JC, AV, AA, FO, MS, etc.)
        - **Columna C**: Código del agente
        - **Columna E**: Nombre del agente
        - **Columna F, H, J...**: Turnos (solo las impares)
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
            turno = ag["turnos"][d]
            fila[f"D{d+1}"] = turno if turno else "—"
        data.append(fila)
    
    df = pd.DataFrame(data)
    
    st.markdown(f"## 📊 Zona {zona}")
    st.caption(f"Total agentes en esta zona: {len(agentes)}")
    
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
        turno1 = agentes[idx1]["turnos"][dia_idx]
        turno2 = agentes[idx2]["turnos"][dia_idx]
        
        st.info(f"📌 **Actual:** {agentes[idx1]['nombre']} → `{turno1 if turno1 else '—'}` | {agentes[idx2]['nombre']} → `{turno2 if turno2 else '—'}`")
        
        if st.button("🔄 Intercambiar y guardar", type="primary", use_container_width=True):
            if intercambiar_turnos(agentes, idx1, idx2, dia_idx):
                ok, msg = guardar_todos_los_cambios(
                    st.session_state.agentes,
                    st.session_state.wb,
                    st.session_state.archivo_path
                )
                if ok:
                    st.success("✅ Turnos intercambiados y guardados")
                    st.balloons()
                    st.rerun()
                else:
                    st.error(msg)
            else:
                st.error("Error al intercambiar")
    else:
        st.warning("Se necesitan al menos 2 agentes en esta zona")
