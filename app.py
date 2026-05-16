import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile
import re

st.set_page_config(page_title="Cuadrantes Metrovalencia", layout="wide")

st.title("📅 Cuadrante de Servicios - Metrovalencia")
st.caption("Detección automática de columnas | Solo turno (primera columna de cada día)")

# ============================================================
# FUNCIONES PRINCIPALES
# ============================================================

def encontrar_fila_encabezados(sheet):
    """Busca la fila donde están los nombres de columna (COD., NOMBRE, 1, 2, 3...)"""
    for fila in range(1, min(20, sheet.max_row + 1)):
        fila_valores = []
        for col in range(1, min(80, sheet.max_column + 1)):
            celda = sheet.cell(row=fila, column=col).value
            if celda:
                fila_valores.append(str(celda).upper())
        
        # Buscar indicadores de fila de encabezados
        if any("NOMBRE" in v for v in fila_valores) or any("AGENTE" in v for v in fila_valores):
            # También debe tener números de días (1, 2, 3...)
            tiene_dias = any(v.isdigit() and 1 <= int(v) <= 31 for v in fila_valores)
            if tiene_dias:
                return fila
    return 2  # Por defecto, fila 2


def encontrar_columna_codigo(sheet, fila_encabezados):
    """Encuentra la columna donde está 'COD.' o 'CODIGO'"""
    for col in range(1, 20):
        celda = sheet.cell(row=fila_encabezados, column=col).value
        if celda and ("COD" in str(celda).upper() or "CÓD" in str(celda).upper()):
            return col
    return 3  # Por defecto, columna C


def encontrar_columna_nombre(sheet, fila_encabezados):
    """Encuentra la columna donde está 'NOMBRE' o 'AGENTE'"""
    for col in range(1, 30):
        celda = sheet.cell(row=fila_encabezados, column=col).value
        if celda and ("NOMBRE" in str(celda).upper() or "AGENTE" in str(celda).upper()):
            return col
    return 5  # Por defecto, columna E


def encontrar_columnas_dias(sheet, fila_encabezados):
    """Encuentra las columnas donde están los días (1, 2, 3...) y devuelve SOLO las primeras de cada par"""
    columnas_dias = []
    for col in range(1, 100):
        celda = sheet.cell(row=fila_encabezados, column=col).value
        if celda:
            # Verificar si es un número de día (1-31)
            try:
                num = int(str(celda).strip())
                if 1 <= num <= 31:
                    columnas_dias.append((col, num))
            except:
                pass
    
    # Ordenar por número de día
    columnas_dias.sort(key=lambda x: x[1])
    
    # Devolver SOLO las primeras de cada par (columnas impares)
    # Esto asume que los días están en columnas consecutivas: día1, (dato), día2, (dato)...
    # Tomamos las que están en posiciones impares de la secuencia
    primeras_columnas = []
    for i, (col, dia) in enumerate(columnas_dias):
        # Si es el primer día o la columna está separada por 2 de la anterior
        if i == 0:
            primeras_columnas.append((col, dia))
        else:
            col_anterior, _ = columnas_dias[i-1]
            # Si la columna actual está a 2 de distancia, es la primera de un nuevo par
            if col - col_anterior >= 2:
                primeras_columnas.append((col, dia))
            # Si está a 1 de distancia, es la segunda columna (la ignoramos)
    
    return primeras_columnas


def cargar_excel(archivo_path):
    """Carga el Excel detectando automáticamente las columnas"""
    try:
        wb = load_workbook(archivo_path, data_only=True)
        sheet = wb["MAYO 2026"]
        
        # Detectar estructura
        fila_encabezados = encontrar_fila_encabezados(sheet)
        col_codigo = encontrar_columna_codigo(sheet, fila_encabezados)
        col_nombre = encontrar_columna_nombre(sheet, fila_encabezados)
        columnas_dias = encontrar_columnas_dias(sheet, fila_encabezados)
        
        st.sidebar.success(f"🔍 Detectado: fila {fila_encabezados}, código col {col_codigo}, nombre col {col_nombre}")
        st.sidebar.write(f"📅 Días encontrados: {len(columnas_dias)}")
        
        if not columnas_dias:
            return None, None, "No se encontraron columnas de días"
        
        agentes = []
        col_zona = 1  # Columna A (identificador JC, AV, etc.)
        
        for fila in range(fila_encabezados + 1, sheet.max_row + 1):
            zona = sheet.cell(row=fila, column=col_zona).value
            codigo = sheet.cell(row=fila, column=col_codigo).value
            nombre = sheet.cell(row=fila, column=col_nombre).value
            
            # Filtrar agentes válidos
            if not codigo or not nombre:
                continue
            if str(codigo).strip() == "0" or str(codigo).strip() == "":
                continue
            
            nombre_str = str(nombre).strip()
            if "DESPLAZADO" in nombre_str.upper() or "VACANTE" in nombre_str.upper():
                continue
            
            # Extraer turnos SOLO de las primeras columnas de cada día
            turnos = []
            celdas_turnos = []
            
            for col, dia in columnas_dias:
                celda = sheet.cell(row=fila, column=col)
                valor = celda.value
                turnos.append(str(valor).strip() if valor else "")
                celdas_turnos.append(celda)
            
            # Asegurar 31 días
            while len(turnos) < 31:
                turnos.append("")
                celdas_turnos.append(None)
            
            agentes.append({
                "zona": str(zona).strip() if zona else "",
                "codigo": str(codigo).strip(),
                "nombre": nombre_str,
                "turnos": turnos[:31],
                "celdas_turnos": celdas_turnos[:31],
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
        
        return agentes_por_zona, wb, f"✅ Cargados {len(agentes)} agentes en {len(agentes_por_zona)} zonas"
    
    except Exception as e:
        return None, None, f"❌ Error: {e}"


def intercambiar_turnos(agentes, idx1, idx2, dia):
    """Intercambia SOLO el turno entre dos agentes"""
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
    if ag1["celdas_turnos"][dia] and ag2["celdas_turnos"][dia]:
        celda1 = ag1["celdas_turnos"][dia]
        celda2 = ag2["celdas_turnos"][dia]
        celda1.value = turno2 if turno2 else None
        celda2.value = turno1 if turno1 else None
    
    return True


def guardar_excel(wb, archivo_path):
    """Guarda los cambios en el archivo"""
    try:
        wb.save(archivo_path)
        return True, "✅ Cambios guardados en el archivo"
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
        
        with st.spinner("Cargando datos..."):
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
        if zonas:
            zona_seleccionada = st.selectbox("📍 Zona", zonas)
            st.session_state['zona'] = zona_seleccionada
            
            agentes = st.session_state['agentes_por_zona'][zona_seleccionada]
            st.metric("👥 Agentes", len(agentes))

# Contenido principal
if not st.session_state.get('cargado', False):
    st.info("👈 Carga el archivo Excel en el panel lateral")
    with st.expander("📖 Instrucciones"):
        st.markdown("""
        ### ¿Cómo funciona?
        1. Carga el archivo **AÑO 2026 ESTACIONES .xlsx**
        2. La aplicación detecta automáticamente:
           - La fila de encabezados
           - Las columnas de código y nombre
           - Las columnas de días (1 al 31)
        3. **Solo lee la primera columna de cada día** (el turno)
        4. Puedes intercambiar turnos entre agentes de la misma zona
        5. Los cambios se guardan directamente en el Excel
        """)
else:
    zona = st.session_state.get('zona', list(st.session_state['agentes_por_zona'].keys())[0])
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
    st.caption("Selecciona dos agentes y un día para intercambiar sus turnos")
    
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
        st.info(f"**Turno actual:** {agentes[idx1]['nombre']} → {turno1 if turno1 else '—'} | {agentes[idx2]['nombre']} → {turno2 if turno2 else '—'}")
        
        if st.button("🔄 Intercambiar y guardar en Excel", type="primary", use_container_width=True):
            if intercambiar_turnos(agentes, idx1, idx2, dia_idx):
                ok, msg = guardar_excel(st.session_state['wb'], st.session_state['archivo_path'])
                if ok:
                    st.success(msg)
                    st.balloons()
                    st.rerun()
                else:
                    st.error(msg)
            else:
                st.error("Error al intercambiar los turnos")
    else:
        st.warning("Se necesitan al menos 2 agentes en esta zona para intercambiar turnos")
