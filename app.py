import streamlit as st
import pandas as pd
from io import BytesIO

# ------------------------------------------------------------
# CONFIGURACIÓN INICIAL
# ------------------------------------------------------------
EXCEL_FILE = "AÑO 2026 ESTACIONES .xlsx"
SHEET_NAME = "MAYO 2026"

# ------------------------------------------------------------
# FUNCIONES DE CARGA DE DATOS
# ------------------------------------------------------------
@st.cache_data
def cargar_datos():
    """Carga el cuadrante de Mayo 2026, extrae agentes y servicios."""
    df_raw = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, header=None)

    # 1. Encontrar la fila donde empiezan los agentes (identificada por "COD." en columna 1)
    start_row = None
    for i, row in df_raw.iterrows():
        if pd.notna(row[1]) and str(row[1]).strip() == "COD.":
            start_row = i + 1      # la siguiente fila contiene los datos
            break
    if start_row is None:
        st.error("No se encontró la cabecera de agentes en el archivo.")
        st.stop()

    # 2. Identificar las columnas de los días (primera columna de cada día)
    #    En la fila 3 (índice 3) están los números de día.
    dias_cols = {}
    for col in range(5, len(df_raw.columns), 2):   # empezamos en col F (índice 5), saltos de 2
        dia_val = df_raw.iloc[3, col]
        if pd.notna(dia_val) and str(dia_val).isdigit():
            dias_cols[int(dia_val)] = col

    # 3. Extraer datos de agentes: filas desde start_row hasta donde haya datos
    agentes = []
    for i in range(start_row, len(df_raw)):
        # Si la primera columna está vacía y la tercera también, terminamos
        if pd.isna(df_raw.iloc[i, 0]) and pd.isna(df_raw.iloc[i, 2]):
            break
        # Código de zona (columna A)
        zona = str(df_raw.iloc[i, 0]).strip() if pd.notna(df_raw.iloc[i, 0]) else ""
        if not zona:
            continue
        # Código de agente (columna C)
        cod_agente = int(df_raw.iloc[i, 2]) if pd.notna(df_raw.iloc[i, 2]) else None
        if cod_agente is None:
            continue
        nombre = str(df_raw.iloc[i, 3]).strip() if pd.notna(df_raw.iloc[i, 3]) else ""

        # Leer los servicios para cada día
        servicios = {}
        for dia, col_serv in dias_cols.items():
            col_comp = col_serv + 1   # columna complementaria (la segunda del día)
            serv = df_raw.iloc[i, col_serv] if pd.notna(df_raw.iloc[i, col_serv]) else ""
            servicios[dia] = serv
        agentes.append({
            "zona": zona,
            "codigo": cod_agente,
            "nombre": nombre,
            "servicios": servicios,
            "fila_excel": i       # guardamos la fila para luego escribir cambios
        })

    return agentes, dias_cols, df_raw

# ------------------------------------------------------------
# FUNCIONES DE GUARDADO
# ------------------------------------------------------------
def guardar_cambios(agentes, dias_cols, df_raw):
    """Vuelca los servicios de los agentes modificados al DataFrame original y guarda en Excel."""
    for ag in agentes:
        fila = ag["fila_excel"]
        for dia, col_serv in dias_cols.items():
            df_raw.iloc[fila, col_serv] = ag["servicios"][dia]
    # Guardar en Excel (se crea un buffer en memoria para no sobreescribir durante la ejecución)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_raw.to_excel(writer, sheet_name=SHEET_NAME, index=False, header=False)
    # Descargar el archivo modificado (o podemos guardar en disco)
    st.session_state["excel_modificado"] = output.getvalue()

# ------------------------------------------------------------
# INTERFAZ STREAMLIT
# ------------------------------------------------------------
st.set_page_config(page_title="Cuadrante de Servicios - Línea 3", layout="wide")

# Inicializar session_state
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.agente = None
    st.session_state.agentes = None
    st.session_state.dias_cols = None
    st.session_state.df_raw = None

# Cargar datos una sola vez
if st.session_state.agentes is None:
    agentes, dias_cols, df_raw = cargar_datos()
    st.session_state.agentes = agentes
    st.session_state.dias_cols = dias_cols
    st.session_state.df_raw = df_raw

# ------------------------------------------------------------
# LOGIN
# ------------------------------------------------------------
if not st.session_state.autenticado:
    st.title("🔐 Acceso al cuadrante de servicios")
    codigo = st.text_input("Código de agente")
    password = st.text_input("Contraseña", type="password")
    if st.button("Entrar"):
        # Buscar el agente por código
        agente = next((a for a in st.session_state.agentes if str(a["codigo"]) == codigo), None)
        if agente and password == codigo:   # contraseña por defecto = código
            st.session_state.autenticado = True
            st.session_state.agente = agente
            st.rerun()
        else:
            st.error("Código o contraseña incorrectos")
    st.stop()

# ------------------------------------------------------------
# APLICACIÓN PRINCIPAL
# ------------------------------------------------------------
agente = st.session_state.agente
agentes = st.session_state.agentes
dias_cols = st.session_state.dias_cols

st.title(f"📅 Cuadrante de Mayo 2026")
st.subheader(f"Bienvenido, {agente['nombre']} (Código {agente['codigo']}) - Zona {agente['zona']}")

# Mostrar los agentes de la misma zona
misma_zona = [a for a in agentes if a["zona"] == agente["zona"]]
st.markdown("---")
st.write("### 📊 Agentes de tu zona")

# Tabla comparativa de servicios (solo días 1 a 31)
dias = sorted(dias_cols.keys())
# Construir DataFrame para mostrar
data = []
for a in misma_zona:
    row = [a["nombre"], a["codigo"]]
    for d in dias:
        row.append(a["servicios"].get(d, ""))
    data.append(row)

df_show = pd.DataFrame(data, columns=["Nombre", "Código"] + [f"Día {d}" for d in dias])
st.dataframe(df_show, use_container_width=True)

# ------------------------------------------------------------
# INTERCAMBIO DE SERVICIOS
# ------------------------------------------------------------
st.markdown("---")
st.write("### 🔄 Intercambiar servicio con otro agente")

# Selección de día
dia_sel = st.selectbox("Día", dias, format_func=lambda x: f"Día {x}")

# Servicio actual del agente logueado en ese día
servicio_actual = agente["servicios"][dia_sel]
st.info(f"Tu servicio actual el día {dia_sel}: **{servicio_actual}**")

# Selección de otro agente de la misma zona (excluyendo al mismo)
otros_agentes = [a for a in misma_zona if a["codigo"] != agente["codigo"]]
if otros_agentes:
    agente_destino = st.selectbox("Agente con quien intercambiar", otros_agentes,
                                   format_func=lambda x: f"{x['nombre']} (Código {x['codigo']})")
    if st.button("Intercambiar"):
        # Obtener servicio del destino en el mismo día
        serv_destino = agente_destino["servicios"][dia_sel]
        # Realizar intercambio
        agente["servicios"][dia_sel], agente_destino["servicios"][dia_sel] = serv_destino, servicio_actual
        # Guardar cambios
        guardar_cambios(st.session_state.agentes, st.session_state.dias_cols, st.session_state.df_raw)
        st.success(f"Intercambio realizado: ahora tienes '{serv_destino}' y {agente_destino['nombre']} tiene '{servicio_actual}'")
        st.rerun()
else:
    st.warning("No hay otros agentes en tu zona para intercambiar.")

# ------------------------------------------------------------
# DESCARGA DEL CUADRANTE MODIFICADO
# ------------------------------------------------------------
st.markdown("---")
if "excel_modificado" in st.session_state:
    st.download_button(
        label="📥 Descargar archivo Excel actualizado",
        data=st.session_state.excel_modificado,
        file_name="cuadrante_modificado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Los cambios aún no se han guardado en un archivo descargable. Realiza un intercambio para generar el archivo.")