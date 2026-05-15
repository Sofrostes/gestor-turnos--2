import streamlit as st
import pandas as pd
from io import BytesIO
import numpy as np

# Configuración de la página
st.set_page_config(page_title="Gestor de Cuadrantes", layout="wide")
st.title("📅 Gestor de Cuadrantes - Mayo 2026")
st.markdown("Intercambia turnos entre empleados del mismo grupo (columna A)")

# --------------------------------------------
# Funciones de carga y procesamiento
# --------------------------------------------
def cargar_excel(uploaded_file):
    """Carga la hoja 'MAYO 2026' y extrae la tabla de empleados y turnos."""
    xls = pd.ExcelFile(uploaded_file)
    if "MAYO 2026" not in xls.sheet_names:
        st.error("La hoja 'MAYO 2026' no existe en el archivo.")
        return None
    # Leer desde la fila donde comienzan los datos (fila índice 11 en Excel)
    df_raw = pd.read_excel(uploaded_file, sheet_name="MAYO 2026", header=11, dtype=str)
    
    # Limpiar columnas: eliminar las que son completamente NaN
    df_raw = df_raw.dropna(axis=1, how='all')
    # La columna A es el grupo, D el nombre, y de la F en adelante son los días
    if len(df_raw.columns) < 6:
        st.error("No se detectaron columnas de turnos.")
        return None
    
    # Identificar las columnas de servicio y detalle por día
    # Las columnas pares a partir de la columna 5 (índice 5 = F) son servicio
    # Las impares son detalle
    col_servicio = []
    col_detalle = []
    for i in range(5, len(df_raw.columns), 2):
        if i+1 < len(df_raw.columns):
            col_servicio.append(df_raw.columns[i])
            col_detalle.append(df_raw.columns[i+1])
        else:
            # Si hay una columna impar sin su par, la ignoramos
            pass
    
    num_dias = len(col_servicio)
    if num_dias != 31:
        st.warning(f"Se encontraron {num_dias} días. Se esperaban 31.")
    
    # Extraer datos relevantes
    grupos = df_raw.iloc[:, 0].fillna('').astype(str).tolist()
    nombres = df_raw.iloc[:, 3].fillna('').astype(str).tolist()
    # Extraer servicios y detalles como DataFrames
    servicios = df_raw[col_servicio].fillna('').astype(str)
    detalles = df_raw[col_detalle].fillna('').astype(str)
    
    # Crear lista de empleados (índice, grupo, nombre)
    empleados = []
    for idx in range(len(grupos)):
        if nombres[idx] and grupos[idx] not in ('0', '', 'nan'):
            empleados.append({
                "indice": idx,
                "grupo": grupos[idx],
                "nombre": nombres[idx]
            })
    
    if not empleados:
        st.error("No se encontraron empleados con nombre y grupo válidos.")
        return None
    
    return {
        "df_raw": df_raw,
        "servicios": servicios,
        "detalles": detalles,
        "col_servicio": col_servicio,
        "col_detalle": col_detalle,
        "empleados": empleados,
        "grupos": sorted(list(set(e["grupo"] for e in empleados)))
    }

def mostrar_cuadrante(data):
    """Muestra el cuadrante interactivo con tooltips (servicio + detalle)."""
    st.subheader("📊 Cuadrante de Turnos (Mayo 2026)")
    empleados = data["empleados"]
    servicios = data["servicios"]
    detalles = data["detalles"]
    
    # Crear dataframe para mostrar
    display_df = pd.DataFrame(index=[e["nombre"] for e in empleados], columns=range(1, 32))
    for i, emp in enumerate(empleados):
        idx = emp["indice"]
        for dia in range(1, 32):
            col_idx = dia - 1
            serv = servicios.iloc[idx, col_idx]
            det = detalles.iloc[idx, col_idx]
            # Texto a mostrar: servicio, y tooltip con detalle si no está vacío
            tooltip = f"{serv} ({det})" if det and det != 'nan' else serv
            display_df.iloc[i, dia-1] = serv if serv else ''
            # No podemos poner tooltip directamente en st.dataframe, usaremos st.write con estilo
    # Usar st.dataframe con highlight
    st.dataframe(display_df, use_container_width=True, height=600)
    
    # Opcional: mostrar tabla con más detalles (servicio + detalle)
    with st.expander("Ver detalles completos (servicio + detalle)"):
        detalle_df = pd.DataFrame(index=[e["nombre"] for e in empleados])
        for dia in range(1, 32):
            detalle_df[f"Día {dia}"] = [
                f"{servicios.iloc[emp['indice'], dia-1]} / {detalles.iloc[emp['indice'], dia-1]}"
                for emp in empleados
            ]
        st.dataframe(detalle_df, use_container_width=True)

def validar_intercambio(emp1, emp2, dia, servicios1, servicios2, detalles1, detalles2):
    """
    Aquí se implementarán las reglas de negocio.
    Por defecto, cualquier intercambio está permitido.
    Devuelve (booleano, mensaje).
    """
    # Ejemplo de regla: no intercambiar si algún turno es 'E' (enfermedad)
    # if servicios1 == 'E' or servicios2 == 'E':
    #     return False, "No se pueden intercambiar turnos con enfermedad (E)"
    return True, "Intercambio permitido"

def realizar_intercambio(data, emp1_idx, emp2_idx, dia):
    """
    Intercambia las columnas de servicio y detalle del día dado
    entre dos empleados.
    """
    dia_col = dia - 1
    serv1 = data["servicios"].iloc[emp1_idx, dia_col]
    serv2 = data["servicios"].iloc[emp2_idx, dia_col]
    det1 = data["detalles"].iloc[emp1_idx, dia_col]
    det2 = data["detalles"].iloc[emp2_idx, dia_col]
    
    # Intercambiar en los DataFrames
    data["servicios"].iloc[emp1_idx, dia_col] = serv2
    data["servicios"].iloc[emp2_idx, dia_col] = serv1
    data["detalles"].iloc[emp1_idx, dia_col] = det2
    data["detalles"].iloc[emp2_idx, dia_col] = det1
    
    # Actualizar también el df_raw para guardar después
    serv_col = data["col_servicio"][dia_col]
    det_col = data["col_detalle"][dia_col]
    data["df_raw"].iloc[emp1_idx, data["df_raw"].columns.get_loc(serv_col)] = serv2
    data["df_raw"].iloc[emp1_idx, data["df_raw"].columns.get_loc(det_col)] = det2
    data["df_raw"].iloc[emp2_idx, data["df_raw"].columns.get_loc(serv_col)] = serv1
    data["df_raw"].iloc[emp2_idx, data["df_raw"].columns.get_loc(det_col)] = det1

def guardar_excel(data, uploaded_file):
    """Guarda los cambios en el archivo Excel y devuelve bytes para descargar."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        data["df_raw"].to_excel(writer, sheet_name="MAYO 2026", index=False, header=True)
    return output.getvalue()

# --------------------------------------------
# Interfaz de usuario
# --------------------------------------------
uploaded_file = st.file_uploader("Sube el archivo Excel del cuadrante", type=["xlsx"])

if uploaded_file is not None:
    data = cargar_excel(uploaded_file)
    if data is None:
        st.stop()
    
    # Mostrar cuadrante
    mostrar_cuadrante(data)
    
    # Panel de intercambio
    st.subheader("🔄 Intercambio de turnos")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        grupo_seleccionado = st.selectbox("Grupo", data["grupos"])
        empleados_grupo = [e for e in data["empleados"] if e["grupo"] == grupo_seleccionado]
        nombres_grupo = [e["nombre"] for e in empleados_grupo]
        if len(nombres_grupo) < 2:
            st.warning("El grupo necesita al menos dos empleados para intercambiar.")
        else:
            emp1_nombre = st.selectbox("Empleado 1", nombres_grupo, key="emp1")
            emp2_nombre = st.selectbox("Empleado 2", nombres_grupo, key="emp2")
    
    with col2:
        dia_seleccionado = st.number_input("Día", min_value=1, max_value=31, step=1, value=1)
    
    with col3:
        if st.button("Intercambiar turnos"):
            # Obtener índices
            emp1 = next(e for e in empleados_grupo if e["nombre"] == emp1_nombre)
            emp2 = next(e for e in empleados_grupo if e["nombre"] == emp2_nombre)
            if emp1["indice"] == emp2["indice"]:
                st.error("No se puede intercambiar con uno mismo.")
            else:
                serv1 = data["servicios"].iloc[emp1["indice"], dia_seleccionado-1]
                serv2 = data["servicios"].iloc[emp2["indice"], dia_seleccionado-1]
                valido, msg = validar_intercambio(
                    emp1, emp2, dia_seleccionado,
                    serv1, serv2,
                    data["detalles"].iloc[emp1["indice"], dia_seleccionado-1],
                    data["detalles"].iloc[emp2["indice"], dia_seleccionado-1]
                )
                if not valido:
                    st.error(msg)
                else:
                    realizar_intercambio(data, emp1["indice"], emp2["indice"], dia_seleccionado)
                    st.success(f"✅ Intercambio realizado entre {emp1_nombre} y {emp2_nombre} para el día {dia_seleccionado}")
                    # Registrar en sesión
                    if "log" not in st.session_state:
                        st.session_state.log = []
                    st.session_state.log.append({
                        "fecha": pd.Timestamp.now(),
                        "grupo": grupo_seleccionado,
                        "empleado1": emp1_nombre,
                        "empleado2": emp2_nombre,
                        "dia": dia_seleccionado,
                        "servicio1_original": serv1,
                        "servicio2_original": serv2
                    })
                    # Actualizar visualización
                    mostrar_cuadrante(data)
    
    # Mostrar log de cambios
    if "log" in st.session_state and st.session_state.log:
        with st.expander("📜 Historial de intercambios"):
            log_df = pd.DataFrame(st.session_state.log)
            st.dataframe(log_df)
    
    # Guardar cambios
    st.subheader("💾 Guardar cambios")
    if st.button("Guardar y descargar archivo modificado"):
        excel_bytes = guardar_excel(data, uploaded_file)
        st.download_button(
            label="Descargar Excel actualizado",
            data=excel_bytes,
            file_name="cuadrante_mayo_modificado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Archivo listo para descargar. No olvides reemplazar el original.")