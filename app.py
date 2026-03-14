import streamlit as st
import pandas as pd
import io
import time

# Importación de módulos propios
# Aseguramos que lector_datos, motor_logica y escritor_excel estén en la misma carpeta
from lector_datos import cargar_todo
from motor_logica import asignar
from escritor_excel import generar_salida_semanal

# 1. Configuración de la página
st.set_page_config(
    page_title="Generador de Zonnings PDG",
    page_icon="📅",
    layout="wide"
)

def main():
    st.title("📅 Generador de Zonnings PDG")
    st.markdown("""
    Esta aplicación genera cuadrantes semanales asignando empleados a zonas según sus habilidades y disponibilidad.
    
    **Instrucciones:**
    1. Sube el archivo de **Datos** (Empleados y Zonas).
    2. Sube el archivo de **Horario Semanal** (Turnos por día).
    3. Sube la **Plantilla Visual** para estilos.
    4. Pulsa 'Generar Cuadrante'.
    """)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("1. Datos")
        archivo_datos = st.file_uploader("Cargar 'Datos.xlsx'", type=["xlsx"], key="datos")
        
    with col2:
        st.subheader("2. Horario Semanal")
        archivo_horario = st.file_uploader("Cargar 'Horario semana.xlsx'", type=["xlsx"], key="horario")
        
    with col3:
        st.subheader("3. Plantilla Visual")
        archivo_plantilla = st.file_uploader("Cargar 'Plantilla_Visual.xlsx'", type=["xlsx"], key="plantilla")
        
    st.divider()
    
    # 4. El botón de arranque
    if st.button("Generar Cuadrante", type="primary"):
        if archivo_datos is None or archivo_horario is None:
            st.error("⚠️ Por favor, sube ambos archivos para continuar.")
        else:
            procesar_pipeline(archivo_datos, archivo_horario, archivo_plantilla)

def procesar_pipeline(buffer_datos, buffer_horario, buffer_plantilla=None):
    """Ejecuta el pipeline completo usando buffers en memoria."""
    
    with st.spinner('🔄 Procesando datos y generando asignaciones...'):
        try:
            # 1. Cargar datos (Paso los objetos UploadedFile directamente)
            # Aseguramos que los punteros estén al inicio por si se re-usan
            buffer_datos.seek(0)
            buffer_horario.seek(0)
            if buffer_plantilla:
                buffer_plantilla.seek(0)
            
            # Llamamos a cargar_todo pasando los buffers en lugar de rutas
            datos = cargar_todo(ruta_datos=buffer_datos, ruta_horario=buffer_horario)
            
            if datos is None:
                st.error("❌ Error al cargar los datos. Verifica el formato de los archivos Excel.")
                return

            # 2. Procesar lógica
            empleados_por_dia = datos['empleados_por_dia']
            zonas = datos['zonas']
            mapeo_hab = datos['mapeo_habilidades']
            
            cuadrantes_por_dia = {}
            # Usamos un expander para mostrar el log sin ocupar toda la pantalla
            log_container = st.expander("Ver detalles del procesamiento", expanded=True)
            
            for dia, empleados_dia in empleados_por_dia.items():
                # Filtrar empleados activos (Lógica original de main.py)
                try:
                    empleados_activos = empleados_dia[
                        empleados_dia['Horas de trabajo'].astype(float) > 0
                    ]
                except Exception:
                    empleados_activos = empleados_dia
                
                if empleados_activos.empty:
                    log_container.write(f"⚠️ {dia}: Omitido (sin empleados activos)")
                    continue
                
                try:
                    # Llamada al motor lógico
                    cuadrante = asignar(empleados_dia, zonas, mapeo_hab)
                    cuadrantes_por_dia[dia] = cuadrante
                    log_container.write(f"✅ {dia}: {len(cuadrante)} asignaciones generadas")
                except Exception as e:
                    log_container.error(f"❌ {dia}: Error en motor ({e})")
            
            if not cuadrantes_por_dia:
                st.error("❌ No se pudieron generar cuadrantes para ningún día.")
                return

            # 3. Generar Excel en memoria
            excel_buffer = io.BytesIO()
            exito = generar_salida_semanal(
                empleados_por_dia,
                cuadrantes_por_dia,
                ruta_salida=excel_buffer,
                ruta_plantilla=buffer_plantilla if buffer_plantilla else 'Plantilla_Visual.xlsx',
                ruta_horario=buffer_horario
            )
            
            if exito:
                excel_buffer.seek(0)
                st.success("✅ ¡Cuadrante generado exitosamente!")
                
                fecha_hoy = pd.Timestamp.now().strftime("%Y-%m-%d")
                nombre_archivo = f"Zonning_Semanal_Completo_{fecha_hoy}.xlsx"
                
                st.download_button(
                    label="📥 Descargar Excel Resultado",
                    data=excel_buffer,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            else:
                st.error("❌ Fallo en la generación del archivo de salida. Asegúrate de que 'Plantilla_Visual.xlsx' esté en la misma carpeta que app.py.")

        except Exception as e:
            st.error(f"❌ Ocurrió un error inesperado: {e}")
            st.exception(e)

if __name__ == "__main__":
    main()
