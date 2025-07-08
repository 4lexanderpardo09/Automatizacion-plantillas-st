# app.py

import streamlit as st
from logic import cargar_y_filtrar_datos, generar_reporte_region, REGIONES

# --- CONFIGURACIÓN E INTERFAZ DE STREAMLIT ---
st.set_page_config(page_title="Generador de Reportes", layout="centered")
st.title('🤖 Generador Automático de Reportes de Servicio')
st.markdown("---")

# --- LÓGICA DE LA APLICACIÓN ---
st.header("Paso 1: Carga tu archivo de consulta")
archivo_cargado = st.file_uploader(
    "Selecciona el archivo 'CRTMPCONSULTA.xlsx' o similar",
    type=['xlsx']
)

# Botón para procesar el archivo. Solo se ejecuta una vez.
if st.button('🚀 Procesar Archivo', type="primary"):
    if archivo_cargado is not None:
        try:
            with st.spinner('Leyendo y procesando el archivo base...'):
                # Guardamos el dataframe procesado en el session_state
                st.session_state['df_pendientes'] = cargar_y_filtrar_datos(archivo_cargado)
            st.success("¡Archivo procesado! Ya puedes descargar los reportes.")
        except Exception as e:
            st.error(f"Ocurrió un error al procesar el archivo: {e}")
            # Limpiamos por si hubo un error
            if 'df_pendientes' in st.session_state:
                del st.session_state['df_pendientes']
    else:
        st.warning("Por favor, carga un archivo primero.")

# Esta sección se mostrará siempre que los datos ya hayan sido procesados y guardados.
if 'df_pendientes' in st.session_state:
    st.markdown("---")
    st.header("Paso 2: Descarga tus reportes generados")
    
    df_pendientes = st.session_state['df_pendientes']
    reportes_generados = 0

    for region in REGIONES:
        df_region = df_pendientes[df_pendientes['REGION'] == region]
        
        if not df_region.empty:
            archivo_bytes = generar_reporte_region(df_region, region)
            st.download_button(
                label=f"📥 Descargar Reporte {region} ({len(df_region)} regs.)",
                data=archivo_bytes,
                file_name=f'reporte_{region.lower()}_actualizado_{REGIONES[region][0]}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                # Clave única para cada botón
                key=f'btn_{region}' 
            )
            reportes_generados += 1

    if reportes_generados == 0:
        st.info("Aunque el archivo se procesó, no se encontraron registros pendientes para ninguna región.")