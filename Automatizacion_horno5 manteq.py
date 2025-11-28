# -*- coding: utf-8 -*-
"""
Created on Mon Nov 24 09:07:27 2025

@author: NCGNpracpim
"""

import pandas as pd
import numpy as np
import streamlit as st
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font 
from openpyxl.utils.dataframe import dataframe_to_rows
from collections import Counter
import re 
import os 
from typing import Tuple, Union

# --- NOMBRES DE COLUMNAS CLAVE Y CONSTANTES ---
COL_CANT_CALCULADA = 'Cant. base calculada'
COL_PESO_NETO = 'peso neto'
COL_SECUENCIA = 'secuencia recurso'
COL_ATIPICO = 'Atipico_Cant_Calculada'
COL_MANO_OBRA = 'Mano de obra' 
HOJA_MANO_OBRA = 'Mano de obra'
COL_SUMA_VALORES = 'suma valores'
COL_PORCENTAJE_RECHAZO = '%de rechazo'
COL_NRO_PERSONAS = 'Cant_Manual' 
COL_NRO_MAQUINAS = 'Cant_Maquinas' 

# Nombres de hojas a crear
HOJA_PRINCIPAL = 'mantequilla'
HOJA_SECUENCIAS = 'Secuencias'
HOJA_SALIDA = 'mantequilla_procesada'

# Definición de columnas de salida (restauradas del código original)
COLUMNAS_LSMW = [
    'PstoTbjo', 'GrpHRuta', 'CGH', 'Material', 'Ce.', 'Op.', 
    COL_CANT_CALCULADA, 'ValPref', 'ValPref1', 'Mano de obra', 'ValPref3', 
    'suma valores', 'ValPref5'
]
COLUMNAS_CAMPOS_USUARIO = [
    'GrpHRuta', 'CGH', 'Material', 'Ce.', 'Op.', 
    'Indicador', 'clase de control', 
     COL_NRO_PERSONAS, COL_NRO_MAQUINAS 
]
COLUMNAS_RECHAZO = [
    'GrPlf', 'Clave_Busqueda', 'Material', 'Ce.', 'alternativa', 'alternativa', 
    'posición', 'Relevancia', COL_PORCENTAJE_RECHAZO, 
    '% rechazo anterior', 'Diferencia', 'Txt.brv.HRuta'
]


# --- FUNCIONES DE LÓGICA (Mantenidas) ---

def detectar_y_marcar_cantidad_atipica(grupo: pd.DataFrame) -> pd.Series:
    valores_no_nan = grupo[COL_CANT_CALCULADA].dropna()
    
    if valores_no_nan.empty:
        return pd.Series(False, index=grupo.index)

    conteo = Counter(valores_no_nan)
    moda = conteo.most_common(1)[0][0]  
    
    es_diferente_a_moda = grupo[COL_CANT_CALCULADA] != moda
    
    return es_diferente_a_moda

def crear_y_guardar_hoja(wb, df_base: pd.DataFrame, nombre_hoja: str, columnas_destino: list):
    """Crea una nueva hoja y la rellena con las columnas especificadas de df_base."""
    if nombre_hoja in wb.sheetnames:
        del wb[nombre_hoja]
    
    ws = wb.create_sheet(nombre_hoja)
    
    # 1. Crear el nuevo DataFrame con las columnas solicitadas
    df_nuevo = pd.DataFrame()
    for col in columnas_destino:
        if col in df_base.columns:
            df_nuevo[col] = df_base[col]
        else:
            df_nuevo[col] = np.nan
            
    # 2. Escribir el nuevo DataFrame en la hoja
    for row in dataframe_to_rows(df_nuevo, header=True, index=False):
        ws.append(row)
        
    

# --- FUNCIÓN PRINCIPAL ADAPTADA PARA RECIBIR OBJETOS DE ARCHIVO ---

def automatizacion_final_diferencia_reforzada(file_original: io.BytesIO, file_info_externa: io.BytesIO) -> Tuple[bool, Union[str, io.BytesIO]]:
    """
    Ejecuta toda la lógica de procesamiento.
    Recibe objetos de archivo (bytes/buffers) en lugar de rutas.
    Devuelve un buffer de bytes para la descarga.
    """
    
    # --- 2. DEFINICIÓN DE ÍNDICES Y NOMBRES CLAVE ---
    IDX_MATERIAL = 2; IDX_GRPLF = 4; IDX_PSTTBJO = 18; IDX_CANTIDAD_BASE = 6
    IDX_MATERIAL_PN = 0; IDX_VALPREF2 = 14; IDX_RECHAZO_EXTERNA = 28  
    
    COL_CLAVE = 'Clave_Busqueda'; COL_DIFERENCIA = 'diferencia';  
    NOMBRE_COL_CANTIDAD_BASE = 'Cantidad base'
    NOMBRE_COL_CLAVE_EXTERNA = 'MaterialHorno'
    NOMBRE_COL_CANT_EXTERNA = 'CantidadBaseXHora'

    FINAL_COL_ORDER = [
        'GrpHRuta', 'CGH', 'Material', COL_CLAVE, 'Ce.', 'GrPlf', 'Op.',
        COL_PORCENTAJE_RECHAZO, NOMBRE_COL_CANTIDAD_BASE, COL_CANT_CALCULADA,  
        COL_DIFERENCIA, COL_PESO_NETO, COL_SECUENCIA, 'ValPref', 'ValPref1', 
        'ValPref2', COL_MANO_OBRA, 'ValPref3', 'ValPref4', COL_SUMA_VALORES,  
        'ValPref5', 
        'Campo de usuario cantidad MANUAL',  
        COL_NRO_PERSONAS,  
        'Campo de usuario cantidad MAQUINAS',  
        COL_NRO_MAQUINAS,  
        'Texto breve operación', 'Ctrl', 'VerF', 'PstoTbjo', 'Cl.', 'Gr.fam.pre',  
        'Texto breve de material', 'Txt.brv.HRuta', 'Bloq.vers.fabric.', 'Campo usuario unidad',  
        'Campo usuario unidad.1', 'Cantidad', 'Contador', 'InBo', 'InBo.1', 'InBo.2',  
        'Unnamed: 31', 'I'
    ]

    try:
        st.write("---")
        st.subheader("Preparando datos... 📊")
        
        # --- 3. LECTURA DE DATOS DESDE LOS OBJETOS DE ARCHIVO ---
        
        # 3.1 Cargar el archivo original y determinar el nombre de la columna en IDX_CANTIDAD_BASE
        cols_original = pd.read_excel(file_original, sheet_name=HOJA_PRINCIPAL, nrows=0).columns.tolist()
        file_original.seek(0) # Resetear el puntero después de leer encabezados
        
        nombre_cant_base_leida = cols_original[IDX_CANTIDAD_BASE]
        nombre_material = cols_original[IDX_MATERIAL]
        nombre_psttbjo = cols_original[IDX_PSTTBJO]
            
        # Carga del DataFrame Principal, de Peso Neto, Secuencias y Mano de Obra
        df_original = pd.read_excel(file_original, sheet_name=HOJA_PRINCIPAL, dtype={nombre_cant_base_leida: str})
        file_original.seek(0) 

        df_peso_neto = pd.read_excel(file_original, sheet_name='Peso neto')
        file_original.seek(0)
        
        df_secuencias = pd.read_excel(file_original, sheet_name=HOJA_SECUENCIAS)
        file_original.seek(0)
        
        # LECTURA CORREGIDA: Leemos todas las columnas (sin usecols) para evitar el error de out-of-bounds.
        # df_mano_obra debe tener suficientes columnas para los nuevos índices A, C, D, E
        df_mano_obra = pd.read_excel(file_original, sheet_name=HOJA_MANO_OBRA, header=None)
        
        # 3.2 Lectura de archivo externo
        cols_externo = pd.read_excel(file_info_externa, sheet_name='Especif y Rutas', nrows=0).columns.tolist()
        file_info_externa.seek(0)
        
        NOMBRE_COL_RECHAZO_EXTERNA = cols_externo[IDX_RECHAZO_EXTERNA] if IDX_RECHAZO_EXTERNA < len(cols_externo) else 'Columna AC'
        cols_a_leer_externo = [NOMBRE_COL_CLAVE_EXTERNA, NOMBRE_COL_CANT_EXTERNA, NOMBRE_COL_RECHAZO_EXTERNA]
        df_externo = pd.read_excel(file_info_externa, sheet_name='Especif y Rutas', header=0, usecols=cols_a_leer_externo)
        
        if nombre_cant_base_leida != NOMBRE_COL_CANTIDAD_BASE:
            df_original = df_original.rename(columns={nombre_cant_base_leida: NOMBRE_COL_CANTIDAD_BASE})

        # --- 4. CÁLCULOS Y BÚSQUEDAS (Clave, Cantidad Calculada, Rechazo, Peso Neto, Secuencia) ---
        
        def limpiar_col(df: pd.DataFrame, idx: int) -> pd.Series:
            return df[cols_original[idx]].astype(str).str.strip().str.replace(r'\W+', '', regex=True)

        df_original[COL_CLAVE] = (
            limpiar_col(df_original, IDX_MATERIAL) + 
            limpiar_col(df_original, IDX_GRPLF) + 
            limpiar_col(df_original, IDX_PSTTBJO)
        )
        
        df_externo[NOMBRE_COL_CLAVE_EXTERNA] = df_externo[NOMBRE_COL_CLAVE_EXTERNA].astype(str).str.strip().str.replace(r'\W+', '', regex=True)
        
        mapa_cantidad = df_externo.drop_duplicates(subset=[NOMBRE_COL_CLAVE_EXTERNA], keep='first').set_index(NOMBRE_COL_CLAVE_EXTERNA)[NOMBRE_COL_CANT_EXTERNA]
        df_original[COL_CANT_CALCULADA] = df_original[COL_CLAVE].map(mapa_cantidad)
        
        mapa_rechazo = df_externo.drop_duplicates(subset=[NOMBRE_COL_CLAVE_EXTERNA], keep='first').set_index(NOMBRE_COL_CLAVE_EXTERNA)[NOMBRE_COL_RECHAZO_EXTERNA]
        df_original[COL_PORCENTAJE_RECHAZO] = df_original[COL_CLAVE].map(mapa_rechazo)
        
        cols_peso_neto = df_peso_neto.columns.tolist()
        nombre_material_pn = cols_peso_neto[IDX_MATERIAL_PN]
        nombre_peso_neto_valor = cols_peso_neto[2]
        mapa_peso_neto = df_peso_neto.drop_duplicates(subset=[nombre_material_pn], keep='first').set_index(nombre_material_pn)[nombre_peso_neto_valor]
        df_original[COL_PESO_NETO] = df_original[nombre_material].map(mapa_peso_neto)

        def obtener_secuencia(puesto_trabajo):
            psttbjo_str = str(puesto_trabajo).strip()
            psttbjo_sec_1 = set(df_secuencias.iloc[:, 0].dropna().astype(str).str.strip())
            psttbjo_sec_2 = set(df_secuencias.iloc[:, 3].dropna().astype(str).str.strip())
            if psttbjo_str in psttbjo_sec_1: return 1
            elif psttbjo_str in psttbjo_sec_2: return 2
            return np.nan

        df_original[COL_SECUENCIA] = df_original[nombre_psttbjo].apply(obtener_secuencia)

        # --- 5. CÁLCULO DE MANO DE OBRA Y MÁQUINAS (CONDICIÓN OP. TERMINADA EN '1') ---
        
        # NUEVOS ÍNDICES BASADOS EN LA TABLA DEL USUARIO (A=0, C=2, D=3, E=4)
        
        # 1. Mano de Obra (Tiempo: Columna A -> Columna C)
        COL_PSTTBJO_MO_TIEMPO = 0  # Columna A
        COL_CANTIDAD_MO_TIEMPO = 2 # Columna C
        
        # 2. Cant. Máquinas (Columna A -> Columna D)
        COL_PSTTBJO_MAQUINAS = 0   # Columna A
        COL_CANTIDAD_MAQUINAS = 3  # Columna D
        
        # 3. Cant. Personas (Columna A -> Columna E)
        COL_PSTTBJO_PERSONAS = 0   # Columna A
        COL_CANTIDAD_PERSONAS = 4  # Columna E
        
        # Limpiar y convertir tipos de datos para los mapas
        df_mano_obra[COL_PSTTBJO_MO_TIEMPO] = df_mano_obra[COL_PSTTBJO_MO_TIEMPO].astype(str).str.strip()
        df_mano_obra[COL_CANTIDAD_MO_TIEMPO] = pd.to_numeric(df_mano_obra[COL_CANTIDAD_MO_TIEMPO], errors='coerce')
        
        df_mano_obra[COL_PSTTBJO_MAQUINAS] = df_mano_obra[COL_PSTTBJO_MAQUINAS].astype(str).str.strip()
        df_mano_obra[COL_CANTIDAD_MAQUINAS] = pd.to_numeric(df_mano_obra[COL_CANTIDAD_MAQUINAS], errors='coerce')

        df_mano_obra[COL_PSTTBJO_PERSONAS] = df_mano_obra[COL_PSTTBJO_PERSONAS].astype(str).str.strip()
        df_mano_obra[COL_CANTIDAD_PERSONAS] = pd.to_numeric(df_mano_obra[COL_CANTIDAD_PERSONAS], errors='coerce')


        # 5.1. Cálculo de TIEMPO de Mano de Obra (Personas * 60)
        df_mano_obra['Calculo_MO'] = df_mano_obra[COL_CANTIDAD_MO_TIEMPO] * 60
        mapa_mano_obra_tiempo = df_mano_obra.drop_duplicates(subset=[COL_PSTTBJO_MO_TIEMPO], keep='first').set_index(COL_PSTTBJO_MO_TIEMPO)['Calculo_MO']
        
        COL_OP = 'Op.' 
        df_original[COL_MANO_OBRA] = np.nan
        op_col = df_original[COL_OP].astype(str).str.strip()
        indices_terminan_en_1 = op_col.str.endswith('1')
        psttbjo_filtrado = df_original.loc[indices_terminan_en_1, nombre_psttbjo].astype(str).str.strip()
        df_original.loc[indices_terminan_en_1, COL_MANO_OBRA] = psttbjo_filtrado.map(mapa_mano_obra_tiempo)


        # 5.2. Búsqueda del Número de PERSONAS (Columna E)
        mapa_personas = df_mano_obra.drop_duplicates(subset=[COL_PSTTBJO_PERSONAS], keep='first').set_index(COL_PSTTBJO_PERSONAS)[COL_CANTIDAD_PERSONAS]
        df_original[COL_NRO_PERSONAS] = np.nan 
        df_original.loc[indices_terminan_en_1, COL_NRO_PERSONAS] = psttbjo_filtrado.map(mapa_personas)


        # 5.3. Búsqueda del Número de MÁQUINAS (Columna D)
        mapa_maquinas = df_mano_obra.drop_duplicates(subset=[COL_PSTTBJO_MAQUINAS], keep='first').set_index(COL_PSTTBJO_MAQUINAS)[COL_CANTIDAD_MAQUINAS]
        df_original[COL_NRO_MAQUINAS] = np.nan
        df_original.loc[indices_terminan_en_1, COL_NRO_MAQUINAS] = psttbjo_filtrado.map(mapa_maquinas)


        # Suma de Valores
        COLUMNAS_A_SUMAR = ['ValPref', 'ValPref1', COL_MANO_OBRA, 'ValPref3']
        df_temp_sum = df_original[COLUMNAS_A_SUMAR].copy()
        for col in COLUMNAS_A_SUMAR:
             df_temp_sum[col] = pd.to_numeric(df_temp_sum[col], errors='coerce')
        df_original[COL_SUMA_VALORES] = df_temp_sum.sum(axis=1, skipna=True)
        
        def formato_excel_regional_suma(x):
            if pd.isna(x) or x is None or x == 0.0: return np.nan
            return f"{x:.2f}".replace('.', ',')
        df_original[COL_SUMA_VALORES] = df_original[COL_SUMA_VALORES].apply(formato_excel_regional_suma)


        # --- 6. CÁLCULO DE DIFERENCIA Y ATÍPICOS ---
        H = pd.to_numeric(df_original[NOMBRE_COL_CANTIDAD_BASE].astype(str).str.replace(',', '.', regex=False).str.strip(), errors='coerce')
        I = pd.to_numeric(df_original[COL_CANT_CALCULADA], errors='coerce')
        diferencia_calculada = H.fillna(0) - I.fillna(0)
        
        def formato_excel_regional(x):
            return f"{x:.2f}".replace('.', ',') if pd.notna(x) else np.nan
        df_original[COL_DIFERENCIA] = diferencia_calculada.apply(formato_excel_regional)
        
        # Atípicos
        df_original[COL_CANT_CALCULADA] = pd.to_numeric(df_original[COL_CANT_CALCULADA], errors='coerce')
        df_original[COL_ATIPICO] = df_original.groupby([COL_PESO_NETO, COL_SECUENCIA], dropna=True).apply(
            detectar_y_marcar_cantidad_atipica
        ).reset_index(level=[0, 1], drop=True) 
        df_original[COL_ATIPICO] = df_original[COL_ATIPICO].fillna(False)


        # --- 8. RECONSTRUCCIÓN FINAL Y GUARDADO CON FORMATO EN MEMORIA ---
        
        
        df_original_final = df_original.reindex(columns=[c for c in FINAL_COL_ORDER if c in df_original.columns])
        
        # 1. Cargar el libro de trabajo desde el buffer (Necesario para mantener las hojas originales)
        file_original.seek(0)
        wb = load_workbook(file_original)
        
        # --- DEFINICIÓN DE ESTILOS DE SOMBREADO Y FUENTE ---
        fill_anomalia = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid') 
        fill_encabezado = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid') 
        font_negrita = Font(bold=True) 
        
        # 2. Eliminar/Crear la hoja principal 'mantequilla_procesada'
        if HOJA_SALIDA in wb.sheetnames: del wb[HOJA_SALIDA]
        ws = wb.create_sheet(HOJA_SALIDA)
        
        # 3. Escribir el DataFrame principal
        for r_idx, row in enumerate(dataframe_to_rows(df_original_final, header=True, index=False)):
            ws.append(row)
        
        # --- 4. APLICACIÓN DE FORMATOS (RESTAURADO) ---
        COLUMNAS_ENCABEZADO_FORMATO = [COL_CANT_CALCULADA, COL_MANO_OBRA, COL_SUMA_VALORES]
        
        indices_encabezado = []
        for col_name in COLUMNAS_ENCABEZADO_FORMATO:
            try:
                idx = df_original_final.columns.get_loc(col_name) + 1
                indices_encabezado.append(idx)
            except KeyError:
                pass 

        for col_idx in indices_encabezado:
            header_cell = ws.cell(row=1, column=col_idx)
            header_cell.fill = fill_encabezado
            header_cell.font = font_negrita 
        
        # Aplicar el sombreado a los datos de 'Cant. base calculada' (naranja para atípicos)
        try:
            col_cant_calculada_idx = df_original_final.columns.get_loc(COL_CANT_CALCULADA) + 1 
        except KeyError:
            col_cant_calculada_idx = 9 

        for r in range(2, len(df_original) + 2):
            es_atipico = df_original.iloc[r-2][COL_ATIPICO]
            if es_atipico:
                cell_to_color = ws.cell(row=r, column=col_cant_calculada_idx)
                cell_to_color.fill = fill_anomalia
        
        # --- CREACIÓN DE HOJAS ADICIONALES (Mantiene las originales y agrega las nuevas) ---
        
        crear_y_guardar_hoja(wb, df_original, "lsmw", COLUMNAS_LSMW)
        crear_y_guardar_hoja(wb, df_original, "campos de usuario", COLUMNAS_CAMPOS_USUARIO)
        crear_y_guardar_hoja(wb, df_original, "% de rechazo", COLUMNAS_RECHAZO)

        # 5. Guardar el libro de trabajo modificado en un buffer de Bytes
        output_buffer = io.BytesIO()
        wb.save(output_buffer)
        output_buffer.seek(0)
        
        return True, output_buffer

    except KeyError as ke:
        return False, f"❌ ERROR CRÍTICO DE ENCABEZADO: El script no encontró la columna {ke}. Verifique las hojas y encabezados del archivo original o externo."
    except IndexError as ie:
        # Esto se ajusta ya que solo necesitamos hasta la columna E (índice 4)
        return False, f"❌ ERROR CRÍTICO DE ÍNDICE: Asegúrese de que la hoja 'Mano de obra' tenga al menos 5 columnas (hasta la Columna E). Mensaje: {ie}"
    except ValueError as ve:
        if 'sheetname' in str(ve) or 'Worksheet' in str(ve):
            return False, f"❌ Error de Lectura de Hoja: Una de las hojas clave ({HOJA_PRINCIPAL}, Peso neto, {HOJA_SECUENCIAS}, {HOJA_MANO_OBRA}, 'Especif y Rutas') no se encontró en los archivos cargados. Mensaje: {ve}"
        return False, f"❌ Ocurrió un error inesperado de valor. Mensaje: {ve}"
    except Exception as e:
        return False, f"❌ Ocurrió un error inesperado. Mensaje: {e}"


# --- INTERFAZ DE STREAMLIT ---

def main():
    st.set_page_config(
        page_title="Automatización Horno 5",
        layout="centered",
        initial_sidebar_state="auto"
    )

    st.title("⚙️ Automatización Verificación de datos- HORNOS")
    st.markdown("Cargue los dos archivos requeridos para generar el reporte procesado")

    col1, col2 = st.columns(2)

    with col1:
        file_original = st.file_uploader(
            "Carga la base de datos original",
            type=['xlsx'], 
            help="El archivo que contiene la base de datos."
        )

    with col2:
        file_externa = st.file_uploader(
            "Carga el archivo externo de toma de información.",
            type=['xlsb', 'xlsx'], 
            help="El archivo que contiene la hoja 'Especif y Rutas' para la Cantidad Base y el % de rechazo."
        )

    st.markdown("---")
    
    # 💡 Botón de ejecución y manejo del proceso
    if st.button("▶️ PROCESAR HORNO 5", type="primary", use_container_width=True):
        if file_original is None or file_externa is None:
            st.error("Por favor, cargue ambos archivos antes de procesar.")
        else:
            # En Streamlit, la magia es usar buffers de bytes (io.BytesIO)
            # para simular el archivo en memoria.
            file_buffer_original = io.BytesIO(file_original.getvalue())
            file_buffer_externa = io.BytesIO(file_externa.getvalue())

            # Usar st.spinner para mostrar progreso
            with st.spinner('Procesando datos y generando reporte...'):
                success, resultado = automatizacion_final_diferencia_reforzada(
                    file_buffer_original, 
                    file_buffer_externa
                )

            st.markdown("---")

            if success:
                st.success("✅ Proceso completado exitosamente.")
                
                # Botón de Descarga
                st.download_button(
                    label="⬇️ Descargar Archivo Procesado (mantequilla.xlsx)",
                    data=resultado, # El buffer de bytes es el archivo final
                    file_name=file_original.name.split('.')[0] + "_procesado.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True
                )
                st.info("El archivo descargado contiene todas las hojas originales más las 4 hojas de reporte.")
            else:
                st.error("❌ Error en el Proceso")
                st.warning(resultado)
                st.write("Verifique el formato de las hojas y los nombres de las columnas en sus archivos.")

if __name__ == "__main__":

    main()









