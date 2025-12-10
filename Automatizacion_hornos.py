# -*- coding: utf-8 -*-
"""
Creado el Lunes 24 de Noviembre de 2025

@author: NCGNpracpim

MODIFICACIONES IMPLEMENTADAS:
1. COL_SECUENCIA, Mano de Obra (Personas/Máquinas) se calculan en Python (valores fijos).
2. Fórmulas de Excel (BUSCARV, Suma, Resta) se usan para Cant. Calculada, Diferencia, Peso Neto y Rechazo.
3. CORRECCIÓN DE ERROR (Log de Recuperación): Las fórmulas SOLO se escriben en la Hoja Principal de Salida.
4. CORRECCIÓN DE ERROR (NameError): Corregida la variable 'IDX_RECHAZA_EXTERNA' a 'IDX_RECHAZO_EXTERNA'.
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
from typing import Tuple, Union, Dict, Any

# --- NOMBRES DE COLUMNAS CLAVE Y CONSTANTES COMUNES ---
COL_CANT_CALCULADA = 'Cant. base calculada'
COL_PESO_NETO = 'peso neto'
COL_SECUENCIA = 'secuencia recurso'
COL_ATIPICO = 'Atipico_Cant_Calculada'
COL_MANO_OBRA = 'Mano de obra'
HOJA_MANO_OBRA = 'Mano de obra' # Esta hoja es común
COL_SUMA_VALORES = 'suma valores'
COL_PORCENTAJE_RECHAZO = '%de rechazo'
COL_NRO_PERSONAS = 'Cant_Manual'
COL_NRO_MAQUINAS = 'Cant_Maquinas'
COL_CLAVE = 'Clave_Busqueda'
COL_DIFERENCIA = 'diferencia'
NOMBRE_COL_CANTIDAD_BASE = 'Cantidad base'
NOMBRE_COL_CLAVE_EXTERNA = 'MaterialHorno'
NOMBRE_COL_CANT_EXTERNA = 'CantidadBaseXHora'
COL_LINEA = 'Linea'
COL_PSTTBJO_CONCATENADO = 'PstoTbjo_Concat' 

# Nombres de hojas a crear (Comunes)
HOJA_SECUENCIAS = 'Secuencias' 
HOJA_LSMW = 'lsmw'
HOJA_CAMPOS_USUARIO = 'campos de usuario'
HOJA_PORCENTAJE_RECHAZO = '% de rechazo'
HOJA_EXTERNA = 'Especif y Rutas' 

# Columnas a resaltar en todas las hojas (solicitado por el usuario)
COLUMNAS_A_RESALTAR = [
    COL_MANO_OBRA,
    COL_SUMA_VALORES,
    COL_NRO_PERSONAS,
    COL_NRO_MAQUINAS
]

# Definición de columnas de salida (Comunes)
COLUMNAS_LSMW = [
    'PstoTbjo', 'GrpHRuta', 'CGH', 'Material', COL_CLAVE, 'Ce.', 'Op.',
    COL_CANT_CALCULADA, 'ValPref', 'ValPref1', COL_MANO_OBRA, 'ValPref3',
    COL_SUMA_VALORES, 'ValPref5'
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

# --- CONSTANTES ESPECÍFICAS DE CADA HORNO ---
HORNOS_CONFIG = {
    'HORNO 5': {
        'HOJA_PRINCIPAL': 'HORNO 5',
        'HOJA_SALIDA': 'HORNO5_procesado',
    },
    'HORNO 12': {
        'HOJA_PRINCIPAL': 'HORNO 12',
        'HOJA_SALIDA': 'HORNO12_procesado',
    },
    'HORNO 1': {
        'HOJA_PRINCIPAL': 'HORNO 1',
        'HOJA_SALIDA': 'HORNO1_procesado',
    },
    'HORNO 3': {
        'HOJA_PRINCIPAL': 'HORNO 3',
        'HOJA_SALIDA': 'HORNO3_procesado',
    },
}

# Índices para el archivo original
IDX_MATERIAL = 2 
IDX_GRPLF = 4 
IDX_CANTIDAD_BASE_LEIDA = 6 
IDX_PSTTBJO = 18 
IDX_MATERIAL_PN = 0
IDX_RECHAZO_EXTERNA = 28 

# --- CONSTANTES DE REFERENCIA EXCEL EN LA HOJA DE SALIDA ---
COL_MATERIAL_OUTPUT_EXCEL = 'C'
COL_CLAVE_OUTPUT_EXCEL = 'D'
COL_OP_OUTPUT_EXCEL = 'G'
COL_CANT_BASE_OUTPUT_EXCEL = 'I' 
COL_CANT_CALC_OUTPUT_EXCEL = 'J' 
COL_DIFERENCIA_OUTPUT_EXCEL = 'K' 
COL_PESO_NETO_OUTPUT_EXCEL = 'L' 
COL_MANO_OBRA_OUTPUT_EXCEL = 'R' 
COL_NRO_PERSONAS_OUTPUT_EXCEL = 'V'
COL_NRO_MAQUINAS_OUTPUT_EXCEL = 'X'
COL_PSTTBJO_OUTPUT_EXCEL = 'AB' 


# --- CONSTANTES DE REFERENCIA EXCEL EN LAS HOJAS DE BÚSQUEDA ---
RANGO_EXTERNO_BUSCARV = '$A:$AC'
COL_CANT_EXTERNA_INDEX = 2 
COL_RECHAZO_EXTERNA_INDEX = 29

HOJA_PESO_NETO = 'Peso neto'
RANGO_PN_BUSCARV = '$A:$C'
COL_PESO_NETO_INDEX = 3

HOJA_MANO_OBRA = 'Mano de obra'
RANGO_MO_BUSCARV = '$A:$E'
COL_TIEMPO_MO_INDEX = 3
COL_CANTIDAD_PERSONAS_MO_INDEX = 5
COL_CANTIDAD_MAQUINAS_MO_INDEX = 4


# --- FUNCIONES DE LÓGICA ---

def detectar_y_marcar_cantidad_atipica(grupo: pd.DataFrame) -> pd.Series:
    """Identifica valores atípicos (diferentes de la moda) en Cant. base calculada dentro de un grupo."""
    valores_no_nan = grupo[COL_CANT_CALCULADA].dropna()
    if valores_no_nan.empty:
        return pd.Series(False, index=grupo.index)

    conteo = Counter(valores_no_nan)
    moda = conteo.most_common(1)[0][0]

    es_diferente_a_moda = grupo[COL_CANT_CALCULADA] != moda

    return es_diferente_a_moda

def crear_y_guardar_hoja_solo_valores(wb, df_base: pd.DataFrame, nombre_hoja: str, columnas_destino: list, fill_encabezado: PatternFill, font_negrita: Font):
    """
    Crea una nueva hoja, reemplazando cualquier fórmula de Excel con NaN antes de la escritura
    para evitar errores de sintaxis en las hojas secundarias.
    """
    if nombre_hoja in wb.sheetnames:
        del wb[nombre_hoja]

    ws = wb.create_sheet(nombre_hoja)
    
    # 1. Crear el nuevo DataFrame con las columnas solicitadas y reemplazar fórmulas por NaN
    df_nuevo = pd.DataFrame()
    for col in columnas_destino:
        data = df_base[col] if col in df_base.columns else np.nan
        
        if isinstance(data, pd.Series):
            # Reemplazar cualquier valor que comience con '=' (fórmula) por NaN
            data = data.apply(lambda x: np.nan if isinstance(x, str) and str(x).startswith('=') else x)
            
        df_nuevo[col] = data

    # 2. Escribir el nuevo DataFrame en la hoja
    for row in dataframe_to_rows(df_nuevo, header=True, index=False):
        ws.append(row)

    # 3. Aplicar Formato a Encabezados Específicos
    indices_a_formatear = [
        df_nuevo.columns.get_loc(col) + 1  
        for col in COLUMNAS_A_RESALTAR  
        if col in df_nuevo.columns
    ]

    for col_idx in indices_a_formatear:
        header_cell = ws.cell(row=1, column=col_idx)
        header_cell.fill = fill_encabezado
        header_cell.font = font_negrita

def obtener_secuencia(puesto_trabajo: str, df_secuencias: pd.DataFrame) -> Union[int, float]:
    """Busca la secuencia del puesto de trabajo (o PstoTbjo_Concat) en la hoja 'Secuencias'."""
    psttbjo_str = str(puesto_trabajo).strip()

    for col_idx in range(df_secuencias.shape[1]):
        col_data = df_secuencias.iloc[:, col_idx].dropna().astype(str).str.strip()
        psttbjo_sec = set(col_data)

        if psttbjo_str in psttbjo_sec:
            return col_idx + 1

    return np.nan

def cargar_y_limpiar_datos(file_original: io.BytesIO, file_info_externa: io.BytesIO, nombre_horno: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, str]]:
    """Carga todos los DataFrames necesarios desde los buffers de archivo."""
    
    config = HORNOS_CONFIG[nombre_horno]
    hoja_principal = config['HOJA_PRINCIPAL']

    # --- 1. Lectura de Archivo Original ---
    cols_original = pd.read_excel(file_original, sheet_name=hoja_principal, nrows=0).columns.tolist()
    file_original.seek(0)
    
    cols_pn = pd.read_excel(file_original, sheet_name='Peso neto', nrows=0).columns.tolist()
    file_original.seek(0)
    
    col_names = {
        'cant_base_leida': cols_original[IDX_CANTIDAD_BASE_LEIDA] if IDX_CANTIDAD_BASE_LEIDA < len(cols_original) else NOMBRE_COL_CANTIDAD_BASE,
        'material': cols_original[IDX_MATERIAL] if IDX_MATERIAL < len(cols_original) else 'Material',
        'psttbjo': cols_original[IDX_PSTTBJO] if IDX_PSTTBJO < len(cols_original) else 'PstoTbjo',
        'material_pn': cols_pn[IDX_MATERIAL_PN] if IDX_MATERIAL_PN < len(cols_pn) else 'Material',
        'peso_neto_valor': cols_pn[2] if 2 < len(cols_pn) else 'Peso neto',
        'cols_original': cols_original,
        'hoja_principal': hoja_principal
    }

    df_original = pd.read_excel(
        file_original, 
        sheet_name=hoja_principal, 
        dtype={col_names['cant_base_leida']: str},
    )
    file_original.seek(0)
    
    # Renombrar columnas clave si es necesario
    if col_names['cant_base_leida'] != NOMBRE_COL_CANTIDAD_BASE:
        df_original = df_original.rename(columns={col_names['cant_base_leida']: NOMBRE_COL_CANTIDAD_BASE})
        col_names['cant_base_leida'] = NOMBRE_COL_CANTIDAD_BASE
    
    if col_names['material'] != 'Material':
        df_original.rename(columns={col_names['material']: 'Material'}, inplace=True)
        col_names['material'] = 'Material'
    
    if col_names['psttbjo'] != 'PstoTbjo':
        df_original.rename(columns={col_names['psttbjo']: 'PstoTbjo'}, inplace=True)
        col_names['psttbjo'] = 'PstoTbjo'

    df_peso_neto = pd.read_excel(file_original, sheet_name='Peso neto')
    file_original.seek(0)

    df_secuencias = pd.read_excel(file_original, sheet_name=HOJA_SECUENCIAS)
    file_original.seek(0)
    
    columnas_mano_obra = [0, 1, 2, 3, 4] 
    df_mano_obra = pd.read_excel(
        file_original, 
        sheet_name=HOJA_MANO_OBRA, 
        header=None, 
        usecols=range(len(columnas_mano_obra)), 
        names=columnas_mano_obra 
    )
    file_original.seek(0)

    # --- 2. Lectura de Archivo Externo ---
    cols_externo = pd.read_excel(file_info_externa, sheet_name=HOJA_EXTERNA, nrows=0).columns.tolist()
    file_info_externa.seek(0)

    # <<<<<<<<<<<<<<<<< CORRECCIÓN DEL ERROR DE NOMBRE AQUÍ >>>>>>>>>>>>>>>>>>
    # Corregido: IDX_RECHAZA_EXTERNA cambiado a IDX_RECHAZO_EXTERNA
    nombre_col_rechazo_externa = cols_externo[IDX_RECHAZO_EXTERNA] if IDX_RECHAZO_EXTERNA < len(cols_externo) else 'Columna AC'
    
    cols_a_leer_externo = [NOMBRE_COL_CLAVE_EXTERNA, NOMBRE_COL_CANT_EXTERNA, nombre_col_rechazo_externa]
    df_externo = pd.read_excel(file_info_externa, sheet_name=HOJA_EXTERNA, header=0, usecols=cols_a_leer_externo)
    file_info_externa.seek(0)

    col_names['nombre_col_rechazo_externa'] = nombre_col_rechazo_externa

    return df_original, df_externo, df_peso_neto, df_secuencias, df_mano_obra, col_names

# --- FUNCIÓN PRINCIPAL DE PROCESAMIENTO ---

def automatizacion_final_diferencia_reforzada(file_original: io.BytesIO, file_info_externa: io.BytesIO, nombre_horno: str) -> Tuple[bool, Union[str, io.BytesIO]]:
    """
    Ejecuta toda la lógica de procesamiento, insertando fórmulas de Excel.
    """
    
    config = HORNOS_CONFIG[nombre_horno]
    HOJA_SALIDA = config['HOJA_SALIDA']
    
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
        st.subheader(f"Preparando datos para **{nombre_horno}**... 📊")

        # 1. Carga y limpieza de datos
        df_original, df_externo, df_peso_neto, df_secuencias, df_mano_obra, col_names = cargar_y_limpiar_datos(file_original, file_info_externa, nombre_horno)
        
        material_col_name = 'Material'
        grplf_col_name = col_names['cols_original'][IDX_GRPLF]
        psttbjo_col_name = 'PstoTbjo' 

        # 2. Creación de la Clave de Búsqueda
        def limpiar_col(df: pd.DataFrame, col_name: str) -> pd.Series:
            if col_name not in df.columns:
                raise KeyError(f"Columna '{col_name}' no encontrada en la hoja '{col_names['hoja_principal']}'.")
            return df[col_name].astype(str).str.strip().str.replace(r'\W+', '', regex=True)

        df_original[COL_CLAVE] = (
            limpiar_col(df_original, material_col_name) +
            limpiar_col(df_original, grplf_col_name) +
            limpiar_col(df_original, psttbjo_col_name)
        )

        df_externo[NOMBRE_COL_CLAVE_EXTERNA] = df_externo[NOMBRE_COL_CLAVE_EXTERNA].astype(str).str.strip().str.replace(r'\W+', '', regex=True)

        # --- LÓGICA DE CONCATENACIÓN Y DETERMINACIÓN DE COLUMNA DE BÚSQUEDA ---
        columna_para_secuencia = psttbjo_col_name 
        linea_existe = COL_LINEA in df_original.columns 

        if linea_existe and df_original[COL_LINEA].astype(str).str.lower().str.contains(r'[a-z0-9]').any():
            st.info(f"⚠️ **Detectada la columna '{COL_LINEA}'**. Aplicando lógica de concatenación **PstoTbjo + Línea** para búsqueda de secuencia.")
            
            linea_limpia = df_original[COL_LINEA].astype(str).str.strip()
            psttbjo_limpio = df_original[psttbjo_col_name].astype(str).str.strip()
            
            df_original[COL_PSTTBJO_CONCATENADO] = np.where(
                linea_limpia.str.lower().isin(['nan', 'none', '']), 
                psttbjo_limpio,
                psttbjo_limpio + linea_limpia
            )
            columna_para_secuencia = COL_PSTTBJO_CONCATENADO
        else:
            st.info("✅ **Columna 'Linea' no detectada o vacía**. Usando solo el Puesto de Trabajo para la búsqueda de secuencia.")
        # ------------------------------------------------------------------

        # Obtener los índices de fila para la construcción de fórmulas de Excel
        indices_fila_excel = range(2, len(df_original) + 2)
        
        # --- 3. Mapeo Temporal de Valores para Cálculo de Atípicos y Secuencia ---
        
        # Función auxiliar para mapeo de Cantidad Máxima (Necesaria solo para Atípicos)
        def mapear_con_maxima_cantidad_temp(df_origen: pd.DataFrame, df_externo: pd.DataFrame, col_clave_origen: str, col_clave_externa: str, col_cantidad_externa: str, col_destino: str):
            df_externo[col_cantidad_externa] = pd.to_numeric(df_externo[col_cantidad_externa], errors='coerce')
            df_mapa = (
                df_externo.sort_values(by=col_cantidad_externa, ascending=False)
                .drop_duplicates(subset=[col_clave_externa], keep='first')
                .set_index(col_clave_externa)[col_cantidad_externa]
            )
            df_origen[col_destino] = df_origen[col_clave_origen].map(df_mapa)
            
        def mapear_columna_temp(df_mapeo: pd.DataFrame, col_indice: str, col_destino: str, col_clave: str, nombre_col_mapa: str):
             mapa = df_mapeo.drop_duplicates(subset=[col_clave], keep='first').set_index(col_clave)[nombre_col_mapa]
             df_original[col_destino] = df_original[col_indice].map(mapa)

        # 3.1. Mapeo de Cantidad Calculada (Temporal, usando MAX para Atípicos)
        mapear_con_maxima_cantidad_temp(
            df_original, df_externo, COL_CLAVE, NOMBRE_COL_CLAVE_EXTERNA, 
            NOMBRE_COL_CANT_EXTERNA, COL_CANT_CALCULADA
        )
        
        # 3.3. Mapeo de Peso Neto (Temporal para Atípicos)
        mapear_columna_temp(df_peso_neto, material_col_name, COL_PESO_NETO, col_names['material_pn'], col_names['peso_neto_valor'])

        # 5. CÁLCULO DE MANO DE OBRA, PERSONAS Y MÁQUINAS (LÓGICA PYTHON - VALOR FIJO)
        COL_PSTTBJO_MO = 0  
        COL_TIEMPO_MO = 2   
        COL_CANTIDAD_MAQUINAS_MO = 3 
        COL_CANTIDAD_PERSONAS_MO = 4 

        df_mano_obra[COL_PSTTBJO_MO] = df_mano_obra[COL_PSTTBJO_MO].astype(str).str.strip()
        for col_idx in [COL_TIEMPO_MO, COL_CANTIDAD_MAQUINAS_MO, COL_CANTIDAD_PERSONAS_MO]:
            df_mano_obra[col_idx] = pd.to_numeric(df_mano_obra[col_idx], errors='coerce')  

        COL_OP = 'Op.'
        op_col = df_original[COL_OP].astype(str).str.strip()
        indices_terminan_en_1 = op_col.str.endswith('1')
        psttbjo_filtrado = df_original.loc[indices_terminan_en_1, psttbjo_col_name].astype(str).str.strip()

        def mapear_mo_filtros(col_origen: int, col_destino: str):
            mapa = df_mano_obra.drop_duplicates(subset=[COL_PSTTBJO_MO], keep='first').set_index(COL_PSTTBJO_MO)[col_origen]
            df_original.loc[indices_terminan_en_1, col_destino] = psttbjo_filtrado.map(mapa)

        # 5.1. Tiempo de Mano de Obra (Personas * 60)
        df_mano_obra['Calculo_MO_Tiempo'] = df_mano_obra[COL_TIEMPO_MO] * 60
        mapa_mano_obra_tiempo = df_mano_obra.drop_duplicates(subset=[COL_PSTTBJO_MO], keep='first').set_index(COL_PSTTBJO_MO)['Calculo_MO_Tiempo']
        df_original[COL_MANO_OBRA] = np.nan
        df_original.loc[indices_terminan_en_1, COL_MANO_OBRA] = psttbjo_filtrado.map(mapa_mano_obra_tiempo)

        # 5.2. Número de Personas (Columna E)
        df_original[COL_NRO_PERSONAS] = np.nan
        mapear_mo_filtros(COL_CANTIDAD_PERSONAS_MO, COL_NRO_PERSONAS)

        # 5.3. Número de Máquinas (Columna D)
        df_original[COL_NRO_MAQUINAS] = np.nan
        mapear_mo_filtros(COL_CANTIDAD_MAQUINAS_MO, COL_NRO_MAQUINAS)
        
        # 4. Cálculo de Secuencia (VALOR FIJO CALCULADO EN PYTHON)
        df_original[COL_SECUENCIA] = df_original[columna_para_secuencia].astype(str).str.strip().apply(
            lambda x: obtener_secuencia(x, df_secuencias)
        )

        # 7.5 CÁLCULO DE ATÍPICOS 
        cols_agrupamiento = [COL_PESO_NETO, COL_SECUENCIA]
        df_original[COL_PESO_NETO] = pd.to_numeric(df_original[COL_PESO_NETO], errors='coerce')
        df_original[COL_SECUENCIA] = pd.to_numeric(df_original[COL_SECUENCIA], errors='coerce')

        df_original[COL_ATIPICO] = df_original.groupby(cols_agrupamiento, dropna=True).apply(
            detectar_y_marcar_cantidad_atipica
        ).reset_index(level=list(range(len(cols_agrupamiento))), drop=True).fillna(False)
        
        # --------------------------------------------------------------------------
        # REEMPLAZO DE COLUMNAS POR FÓRMULAS DE EXCEL (solo en Hoja de Salida)
        # --------------------------------------------------------------------------

        # --- 3. Fórmulas BUSCARV ---
        
        # 3.1. COL_CANT_CALCULADA (Fórmula)
        formulas_cant_calc = [
            f'=BUSCARV({COL_CLAVE_OUTPUT_EXCEL}{r};\'{HOJA_EXTERNA}\'!{RANGO_EXTERNO_BUSCARV};{COL_CANT_EXTERNA_INDEX};FALSO)' 
            for r in indices_fila_excel
        ]
        df_original[COL_CANT_CALCULADA] = formulas_cant_calc
        
        # 3.2. COL_PORCENTAJE_RECHAZO (Fórmula)
        formulas_rechazo = [
            f'=BUSCARV({COL_CLAVE_OUTPUT_EXCEL}{r};\'{HOJA_EXTERNA}\'!{RANGO_EXTERNO_BUSCARV};{COL_RECHAZO_EXTERNA_INDEX};FALSO)'
            for r in indices_fila_excel
        ]
        df_original[COL_PORCENTAJE_RECHAZO] = formulas_rechazo

        # 3.3. COL_PESO_NETO (Fórmula)
        formulas_peso_neto = [
            f'=BUSCARV({COL_MATERIAL_OUTPUT_EXCEL}{r};\'{HOJA_PESO_NETO}\'!{RANGO_PN_BUSCARV};{COL_PESO_NETO_INDEX};FALSO)'
            for r in indices_fila_excel
        ]
        df_original[COL_PESO_NETO] = formulas_peso_neto

        # --- 6. Suma de Valores (Fórmula SUMA) ---
        formulas_suma = [f'=O{r}+P{r}+R{r}+S{r}' for r in indices_fila_excel]
        df_original[COL_SUMA_VALORES] = formulas_suma
        
        # --- 7. Cálculo de Diferencia (Fórmula Resta) ---
        
        # Truncado de Cantidad base (Columna I) en Python (valor fijo formateado)
        H_str = df_original[NOMBRE_COL_CANTIDAD_BASE].astype(str).str.replace(',', '.', regex=False).str.strip()
        H_float = pd.to_numeric(H_str, errors='coerce')
        H_trunc = np.trunc(H_float)
        
        def formato_sin_decimales_str(x):
            return f"{x:.0f}".replace('.', ',') if pd.notna(x) and pd.api.types.is_number(x) else np.nan

        df_original[NOMBRE_COL_CANTIDAD_BASE] = H_trunc.apply(formato_sin_decimales_str)
        
        # Fórmula de Diferencia: Cantidad base (I) - Cant. base calculada (J)
        formulas_diferencia = [f'={COL_CANT_BASE_OUTPUT_EXCEL}{r}-{COL_CANT_CALC_OUTPUT_EXCEL}{r}' for r in indices_fila_excel]
        df_original[COL_DIFERENCIA] = formulas_diferencia
        
        # 8. Reconstrucción Final y Guardado con Formato
        
        # Formato de valores fijos de Python
        def formato_excel_regional_2_dec(x):
             return f"{x:.2f}".replace('.', ',') if pd.notna(x) and pd.api.types.is_number(x) else x
        
        # Aplicar formato de 2 decimales a Mano de Obra (Valor fijo)
        df_original[COL_MANO_OBRA] = df_original[COL_MANO_OBRA].apply(formato_excel_regional_2_dec)
        
        # Aplicar formato de entero a Personas y Máquinas (Valores fijos)
        df_original[COL_NRO_PERSONAS] = df_original[COL_NRO_PERSONAS].apply(lambda x: formato_sin_decimales_str(x))
        df_original[COL_NRO_MAQUINAS] = df_original[COL_NRO_MAQUINAS].apply(lambda x: formato_sin_decimales_str(x))

        if COL_PSTTBJO_CONCATENADO in df_original.columns:
             df_original = df_original.drop(columns=[COL_PSTTBJO_CONCATENADO])

        df_original_final = df_original.reindex(columns=[c for c in FINAL_COL_ORDER if c in df_original.columns])

        file_original.seek(0)
        wb = load_workbook(file_original)
        
        # Definición de Estilos
        fill_anomalia = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid') 
        fill_encabezado = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid') 
        font_negrita = Font(bold=True)

        # Crear y escribir la hoja principal procesada (AQUÍ SE ESCRIBEN LAS FÓRMULAS)
        if HOJA_SALIDA in wb.sheetnames: del wb[HOJA_SALIDA]
        ws = wb.create_sheet(HOJA_SALIDA)

        for row in dataframe_to_rows(df_original_final, header=True, index=False):
            ws.append(row)

        # 8.2 APLICACIÓN DE FORMATOS Y ATÍPICOS
        COLUMNAS_ENCABEZADO_FORMATO = [
            COL_CANT_CALCULADA, NOMBRE_COL_CANTIDAD_BASE, COL_DIFERENCIA, COL_SUMA_VALORES
        ] + COLUMNAS_A_RESALTAR 

        indices_encabezado = [
            df_original_final.columns.get_loc(col_name) + 1 
            for col_name in COLUMNAS_ENCABEZADO_FORMATO
            if col_name in df_original_final.columns
        ]

        for col_idx in indices_encabezado:
            header_cell = ws.cell(row=1, column=col_idx)
            header_cell.fill = fill_encabezado
            header_cell.font = font_negrita

        # Aplicar el sombreado a los datos de 'Cant. base calculada' (naranja para atípicos)
        try:
            col_cant_calculada_idx = df_original_final.columns.get_loc(COL_CANT_CALCULADA) + 1
        except KeyError:
            col_cant_calculada_idx = 10 

        for r in range(2, len(df_original) + 2):
            if df_original.iloc[r-2][COL_ATIPICO]:
                cell_to_color = ws.cell(row=r, column=col_cant_calculada_idx)
                cell_to_color.fill = fill_anomalia
                
        # 8.3 APLICACIÓN DE FORMATO NUMÉRICO A LAS COLUMNAS CON FÓRMULAS Y VALORES FIJOS
        
        EXCEL_FORMATO_2_DECIMALES = '#,##0.00' 
        EXCEL_FORMATO_PORCENTAJE = '0.00%'
        EXCEL_FORMATO_ENTERO = '0'

        columnas_a_formatear = {
            # Columnas con FÓRMULA
            COL_CANT_CALCULADA: EXCEL_FORMATO_2_DECIMALES, 
            COL_DIFERENCIA: EXCEL_FORMATO_2_DECIMALES,
            COL_SUMA_VALORES: EXCEL_FORMATO_2_DECIMALES,
            COL_PESO_NETO: EXCEL_FORMATO_2_DECIMALES,
            COL_PORCENTAJE_RECHAZO: EXCEL_FORMATO_PORCENTAJE,
            # Columnas con VALOR FIJO (Calculado en Python)
            COL_MANO_OBRA: EXCEL_FORMATO_2_DECIMALES, 
            COL_SECUENCIA: EXCEL_FORMATO_ENTERO,
            COL_NRO_PERSONAS: EXCEL_FORMATO_ENTERO,
            COL_NRO_MAQUINAS: EXCEL_FORMATO_ENTERO,
        }
        
        for col_name, number_format in columnas_a_formatear.items():
            if col_name in df_original_final.columns:
                col_idx = df_original_final.columns.get_loc(col_name) + 1
                for r in range(2, len(df_original) + 2):
                    cell = ws.cell(row=r, column=col_idx)
                    cell.number_format = number_format
                    
        # --- CREACIÓN DE HOJAS ADICIONALES (SOLO VALORES para evitar el error de Excel) ---
        crear_y_guardar_hoja_solo_valores(wb, df_original, HOJA_LSMW, COLUMNAS_LSMW, fill_encabezado, font_negrita)
        crear_y_guardar_hoja_solo_valores(wb, df_original, HOJA_CAMPOS_USUARIO, COLUMNAS_CAMPOS_USUARIO, fill_encabezado, font_negrita)
        crear_y_guardar_hoja_solo_valores(wb, df_original, HOJA_PORCENTAJE_RECHAZO, COLUMNAS_RECHAZO, fill_encabezado, font_negrita)

        # Guardar el libro de trabajo modificado en un buffer de Bytes
        output_buffer = io.BytesIO()
        wb.save(output_buffer)
        output_buffer.seek(0)

        return True, output_buffer

    except KeyError as ke:
        return False, f"❌ ERROR CRÍTICO DE ENCABEZADO: El script no encontró la columna {ke}. Verifique las hojas y encabezados del archivo original o externo. Asegúrese que el nombre de la hoja principal **{config['HOJA_PRINCIPAL']}** es correcto."
    except IndexError as ie:
        return False, f"❌ ERROR CRÍTICO DE ÍNDICE: Un índice de columna está fuera de rango. Mensaje: {ie}"
    except ValueError as ve:
        if 'sheetname' in str(ve) or 'Worksheet' in str(ve):
            hojas_requeridas = [config['HOJA_PRINCIPAL'], 'Peso neto', HOJA_SECUENCIAS, HOJA_MANO_OBRA, HOJA_EXTERNA]
            return False, f"❌ Error de Lectura de Hoja: Una de las hojas clave ({', '.join(hojas_requeridas)}) no se encontró en los archivos cargados. Mensaje: {ve}"
        return False, f"❌ Ocurrió un error inesperado de valor. Mensaje: {ve}"
    except Exception as e:
        return False, f"❌ Ocurrió un error inesperado. Mensaje: {e}"


# --- INTERFAZ DE STREAMLIT (CON SELECTOR DE HORNO) ---

def main():
    """Configura la interfaz de usuario de Streamlit."""
    st.set_page_config(
        page_title="Automatización Hornos",
        layout="centered",
        initial_sidebar_state="auto"
    )

    st.title("⚙️ Automatización Verificación de datos - HORNOS")
    st.markdown("Seleccione el Horno a procesar y luego cargue los archivos.")

    # SELECCIÓN DEL HORNO
    hornos_disponibles = list(HORNOS_CONFIG.keys())
    selected_horno = st.radio(
        "**1. Seleccione el Horno a Procesar:**",
        hornos_disponibles,
        index=hornos_disponibles.index('HORNO 5') if 'HORNO 5' in hornos_disponibles else 0,
        horizontal=True,
        key="horno_selector" 
    )
    st.markdown("---")
    
    config = HORNOS_CONFIG[selected_horno]
    hoja_principal = config['HOJA_PRINCIPAL']
    hoja_salida = config['HOJA_SALIDA']

    st.subheader(f"2. Carga de Archivos para **{selected_horno}** (Hoja Principal: '{hoja_principal}')")
    
    col1, col2 = st.columns(2)

    with col1:
        file_original = st.file_uploader(
            f"Carga la base de datos original",
            type=['xlsx'],
            help=f"El archivo debe contener las hojas: **{hoja_principal}**, 'Peso neto', '{HOJA_SECUENCIAS}' y '{HOJA_MANO_OBRA}'.",
            key="file_original_uploader" 
        )

    with col2:
        file_externa = st.file_uploader(
            "Carga el archivo externo de toma de información.",
            type=['xlsb', 'xlsx'],
            help=f"El archivo que contiene la hoja '{HOJA_EXTERNA}'.",
            key="file_externa_uploader" 
        )

    st.markdown("---")

    # Botón de ejecución y manejo del proceso
    if st.button(f"▶️ PROCESAR {selected_horno}", type="primary", use_container_width=True, key="process_button"):
        if file_original is None or file_externa is None:
            st.error("Por favor, cargue ambos archivos antes de procesar.")
        else:
            file_buffer_original = io.BytesIO(file_original.getvalue())
            file_buffer_externa = io.BytesIO(file_externa.getvalue())

            with st.spinner(f'Procesando datos y generando reporte para {selected_horno}...'):
                success, resultado = automatizacion_final_diferencia_reforzada(
                    file_buffer_original,
                    file_buffer_externa,
                    selected_horno
                )

            st.markdown("---")

            if success:
                st.success(f"✅ Proceso para **{selected_horno}** completado exitosamente.")

                # Nombre de archivo de salida
                base_name = file_original.name.split('.')[0]
                if hoja_principal in base_name:
                    file_name_output = base_name.replace(hoja_principal, '') + f"{hoja_salida}.xlsx"
                else:
                    file_name_output = f"{base_name}_{hoja_salida}.xlsx"
                
                st.download_button(
                    label="⬇️ Descargar Archivo Procesado",
                    data=resultado,
                    file_name=file_name_output,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True
                )
                st.info(f"El archivo descargado contiene todas las hojas originales más las 4 hojas de reporte: **{hoja_salida}** (con fórmulas), '{HOJA_LSMW}', '{HOJA_CAMPOS_USUARIO}' y '{HOJA_PORCENTAJE_RECHAZO}' (con valores fijos).")
            else:
                st.error("❌ Error en el Proceso")
                st.warning(resultado)
                st.write("Verifique el formato de las hojas y los nombres de las columnas en sus archivos.")

if __name__ == "__main__":
    main()




