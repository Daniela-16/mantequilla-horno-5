# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
import streamlit as st
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook 
from openpyxl.workbook.properties import CalcProperties # <--- CORRECCIÓN CLAVE
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
COL_PSTTBJO_CONCATENADO = 'PstoTbjo_Concat' # Nombre temporal para la columna concatenada

# Nombres de hojas a crear (Comunes)
HOJA_SECUENCIAS = 'Secuencias'
HOJA_LSMW = 'lsmw'
HOJA_CAMPOS_USUARIO = 'campos de usuario'
HOJA_PORCENTAJE_RECHAZO = '% de rechazo'
COL_OP = 'Op.'

# Columnas a resaltar en todas las hojas (solicitado por el usuario)
COLUMNAS_A_RESALTAR = [
    COL_MANO_OBRA,
    COL_SUMA_VALORES,
    COL_NRO_PERSONAS,
    COL_NRO_MAQUINAS
]

# Definición de columnas de salida (Comunes)
COLUMNAS_LSMW = [
    'PstoTbjo', 'GrpHRuta', 'CGH', 'Material', COL_CLAVE, 'Ce.', COL_OP,
    COL_CANT_CALCULADA, 'ValPref', 'ValPref1', COL_MANO_OBRA, 'ValPref3',
    COL_SUMA_VALORES, 'ValPref5'
]
COLUMNAS_CAMPOS_USUARIO = [
    'GrpHRuta', 'CGH', 'Material', 'Ce.', COL_OP,
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
IDX_MATERIAL = 2 # Columna C
IDX_GRPLF = 4 # Columna E
IDX_CANTIDAD_BASE_LEIDA = 6 # Columna G
IDX_PSTTBJO = 18 # Columna S (Puesto de Trabajo)
IDX_MATERIAL_PN = 0
IDX_RECHAZO_EXTERNA = 28 # CONSTANTE CORRECTA


# --- FUNCIONES DE LÓGICA (Mantenidas) ---

def detectar_y_marcar_cantidad_atipica(grupo: pd.DataFrame) -> pd.Series:
    
    valores_no_nan = grupo[COL_CANT_CALCULADA].dropna()
    if valores_no_nan.empty:
        return pd.Series(False, index=grupo.index)

    conteo = Counter(valores_no_nan)
    moda = conteo.most_common(1)[0][0]

    es_diferente_a_moda = grupo[COL_CANT_CALCULADA] != moda

    return es_diferente_a_moda

def filtrar_operaciones_impares_desde_31(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filtra el DataFrame para incluir solo filas donde 'Op.' es un número impar >= 31.
    """
    if COL_OP not in df.columns:
        st.warning("Columna 'Op.' no encontrada para aplicar filtro en 'campos de usuario'.")
        return pd.DataFrame()

    df_temp = df.copy()
    
    # 1. Intentar convertir la columna 'Op.' a numérico
    df_temp['Op_Num'] = pd.to_numeric(df_temp[COL_OP].astype(str).str.strip(), errors='coerce')
    
    # 2. Definir la condición: No es NaN AND es >= 31 AND es impar
    condicion_impar_desde_31 = (
        df_temp['Op_Num'].notna() & 
        (df_temp['Op_Num'] >= 31) & 
        (df_temp['Op_Num'] % 2 != 0)
    )
    
    # 3. Aplicar el filtro y eliminar la columna temporal ANTES de devolver
    df_filtrado = df_temp[condicion_impar_desde_31].drop(columns=['Op_Num'])
    
    return df_filtrado


# FUNCIÓN crear_y_guardar_hoja (Mantenida)
def crear_y_guardar_hoja(wb, df_base: pd.DataFrame, nombre_hoja: str, columnas_destino: list, fill_encabezado: PatternFill, font_negrita: Font, hoja_salida_name: str = None):
    """
    Crea y guarda una hoja de cálculo en el workbook, aplicando filtros y fórmulas de vinculación si es LSMW.
    """
    
    # Si la hoja a crear es la de campos de usuario, aplicamos el filtro
    if nombre_hoja == HOJA_CAMPOS_USUARIO:
        df_a_guardar = filtrar_operaciones_impares_desde_31(df_base)
        
    else:
        df_a_guardar = df_base.copy()
    
    if nombre_hoja in wb.sheetnames:
        del wb[nombre_hoja]

    ws = wb.create_sheet(nombre_hoja)

    # 1. Crear el nuevo DataFrame con las columnas solicitadas
    df_nuevo = pd.DataFrame()
    for col in columnas_destino:
        if col in df_a_guardar.columns:
            # Los campos vinculados se ponen como NaN para Openpyxl escriba la fórmula
            # Esta lista asegura que las columnas solicitadas por el usuario se pongan como fórmula
            COLUMNAS_A_VINCULAR = [
                COL_CANT_CALCULADA, 'ValPref', 'ValPref1', COL_MANO_OBRA, 
                'ValPref3', COL_SUMA_VALORES, 'ValPref5'
            ]
            
            if nombre_hoja == HOJA_LSMW and col in COLUMNAS_A_VINCULAR:
                 df_nuevo[col] = np.nan
            else:
                 df_nuevo[col] = df_a_guardar[col]
        elif col == 'Indicador' and nombre_hoja == HOJA_CAMPOS_USUARIO:
            df_nuevo[col] = 'x'
        elif col == 'clase de control' and nombre_hoja == HOJA_CAMPOS_USUARIO:
            df_nuevo[col] = 'ZPP0006'
        else:
            df_nuevo[col] = np.nan
            
    # Manejar el caso de DataFrame vacío (si el filtro no encuentra nada)
    if df_nuevo.empty and not df_a_guardar.empty:
        st.error(f"Fallo al crear el DataFrame para la hoja '{nombre_hoja}'. Verifique los nombres de las columnas: {columnas_destino}")
        df_nuevo = pd.DataFrame(columns=columnas_destino)

    # 2. Escribir el nuevo DataFrame en la hoja (Encabezado y Datos)
    for row in dataframe_to_rows(df_nuevo, header=True, index=False):
        ws.append(row)

    # LÓGICA DE FÓRMULA DE VINCULACIÓN PARA LSMW
    if nombre_hoja == HOJA_LSMW and hoja_salida_name:
        
        # Columnas que deben ser vinculadas a la hoja procesada
        COLUMNAS_A_VINCULAR_LSMW = [
            COL_CANT_CALCULADA, 'ValPref', 'ValPref1', COL_MANO_OBRA,  
            'ValPref3', COL_SUMA_VALORES, 'ValPref5'
        ]
        
        # Todas las columnas vinculadas tendrán el condicional CERO
        COLUMNAS_CON_CONDICIONAL_CERO = COLUMNAS_A_VINCULAR_LSMW 
        
        try:
            df_referencia = df_base # df_original_final (contiene los índices de columna correctos)

            # Iterar sobre las columnas a vincular
            for col_name_to_link in COLUMNAS_A_VINCULAR_LSMW:
                if col_name_to_link not in df_nuevo.columns:
                    continue

                # 1. Obtener la columna de la hoja LSMW donde se colocará la fórmula
                lsmw_col_idx = df_nuevo.columns.get_loc(col_name_to_link) + 1
                
                # 2. Obtener la letra de la columna en la HOJA_SALIDA
                try:
                    source_col_idx = df_referencia.columns.get_loc(col_name_to_link) + 1
                    source_col_letter = get_column_letter(source_col_idx)
                except KeyError:
                    st.warning(f"La columna '{col_name_to_link}' no se encontró en la base de datos de origen ('df_base'). No se puede vincular.")
                    continue

                # 3. Iterar sobre las filas de datos (a partir de la fila 2)
                for r_idx in range(len(df_nuevo)):
                    excel_row = r_idx + 2 # Fila de datos en Excel (Empieza en 2)
                    
                    # Referencia simple a la celda en la hoja de salida
                    referencia_celda = f"'{hoja_salida_name}'!{source_col_letter}{excel_row}"
                    
                    # Aplicar la lógica condicional SI(CELDA=0,"",CELDA)
                    if col_name_to_link in COLUMNAS_CON_CONDICIONAL_CERO:
                        # La fórmula se escribe en sintaxis universal de Excel (coma).
                        # Excel la convierte a punto y coma (;) en el entorno regional del usuario.
                        formula = f"=SI({referencia_celda}=0,\"\",{referencia_celda})"
                    else:
                        formula = f"={referencia_celda}"
                    
                    # Sobrescribir la celda con la fórmula
                    cell = ws.cell(row=excel_row, column=lsmw_col_idx, value=formula)
                    
                    # Aplicar formato numérico a todas las celdas vinculadas
                    cell.number_format = '#,##0.00' # Esto es CRÍTICO para que la celda se muestre con coma decimal.
            
        except KeyError:
            st.error(f"Error al aplicar fórmulas en '{HOJA_LSMW}'. Verifique la existencia de las columnas.")


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

# FUNCIÓN cargar_y_limpiar_datos
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
    
    # Renombrar columnas clave si es necesario (basado en índices leídos)
    if col_names['cant_base_leida'] != NOMBRE_COL_CANTIDAD_BASE:
        df_original = df_original.rename(columns={col_names['cant_base_leida']: NOMBRE_COL_CANTIDAD_BASE})
        col_names['cant_base_leida'] = NOMBRE_COL_CANTIDAD_BASE
    
    if col_names['material'] != 'Material':
        df_original.rename(columns={col_names['material']: 'Material'}, inplace=True)
        col_names['material'] = 'Material'
    
    if col_names['psttbjo'] != 'PstoTbjo':
        df_original.rename(columns={col_names['psttbjo']: 'PstoTbjo'}, inplace=True)
        col_names['psttbjo'] = 'PstoTbjo'

    # Cargar DataFrames adicionales
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
    cols_externo = pd.read_excel(file_info_externa, sheet_name='Especif y Rutas', nrows=0).columns.tolist()
    file_info_externa.seek(0)

    nombre_col_rechazo_externa = cols_externo[IDX_RECHAZO_EXTERNA] if IDX_RECHAZO_EXTERNA < len(cols_externo) else 'Columna AC'
    
    cols_a_leer_externo = [NOMBRE_COL_CLAVE_EXTERNA, NOMBRE_COL_CANT_EXTERNA, nombre_col_rechazo_externa]
    df_externo = pd.read_excel(file_info_externa, sheet_name='Especif y Rutas', header=0, usecols=cols_a_leer_externo)
    file_info_externa.seek(0)

    # Guardar nombres de columnas externas leídas
    col_names['nombre_col_rechazo_externa'] = nombre_col_rechazo_externa

    return df_original, df_externo, df_peso_neto, df_secuencias, df_mano_obra, col_names

# --- FUNCIÓN PRINCIPAL DE PROCESAMIENTO ---

def automatizacion_final_diferencia_reforzada(file_original: io.BytesIO, file_info_externa: io.BytesIO, nombre_horno: str) -> Tuple[bool, Union[str, io.BytesIO]]:
    """
    Ejecuta toda la lógica de procesamiento.
    """
    
    config = HORNOS_CONFIG[nombre_horno]
    HOJA_SALIDA = config['HOJA_SALIDA']
    
    # Orden final de las columnas en la hoja de salida
    FINAL_COL_ORDER = [
        'GrpHRuta', 'CGH', 'Material', COL_CLAVE, 'Ce.', 'GrPlf', COL_OP,
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
    
    COLUMNAS_A_SUMAR = ['ValPref', 'ValPref1', COL_MANO_OBRA, 'ValPref3'] # Columnas que se sumarán en Excel

    try:
        st.write("---")
        st.subheader(f"Preparando datos para **{nombre_horno}**... 📊")

        # 1. Carga y limpieza de datos (Se pasa el nombre del horno)
        df_original, df_externo, df_peso_neto, df_secuencias, df_mano_obra, col_names = cargar_y_limpiar_datos(file_original, file_info_externa, nombre_horno)
        
        # Asegurar nombres de columnas clave 
        material_col_name = 'Material'
        grplf_col_name = col_names['cols_original'][IDX_GRPLF]
        psttbjo_col_name = 'PstoTbjo' 

        # 2. Creación de la Clave de Búsqueda
        def limpiar_col(df: pd.DataFrame, col_name: str) -> pd.Series:
            """Limpia (quita caracteres no alfanuméricos) la columna especificada."""
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
        
        # 1. Verificar si la columna 'Linea' existe por nombre
        linea_existe = COL_LINEA in df_original.columns 

        if linea_existe:
            linea_data = df_original[COL_LINEA]
            linea_limpia = linea_data.astype(str).str.strip()
            
            # Verificar si la columna 'Linea' contiene al menos un valor relevante 
            linea_existe_y_es_relevante = linea_limpia.str.lower().str.contains(r'[a-z0-9]').any()
        else:
            linea_existe_y_es_relevante = False
        
        
        if linea_existe_y_es_relevante:
            st.info(f"⚠️ **Detectada la columna '{COL_LINEA}'**. Aplicando lógica de concatenación **PstoTbjo + Línea** para búsqueda de secuencia en {nombre_horno}.")
            
            psttbjo_limpio = df_original[psttbjo_col_name].astype(str).str.strip()
            
            # Crear la columna concatenada: PstoTbjo + Línea si Línea tiene valor, sino solo PstoTbjo
            df_original[COL_PSTTBJO_CONCATENADO] = np.where(
                # La condición usa una serie de booleanos para detectar celdas vacías/nulas
                linea_limpia.str.lower().isin(['nan', 'none', '']), 
                psttbjo_limpio,
                psttbjo_limpio + linea_limpia
            )
            columna_para_secuencia = COL_PSTTBJO_CONCATENADO
        else:
            st.info("✅ **Columna 'Linea' no detectada o vacía**. Usando solo el Puesto de Trabajo para la búsqueda de secuencia.")
        # ------------------------------------------------------------------


        # 3. Mapeo de Cantidad Calculada, Rechazo y Peso Neto
        
        def mapear_columna(df_mapeo: pd.DataFrame, col_indice: str, col_destino: str, col_clave: str, nombre_col_mapa: str):
            """Realiza el mapeo de una columna externa al DataFrame principal (usa 'first')."""
            mapa = df_mapeo.drop_duplicates(subset=[col_clave], keep='first').set_index(col_clave)[nombre_col_mapa]
            df_original[col_destino] = df_original[col_indice].map(mapa)

        def mapear_con_maxima_cantidad(df_origen: pd.DataFrame, df_externo: pd.DataFrame, col_clave_origen: str, col_clave_externa: str, col_cantidad_externa: str, col_destino: str):
            """
            Realiza el mapeo de la Cantidad Base Calculada, seleccionando el valor máximo.
            """
            # 1. Asegurar la columna de cantidad como numérica
            df_externo[col_cantidad_externa] = pd.to_numeric(df_externo[col_cantidad_externa], errors='coerce')
            
            # 2. Encontrar la Cantidad Máxima por Clave
            df_mapa = (
                df_externo.sort_values(by=col_cantidad_externa, ascending=False)
                .drop_duplicates(subset=[col_clave_externa], keep='first')
                .set_index(col_clave_externa)[col_cantidad_externa]
            )
            
            # 3. Aplicar el mapeo al DataFrame de origen
            df_origen[col_destino] = df_origen[col_clave_origen].map(df_mapa)
        
        
        # 3.1. Mapeo de Cantidad Calculada (usando la nueva lógica de MAX)
        mapear_con_maxima_cantidad(
            df_original, 
            df_externo, 
            COL_CLAVE, 
            NOMBRE_COL_CLAVE_EXTERNA, 
            NOMBRE_COL_CANT_EXTERNA, 
            COL_CANT_CALCULADA
        )

        # 3.2. Mapeo de Porcentaje de Rechazo (usa la lógica original, 'first' dupe)
        mapear_columna(df_externo, COL_CLAVE, COL_PORCENTAJE_RECHAZO, NOMBRE_COL_CLAVE_EXTERNA, col_names['nombre_col_rechazo_externa'])
        
        # 3.3. Mapeo de Peso Neto
        mapear_columna(df_peso_neto, material_col_name, COL_PESO_NETO, col_names['material_pn'], col_names['peso_neto_valor'])

        # 4. Cálculo de Secuencia (Usando la columna determinada)
        
        df_original[COL_SECUENCIA] = df_original[columna_para_secuencia].astype(str).str.strip().apply(
            lambda x: obtener_secuencia(x, df_secuencias)
        )

        # 5. Cálculo de Mano de Obra, Personas y Máquinas
        # Índices de df_mano_obra (A=0, C=2, D=3, E=4)
        COL_PSTTBJO_MO = 0 
        COL_TIEMPO_MO = 2 
        COL_CANTIDAD_MAQUINAS_MO = 3 
        COL_CANTIDAD_PERSONAS_MO = 4 

        # Limpieza de datos en df_mano_obra
        df_mano_obra[COL_PSTTBJO_MO] = df_mano_obra[COL_PSTTBJO_MO].astype(str).str.strip()
        for col_idx in [COL_TIEMPO_MO, COL_CANTIDAD_MAQUINAS_MO, COL_CANTIDAD_PERSONAS_MO]:
            df_mano_obra[col_idx] = pd.to_numeric(df_mano_obra[col_idx], errors='coerce') 

        # Filtro: Solo operaciones que terminan en '1'
        op_col = df_original[COL_OP].astype(str).str.strip()
        indices_terminan_en_1 = op_col.str.endswith('1')
        psttbjo_filtrado = df_original.loc[indices_terminan_en_1, psttbjo_col_name].astype(str).str.strip() # Usamos el PstoTbjo original

        # Mapeos para Mano de Obra, Personas y Máquinas
        def mapear_mo_filtros(col_origen: int, col_destino: str):
            """Genera el mapa y aplica el mapeo solo a las filas filtradas."""
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

        # 6. Suma de Valores (Se reserva el espacio para la fórmula de Excel)
        df_original[COL_SUMA_VALORES] = np.nan 

        # 7. Cálculo de Diferencia y Atípicos
        
        H_str = df_original[NOMBRE_COL_CANTIDAD_BASE].astype(str).str.replace(',', '.', regex=False).str.strip()
        H_float = pd.to_numeric(H_str, errors='coerce')
        # Truncar (eliminar decimales)
        H_trunc = np.trunc(H_float)
        
        I = pd.to_numeric(df_original[COL_CANT_CALCULADA], errors='coerce')

        # *** MODIFICACIÓN CRÍTICA: Asegurar que los datos base se guarden como floats (números) ***
        df_original[NOMBRE_COL_CANTIDAD_BASE] = H_trunc 

        # Rellenar la columna de diferencia con NaN; la fórmula se añadirá después con openpyxl.
        df_original[COL_DIFERENCIA] = np.nan 

        # Atípicos
        df_original[COL_CANT_CALCULADA] = I # I es float, se usa para el cálculo de atípicos
        cols_agrupamiento = [COL_PESO_NETO, COL_SECUENCIA]
        for col in cols_agrupamiento:
            if col not in df_original.columns:
                raise KeyError(f"Columna de agrupamiento '{col}' falta en el DataFrame. Se necesita para calcular atípicos.")

        df_original[COL_ATIPICO] = df_original.groupby(cols_agrupamiento, dropna=True).apply(
            detectar_y_marcar_cantidad_atipica
        ).reset_index(level=list(range(len(cols_agrupamiento))), drop=True).fillna(False)


        # 8. Reconstrucción Final y Guardado con Formato
        
        # --- APLICACIÓN DE VALORES FIJOS PARA CAMPOS DE USUARIO ---
        df_original['Indicador'] = 'x'
        df_original['clase de control'] = 'ZPP0006'
        
        # Si se creó la columna de concatenación, la eliminamos para el output final
        if COL_PSTTBJO_CONCATENADO in df_original.columns:
            df_original = df_original.drop(columns=[COL_PSTTBJO_CONCATENADO])
            
        # Reindexar el DataFrame final con el orden deseado. ESTE ES EL ORDEN DE LA HOJA PROCESADA
        df_original_final = df_original.reindex(columns=[c for c in FINAL_COL_ORDER if c in df_original.columns])

        # Cargar el libro de trabajo desde el buffer (Para mantener las hojas originales)
        file_original.seek(0)
        wb = load_workbook(file_original)

        # Definición de Estilos
        fill_anomalia = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid') # Naranja (Atípicos)
        fill_encabezado = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid') # Azul claro (Calculadas)
        font_negrita = Font(bold=True)

        # Crear y escribir la hoja principal procesada
        if HOJA_SALIDA in wb.sheetnames: del wb[HOJA_SALIDA]
        ws = wb.create_sheet(HOJA_SALIDA)

        # Escribir el DataFrame con valores y NaNs
        for row in dataframe_to_rows(df_original_final, header=True, index=False):
            ws.append(row)

        # --- CÁLCULO DE ÍNDICES Y LETRAS PARA LA FÓRMULA DE EXCEL (HOJA PROCESADA) ---
        
        try:
            col_diferencia_idx = df_original_final.columns.get_loc(COL_DIFERENCIA) + 1 
            col_cant_base_idx = df_original_final.columns.get_loc(NOMBRE_COL_CANTIDAD_BASE) + 1 
            col_cant_calculada_idx = df_original_final.columns.get_loc(COL_CANT_CALCULADA) + 1 
        except KeyError as e:
            st.warning(f"No se pudo determinar la posición exacta de las columnas de diferencia: {e}.")
            return False, f"Error: No se encontró una columna clave para cálculo en la hoja procesada: {e}"

        col_base_letter = get_column_letter(col_cant_base_idx) 
        col_calculada_letter = get_column_letter(col_cant_calculada_idx) 
        
        # --- APLICACIÓN DE FÓRMULA DE EXCEL EN COLUMNA 'diferencia' ---
        
        for r in range(2, len(df_original_final) + 2):
            # Usar REDONDEAR.MENOS para replicar el truncamiento
            formula_dif = f'=REDONDEAR.MENOS({col_base_letter}{r}, 0) - {col_calculada_letter}{r}'
            
            cell = ws.cell(row=r, column=col_diferencia_idx, value=formula_dif)
            cell.number_format = '#,##0.00'

        # --- APLICACIÓN DE FÓRMULA DE SUMA DE VALORES ---
        try:
            col_suma_valores_idx = df_original_final.columns.get_loc(COL_SUMA_VALORES) + 1
            
            col_sum_letters = [
                get_column_letter(df_original_final.columns.get_loc(col) + 1)
                for col in COLUMNAS_A_SUMAR 
                if col in df_original_final.columns
            ]

            for r in range(2, len(df_original_final) + 2):
                sum_expression = f'SUMA({",".join([f"{letter}{r}" for letter in col_sum_letters])})'
                
                formula_sum = f'=SI({sum_expression}=0,"",{sum_expression})'
                
                cell = ws.cell(row=r, column=col_suma_valores_idx, value=formula_sum)
                cell.number_format = '#,##0.00' 
                
        except KeyError as e:
            st.error(f"Error al aplicar fórmula de suma en '{HOJA_SALIDA}'. Una columna de suma no fue encontrada: {e}")


        # 4. APLICACIÓN DE FORMATOS EN HOJA PRINCIPAL
        # --- CRÍTICO: Definimos TODAS las columnas que son VALORES NUMÉRICOS para aplicarles formato de número en Excel
        COLUMNAS_VALOR_NUMERICO = [
            NOMBRE_COL_CANTIDAD_BASE, COL_CANT_CALCULADA, COL_PESO_NETO, COL_SECUENCIA,
            'ValPref', 'ValPref1', 'ValPref2', COL_MANO_OBRA, 'ValPref3', 'ValPref4', 'ValPref5',
            COL_NRO_PERSONAS, COL_NRO_MAQUINAS, COL_PORCENTAJE_RECHAZO
        ]
        
        # Columnas para encabezado
        COLUMNAS_ENCABEZADO_FORMATO = [COL_CANT_CALCULADA, NOMBRE_COL_CANTIDAD_BASE] + COLUMNAS_A_RESALTAR

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
        # Y aplicar formato numérico a las columnas de valor
        
        col_indices_valor_numerico = [
             df_original_final.columns.get_loc(col_name) + 1 
             for col_name in COLUMNAS_VALOR_NUMERICO
             if col_name in df_original_final.columns
        ]
        
        # Iterar sobre las filas de datos (a partir de la fila 2)
        for r in range(2, len(df_original_final) + 2):
            # Aplicar formato numérico a los valores de las columnas clave
            for col_idx in col_indices_valor_numerico:
                cell = ws.cell(row=r, column=col_idx)
                # Asegurar que los números se muestren con el formato regional adecuado
                cell.number_format = '#,##0.00' 
                
            # Aplicar el color de anomalía
            if df_original.iloc[r-2][COL_ATIPICO]:
                cell_to_color = ws.cell(row=r, column=col_cant_calculada_idx)
                cell_to_color.fill = fill_anomalia

        # --- CREACIÓN DE HOJAS ADICIONALES ---
        
        # 1. HOJA LSMW 
        crear_y_guardar_hoja(
            wb, 
            df_original_final, 
            HOJA_LSMW, 
            COLUMNAS_LSMW, 
            fill_encabezado, 
            font_negrita,
            hoja_salida_name=HOJA_SALIDA 
        )
        
        # 2. HOJA CAMPOS DE USUARIO 
        crear_y_guardar_hoja(wb, df_original_final, HOJA_CAMPOS_USUARIO, COLUMNAS_CAMPOS_USUARIO, fill_encabezado, font_negrita)
        
        # 3. HOJA PORCENTAJE DE RECHAZO
        crear_y_guardar_hoja(wb, df_original_final, HOJA_PORCENTAJE_RECHAZO, COLUMNAS_RECHAZO, fill_encabezado, font_negrita)
        
        
        # --- SOLUCIÓN AL PROBLEMA DE RECALCULO: Forzar FullCalcOnLoad ---
        # 
        # Si la propiedad de cálculo no existe, la creamos usando CalcProperties
        if wb.calcProperties is None:
            wb.calcProperties = CalcProperties()
            
        # Esta línea fuerza a Excel a recalcular todas las fórmulas al abrir el archivo.
        wb.calcProperties.fullCalcOnLoad = True 
        # -------------------------------------------


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
            hojas_requeridas = [config['HOJA_PRINCIPAL'], 'Peso neto', HOJA_SECUENCIAS, HOJA_MANO_OBRA, 'Especif y Rutas']
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
            help="El archivo que contiene la hoja 'Especif y Rutas'.",
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
                st.info(f"El archivo descargado contiene todas las hojas originales más las 4 hojas de reporte: **{hoja_salida}**, '{HOJA_LSMW}', '{HOJA_CAMPOS_USUARIO}' y '{HOJA_PORCENTAJE_RECHAZO}'.")
            else:
                st.error("❌ Error en el Proceso")
                st.warning(resultado)
                st.write("Verifique el formato de las hojas y los nombres de las columnas en sus archivos.")

if __name__ == "__main__":
    main()

