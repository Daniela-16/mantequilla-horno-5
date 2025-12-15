import pandas as pd
import numpy as np
import streamlit as st
import io
import re
from typing import Tuple, Union, Dict, Any
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from collections import Counter

# --- 1. CONSTANTES CENTRALIZADAS Y DEFINICIONES ---
# Nombres de columnas usadas en el DataFrame final
COL = {
    # Columnas calculadas/mapeadas
    'CANT_CALCULADA': 'Cant. base calculada',
    'PESO_NETO': 'peso neto',
    'SECUENCIA': 'secuencia recurso',
    'ATIPICO': 'Atipico_Cant_Calculada',
    'MANO_OBRA': 'Mano de obra',
    'SUMA_VALORES': 'suma valores',
    'PORCENTAJE_RECHAZO': '%de rechazo',
    'NRO_PERSONAS': 'Cant_Manual',
    'NRO_MAQUINAS': 'Cant_Maquinas',
    'CLAVE_BUSQUEDA': 'Clave_Busqueda',
    'DIFERENCIA': 'diferencia',
    'CANTIDAD_BASE': 'Cantidad base', # Columna leída del original
    'OP': 'Op.',
    'LINEA': 'Linea',
    # Columnas de mapeo (archivo externo)
    'CLAVE_EXTERNA': 'MaterialHorno',
    'CANT_EXTERNA': 'CantidadBaseXHora',
    # Nombres de hojas a crear (Comunes)
    'HOJA_SALIDA_SECUENCIAS': 'Secuencias',
    'HOJA_SALIDA_LSMW': 'lsmw',
    'HOJA_SALIDA_CAMPOS_USUARIO': 'campos de usuario',
    'HOJA_SALIDA_RECHAZO': '% de rechazo',
    'HOJA_MANO_OBRA': 'Mano de obra',
    # Estilos/Resaltados
    'RESALTAR': ['Mano de obra', 'suma valores', 'Cant_Manual', 'Cant_Maquinas']
}

# Configuración específica de cada Horno
HORNOS_CONFIG = {'HORNO 1': {'HOJA_PRINCIPAL': 'HORNO 1', 'HOJA_SALIDA': 'HORNO1_procesado'},
    'HORNO 2': {'HOJA_PRINCIPAL': 'HORNO 2', 'HOJA_SALIDA': 'HORNO2_procesado'},
    'HORNO 3': {'HOJA_PRINCIPAL': 'HORNO 3', 'HOJA_SALIDA': 'HORNO3_procesado'},
    'HORNO 4': {'HOJA_PRINCIPAL': 'HORNO 4', 'HOJA_SALIDA': 'HORNO4_procesado'},
    'HORNO 5': {'HOJA_PRINCIPAL': 'HORNO 5', 'HOJA_SALIDA': 'HORNO5_procesado'},
    'HORNO 6': {'HOJA_PRINCIPAL': 'HORNO 6', 'HOJA_SALIDA': 'HORNO6_procesado'},
    'HORNO 7': {'HOJA_PRINCIPAL': 'HORNO 7', 'HOJA_SALIDA': 'HORNO7_procesado'},
    'HORNO 8': {'HOJA_PRINCIPAL': 'HORNO 8', 'HOJA_SALIDA': 'HORNO8_procesado'},
    'HORNO 9': {'HOJA_PRINCIPAL': 'HORNO 9', 'HOJA_SALIDA': 'HORNO9_procesado'},
    'HORNO 10': {'HOJA_PRINCIPAL': 'HORNO 10', 'HOJA_SALIDA': 'HORNO10_procesado'},
    'HORNO 11': {'HOJA_PRINCIPAL': 'HORNO 11', 'HOJA_SALIDA': 'HORNO11_procesado'},
    'HORNO 12': {'HOJA_PRINCIPAL': 'HORNO 12', 'HOJA_SALIDA': 'HORNO12_procesado'},
    'HORNO 18': {'HOJA_PRINCIPAL': 'HORNO 18', 'HOJA_SALIDA': 'HORNO18_procesado'},
}

# Índices fijos del archivo original (para obtener nombres de columnas)
IDX = {
    'MATERIAL': 2, # Columna C
    'GRPLF': 4, # Columna E
    'CANTIDAD_BASE_LEIDA': 6, # Columna G
    'PSTTBJO': 18, # Columna S (Puesto de Trabajo)
    'MATERIAL_PN': 0, # Columna A de 'Peso neto'
    'RECHAZO_EXTERNA': 28, # Columna AC de 'Especif y Rutas'
    'PESO_NETO_VALOR': 2, # Columna C de 'Peso neto'
}

# Columnas de salida (Definiciones más claras)
COLUMNAS_OUTPUT = {
    'LSMW': ['PstoTbjo', 'GrpHRuta', 'CGH', 'Material', COL['CLAVE_BUSQUEDA'], 'Ce.', COL['OP'],
             COL['CANT_CALCULADA'], 'ValPref', 'ValPref1', COL['MANO_OBRA'], 'ValPref3',
             COL['SUMA_VALORES'], 'ValPref5'],
    'CAMPOS_USUARIO': ['GrpHRuta', 'CGH', 'Material', 'Ce.', COL['OP'],
                       'Indicador', 'clase de control', COL['NRO_PERSONAS'], COL['NRO_MAQUINAS']],
    'RECHAZO': ['GrPlf', COL['CLAVE_BUSQUEDA'], 'Material', 'Ce.', 'alternativa', 'alternativa',
                'posición', 'Relevancia', COL['PORCENTAJE_RECHAZO'],
                '% rechazo anterior', COL['DIFERENCIA'], 'Txt.brv.HRuta']
}

# Orden final de las columnas en la hoja de salida (simplificado)
FINAL_COL_ORDER = [
    'GrpHRuta', 'CGH', 'Material', COL['CLAVE_BUSQUEDA'], 'Ce.', 'GrPlf', COL['OP'],
    COL['PORCENTAJE_RECHAZO'], COL['CANTIDAD_BASE'], COL['CANT_CALCULADA'],
    COL['DIFERENCIA'], COL['PESO_NETO'], COL['SECUENCIA'], 'ValPref', 'ValPref1',
    'ValPref2', COL['MANO_OBRA'], 'ValPref3', 'ValPref4', COL['SUMA_VALORES'],
    'ValPref5', 'Campo de usuario cantidad MANUAL', COL['NRO_PERSONAS'],
    'Campo de usuario cantidad MAQUINAS', COL['NRO_MAQUINAS'],
    'Texto breve operación', 'Ctrl', 'VerF', 'PstoTbjo', 'Cl.', 'Gr.fam.pre',
    'Texto breve de material', 'Txt.brv.HRuta', 'Bloq.vers.fabric.', 'Campo usuario unidad',
    'Campo usuario unidad.1', 'Cantidad', 'Contador', 'InBo', 'InBo.1', 'InBo.2',
    'Unnamed: 31', 'I'
]

# Columnas cuyos valores se suman en Excel para 'suma valores'
COLUMNAS_A_SUMAR = ['ValPref', 'ValPref1', COL['MANO_OBRA'], 'ValPref3']

# Columnas que se vinculan con la fórmula universal =IF(CELDA=0,"",CELDA) en LSMW
COLUMNAS_A_VINCULAR_LSMW = [
    COL['CANT_CALCULADA'], 'ValPref', 'ValPref1', COL['MANO_OBRA'],
    'ValPref3', COL['SUMA_VALORES'], 'ValPref5'
]


# --- 2. FUNCIONES DE LÓGICA ABSTRAÍDA ---

def _obtener_nombre_columna(cols: list, idx: int, default_name: str) -> str:
    """Retorna el nombre de la columna en el índice, o un nombre por defecto si no existe."""
    return cols[idx] if idx < len(cols) else default_name

def _mapear_df(df_origen: pd.DataFrame, df_mapa: pd.DataFrame, col_clave_origen: str, col_clave_mapa: str, col_valor_mapa: str, col_destino: str, keep_mode: str = 'first'):
    """Función utilitaria para realizar mapeos (vlookup) de forma concisa."""
    mapa_series = (
        df_mapa.sort_values(by=col_valor_mapa, ascending=(keep_mode == 'first'), na_position='last')
        .drop_duplicates(subset=[col_clave_mapa], keep=keep_mode)
        .set_index(col_clave_mapa)[col_valor_mapa]
    )
    df_origen[col_destino] = df_origen[col_clave_origen].map(mapa_series)

def detectar_y_marcar_cantidad_atipica(grupo: pd.DataFrame) -> pd.Series:
    """Detecta si 'Cant. base calculada' es atípica (diferente a la moda).
       Retorna una Serie booleana alineada con el índice del grupo."""
    valores_no_nan = grupo[COL['CANT_CALCULADA']].dropna()
    if valores_no_nan.empty:
        return pd.Series(False, index=grupo.index)

    moda = Counter(valores_no_nan).most_common(1)[0][0]
    return grupo[COL['CANT_CALCULADA']] != moda

def filtrar_operaciones_impares_desde_31(df: pd.DataFrame) -> pd.DataFrame:
    """Filtra filas donde 'Op.' es un número impar >= 31."""
    if COL['OP'] not in df.columns:
        return pd.DataFrame()
    
    df_temp = df.copy()
    op_num = pd.to_numeric(df_temp[COL['OP']].astype(str).str.strip(), errors='coerce')
    
    condicion = (op_num.notna()) & (op_num >= 31) & (op_num % 2 != 0)
    
    return df_temp[condicion]

def obtener_secuencia(puesto_trabajo: str, df_secuencias: pd.DataFrame) -> Union[int, float]:
    """Busca la secuencia del puesto de trabajo en la hoja 'Secuencias'."""
    psttbjo_str = str(puesto_trabajo).strip()
    
    # La columna de la secuencia es el índice de la columna en df_secuencias + 1
    for col_idx in range(df_secuencias.shape[1]):
        col_data = df_secuencias.iloc[:, col_idx].dropna().astype(str).str.strip()
        if psttbjo_str in set(col_data):
            return col_idx + 1

    return np.nan

# --- 3. FUNCIÓN DE CARGA Y LIMPIEZA SIMPLIFICADA ---

def cargar_y_limpiar_datos(file_original: io.BytesIO, file_info_externa: io.BytesIO, nombre_horno: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, str]]:
    """Carga todos los DataFrames necesarios y prepara los nombres de columnas."""
    config = HORNOS_CONFIG[nombre_horno]
    hoja_principal = config['HOJA_PRINCIPAL']

    # Lectura inicial para obtener nombres de columnas
    cols_original = pd.read_excel(file_original, sheet_name=hoja_principal, nrows=0).columns.tolist()
    file_original.seek(0)
    cols_pn = pd.read_excel(file_original, sheet_name='Peso neto', nrows=0).columns.tolist()
    file_original.seek(0)
    cols_externo = pd.read_excel(file_info_externa, sheet_name='Especif y Rutas', nrows=0).columns.tolist()
    file_info_externa.seek(0)
    
    # Mapeo de nombres originales a nombres estandarizados (usando el diccionario IDX)
    col_names = {
        'cant_base_leida': _obtener_nombre_columna(cols_original, IDX['CANTIDAD_BASE_LEIDA'], COL['CANTIDAD_BASE']),
        'material': _obtener_nombre_columna(cols_original, IDX['MATERIAL'], 'Material'),
        'psttbjo': _obtener_nombre_columna(cols_original, IDX['PSTTBJO'], 'PstoTbjo'),
        'grplf': _obtener_nombre_columna(cols_original, IDX['GRPLF'], 'GrPlf'),
        'material_pn': _obtener_nombre_columna(cols_pn, IDX['MATERIAL_PN'], 'Material'),
        'peso_neto_valor': _obtener_nombre_columna(cols_pn, IDX['PESO_NETO_VALOR'], 'Peso neto'),
        'nombre_col_rechazo_externa': _obtener_nombre_columna(cols_externo, IDX['RECHAZO_EXTERNA'], 'Columna AC'),
        'hoja_principal': hoja_principal
    }

    # Carga de DataFrames
    df_original = pd.read_excel(file_original, sheet_name=hoja_principal, dtype={col_names['cant_base_leida']: str})
    file_original.seek(0)
    
    df_peso_neto = pd.read_excel(file_original, sheet_name='Peso neto')
    file_original.seek(0)

    df_secuencias = pd.read_excel(file_original, sheet_name=COL['HOJA_SALIDA_SECUENCIAS'])
    file_original.seek(0)
    
    # Carga de Mano de Obra (se mantienen los índices numéricos para consistencia con el código original)
    columnas_mano_obra = [0, 1, 2, 3, 4]
    df_mano_obra = pd.read_excel(file_original, sheet_name=COL['HOJA_MANO_OBRA'], header=None,
                                 usecols=range(len(columnas_mano_obra)), names=columnas_mano_obra)
    file_original.seek(0)

    # Carga del Archivo Externo
    cols_a_leer_externo = [COL['CLAVE_EXTERNA'], COL['CANT_EXTERNA'], col_names['nombre_col_rechazo_externa']]
    df_externo = pd.read_excel(file_info_externa, sheet_name='Especif y Rutas', usecols=cols_a_leer_externo)
    file_info_externa.seek(0)
    
    # Renombrar columnas clave en df_original
    rename_map = {
        col_names['cant_base_leida']: COL['CANTIDAD_BASE'],
        col_names['material']: 'Material',
        col_names['psttbjo']: 'PstoTbjo',
        col_names['grplf']: 'GrPlf'
    }
    df_original.rename(columns={k: v for k, v in rename_map.items() if k in df_original.columns}, inplace=True)
    
    return df_original, df_externo, df_peso_neto, df_secuencias, df_mano_obra, col_names

# --- 4. FUNCIÓN DE CREACIÓN DE HOJAS DE EXCEL ---

def crear_y_guardar_hoja(wb, df_base: pd.DataFrame, nombre_hoja: str, columnas_destino: list, fill_encabezado: PatternFill, font_negrita: Font, hoja_salida_name: str = None):
    """Crea y guarda una hoja de cálculo con formato y fórmulas de vinculación (si es LSMW)."""
    
    df_a_guardar = filtrar_operaciones_impares_desde_31(df_base) if nombre_hoja == COL['HOJA_SALIDA_CAMPOS_USUARIO'] else df_base.copy()
    
    if nombre_hoja in wb.sheetnames: del wb[nombre_hoja]
    ws = wb.create_sheet(nombre_hoja)
    
    # 1. Crear el nuevo DataFrame con las columnas solicitadas y valores fijos
    df_nuevo = pd.DataFrame()
    for col in columnas_destino:
        if col in df_a_guardar.columns:
            # Poner NaN en columnas de LSMW que serán fórmulas
            es_col_a_vincular_lsmw = nombre_hoja == COL['HOJA_SALIDA_LSMW'] and col in COLUMNAS_A_VINCULAR_LSMW
            df_nuevo[col] = np.nan if es_col_a_vincular_lsmw else df_a_guardar[col]
        elif nombre_hoja == COL['HOJA_SALIDA_CAMPOS_USUARIO']:
            if col == 'Indicador': df_nuevo[col] = 'x'
            elif col == 'clase de control': df_nuevo[col] = 'ZPP0006'
            else: df_nuevo[col] = np.nan
        else:
            df_nuevo[col] = np.nan
            
    # Manejar DataFrame vacío después del filtro (o si la creación falla)
    if df_nuevo.empty and not df_a_guardar.empty:
        st.warning(f"DataFrame vacío para '{nombre_hoja}' después del filtrado o error de columnas.")
        df_nuevo = pd.DataFrame(columns=columnas_destino)
    elif df_nuevo.empty:
        # Asegurar que las cabeceras se escriban incluso si no hay datos
        ws.append(columnas_destino)
        # Aplicar formato a encabezados (parte 3)
        indices_a_formatear = [
            columnas_destino.index(col) + 1 for col in COL['RESALTAR'] if col in columnas_destino
        ]
        for col_idx in indices_a_formatear:
            ws.cell(row=1, column=col_idx).fill = fill_encabezado
            ws.cell(row=1, column=col_idx).font = font_negrita
        return

    # 2. Escribir el nuevo DataFrame en la hoja (Encabezado y Datos)
    for row in dataframe_to_rows(df_nuevo, header=True, index=False):
        ws.append(row)

    # 3. Lógica de FÓRMULA DE VINCULACIÓN para LSMW
    if nombre_hoja == COL['HOJA_SALIDA_LSMW'] and hoja_salida_name:
        try:
            df_referencia = df_base # Usamos el df_original_final con índices correctos
            
            for col_name_to_link in COLUMNAS_A_VINCULAR_LSMW:
                if col_name_to_link not in df_nuevo.columns: continue

                lsmw_col_idx = df_nuevo.columns.get_loc(col_name_to_link) + 1
                
                # Obtener la letra de la columna en la HOJA_SALIDA
                source_col_idx = df_referencia.columns.get_loc(col_name_to_link) + 1
                source_col_letter = get_column_letter(source_col_idx)

                # Iterar sobre las filas de datos (a partir de la fila 2)
                for r_idx in range(len(df_nuevo)):
                    excel_row = r_idx + 2
                    referencia_celda = f"'{hoja_salida_name}'!{source_col_letter}{excel_row}"
                    
                    # Fórmula universal: =IF(CELDA=0,"",CELDA)
                    formula = f"=IF({referencia_celda}=0,\"\",{referencia_celda})"
                    
                    cell = ws.cell(row=excel_row, column=lsmw_col_idx, value=formula)
                    cell.number_format = '#,##0.00'
        except KeyError as e:
            st.error(f"Error al aplicar fórmulas en '{COL['HOJA_SALIDA_LSMW']}': Columna de referencia {e} no encontrada.")

    # 4. Aplicar Formato a Encabezados Específicos
    indices_a_formatear = [
        df_nuevo.columns.get_loc(col) + 1 for col in COL['RESALTAR'] if col in df_nuevo.columns
    ]

    for col_idx in indices_a_formatear:
        header_cell = ws.cell(row=1, column=col_idx)
        header_cell.fill = fill_encabezado
        header_cell.font = font_negrita


# --- 5. FUNCIÓN PRINCIPAL DE PROCESAMIENTO REFACTORIZADA ---

def automatizacion_final_diferencia_reforzada(file_original: io.BytesIO, file_info_externa: io.BytesIO, nombre_horno: str) -> Tuple[bool, Union[str, io.BytesIO]]:
    """Ejecuta toda la lógica de procesamiento."""

    config = HORNOS_CONFIG[nombre_horno]
    HOJA_SALIDA = config['HOJA_SALIDA']

    try:
        st.subheader(f"Preparando datos para **{nombre_horno}**... 📊")
        
        # 1. Carga y limpieza
        df_original, df_externo, df_peso_neto, df_secuencias, df_mano_obra, col_names = cargar_y_limpiar_datos(file_original, file_info_externa, nombre_horno)
        
        # 2. Creación y Mapeo de Clave de Búsqueda
        # Función de limpieza (quita caracteres no alfanuméricos)
        def limpiar(serie: pd.Series) -> pd.Series:
            return serie.astype(str).str.strip().str.replace(r'\W+', '', regex=True)

        df_original[COL['CLAVE_BUSQUEDA']] = (
            limpiar(df_original['Material']) +
            limpiar(df_original['GrPlf']) +
            limpiar(df_original['PstoTbjo'])
        )
        df_externo[COL['CLAVE_EXTERNA']] = limpiar(df_externo[COL['CLAVE_EXTERNA']])
        
        # 3. Lógica de Secuencia (PstoTbjo vs. PstoTbjo + Linea)
        columna_para_secuencia = 'PstoTbjo'
        if COL['LINEA'] in df_original.columns and limpiar(df_original[COL['LINEA']]).str.len().gt(0).any():
            st.info(f"⚠️ **Detectada la columna '{COL['LINEA']}'**. Aplicando lógica de concatenación **PstoTbjo + Línea** para búsqueda de secuencia.")
            
            linea_limpia = df_original[COL['LINEA']].astype(str).str.strip()
            psttbjo_limpio = df_original['PstoTbjo'].astype(str).str.strip()
            
            # Crear la columna concatenada solo si 'Linea' tiene un valor
            df_original['PstoTbjo_Concat'] = np.where(
                linea_limpia.str.lower().isin(['nan', 'none', '', '-']),
                psttbjo_limpio,
                psttbjo_limpio + linea_limpia
            )
            columna_para_secuencia = 'PstoTbjo_Concat'
            
        df_original[COL['SECUENCIA']] = df_original[columna_para_secuencia].apply(lambda x: obtener_secuencia(x, df_secuencias))

        # 4. Mapeo de Cantidad Calculada, Rechazo y Peso Neto
        # 4.1. Cantidad Calculada (usando MAX)
        _mapear_df(df_original, df_externo, COL['CLAVE_BUSQUEDA'], COL['CLAVE_EXTERNA'], COL['CANT_EXTERNA'], COL['CANT_CALCULADA'], keep_mode='last') # 'last' en sort_values, 'first' en keep
        
        # 4.2. Porcentaje de Rechazo (usando FIRST)
        _mapear_df(df_original, df_externo, COL['CLAVE_BUSQUEDA'], COL['CLAVE_EXTERNA'], col_names['nombre_col_rechazo_externa'], COL['PORCENTAJE_RECHAZO'], keep_mode='first')
        
        # 4.3. Peso Neto (usando FIRST)
        _mapear_df(df_original, df_peso_neto, 'Material', col_names['material_pn'], col_names['peso_neto_valor'], COL['PESO_NETO'], keep_mode='first')

        # 5. Cálculo y Mapeo de Mano de Obra, Personas y Máquinas (solo para Op. que terminan en '1')
        COL_PSTTBJO_MO, COL_TIEMPO_MO, COL_CANTIDAD_MAQUINAS_MO, COL_CANTIDAD_PERSONAS_MO = 0, 2, 3, 4 # Índices de df_mano_obra
        
        df_mano_obra[COL_PSTTBJO_MO] = df_mano_obra[COL_PSTTBJO_MO].astype(str).str.strip()
        df_mano_obra['Calculo_MO_Tiempo'] = pd.to_numeric(df_mano_obra[COL_TIEMPO_MO], errors='coerce') * 60

        indices_terminan_en_1 = df_original[COL['OP']].astype(str).str.strip().str.endswith('1')
        psttbjo_filtrado = df_original.loc[indices_terminan_en_1, 'PstoTbjo'].astype(str).str.strip()
        
        # Mapeos
        map_mo_tiempo = df_mano_obra.drop_duplicates(subset=[COL_PSTTBJO_MO], keep='first').set_index(COL_PSTTBJO_MO)['Calculo_MO_Tiempo']
        map_personas = df_mano_obra.drop_duplicates(subset=[COL_PSTTBJO_MO], keep='first').set_index(COL_PSTTBJO_MO)[COL_CANTIDAD_PERSONAS_MO]
        map_maquinas = df_mano_obra.drop_duplicates(subset=[COL_PSTTBJO_MO], keep='first').set_index(COL_PSTTBJO_MO)[COL_CANTIDAD_MAQUINAS_MO]
        
        df_original.loc[indices_terminan_en_1, COL['MANO_OBRA']] = psttbjo_filtrado.map(map_mo_tiempo)
        df_original.loc[indices_terminan_en_1, COL['NRO_PERSONAS']] = psttbjo_filtrado.map(map_personas)
        df_original.loc[indices_terminan_en_1, COL['NRO_MAQUINAS']] = psttbjo_filtrado.map(map_maquinas)
        
        df_original[COL['SUMA_VALORES']] = np.nan # Campo para fórmula en Excel

        # 6. Cálculo de Cantidad Base Truncada y Atípicos
        cant_base_float = pd.to_numeric(df_original[COL['CANTIDAD_BASE']].astype(str).str.replace(',', '.', regex=False).str.strip(), errors='coerce')
        df_original[COL['CANTIDAD_BASE']] = np.trunc(cant_base_float) # Sobreescribir con el valor truncado
        df_original[COL['DIFERENCIA']] = np.nan # Campo para fórmula en Excel
        
        # Atípicos
        cols_agrupamiento = [COL['PESO_NETO'], COL['SECUENCIA']]
        
        # FIX: Se cambia .reset_index() por .droplevel() para evitar el error de asignación de columnas múltiples.
        df_original[COL['ATIPICO']] = df_original.groupby(cols_agrupamiento, dropna=True).apply(
            detectar_y_marcar_cantidad_atipica
        ).droplevel(level=list(range(len(cols_agrupamiento)))).fillna(False)

        # 7. Reconstrucción Final y Guardado con Formato
        df_original = df_original.drop(columns=['PstoTbjo_Concat']) if 'PstoTbjo_Concat' in df_original.columns else df_original
        
        # Reindexar con el orden final deseado
        df_original_final = df_original.reindex(columns=[c for c in FINAL_COL_ORDER if c in df_original.columns])

        # Cargar/Crear Workbook y Estilos
        file_original.seek(0)
        wb = load_workbook(file_original)
        fill_anomalia = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
        fill_encabezado = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid')
        font_negrita = Font(bold=True)
        
        # Escribir hoja principal procesada
        if HOJA_SALIDA in wb.sheetnames: del wb[HOJA_SALIDA]
        ws = wb.create_sheet(HOJA_SALIDA)
        for row in dataframe_to_rows(df_original_final, header=True, index=False):
            ws.append(row)

        # Aplicación de FÓRMULAS DE EXCEL (DIFERENCIA y SUMA) y FORMATOS
        
        # 7.1. Fórmulas (Diferencia y Suma)
        col_dif_idx = df_original_final.columns.get_loc(COL['DIFERENCIA']) + 1
        col_base_letter = get_column_letter(df_original_final.columns.get_loc(COL['CANTIDAD_BASE']) + 1)
        col_calculada_letter = get_column_letter(df_original_final.columns.get_loc(COL['CANT_CALCULADA']) + 1)
        col_suma_valores_idx = df_original_final.columns.get_loc(COL['SUMA_VALORES']) + 1
        
        col_sum_letters = [
            get_column_letter(df_original_final.columns.get_loc(col) + 1)
            for col in COLUMNAS_A_SUMAR if col in df_original_final.columns
        ]

        for r in range(2, len(df_original_final) + 2):
            # Diferencia: =ROUNDDOWN(BASE, 0) - CALCULADA
            formula_dif = f'=ROUNDDOWN({col_base_letter}{r}, 0) - {col_calculada_letter}{r}'
            ws.cell(row=r, column=col_dif_idx, value=formula_dif).number_format = '#,##0.00'
            
            # Suma: =IF(SUM(...)=0,"",SUM(...))
            sum_expression = f'SUM({",".join([f"{letter}{r}" for letter in col_sum_letters])})'
            formula_sum = f'=IF({sum_expression}=0,"",{sum_expression})'
            ws.cell(row=r, column=col_suma_valores_idx, value=formula_sum).number_format = '#,##0.00'

            # Aplicar color de anomalía
            if df_original.iloc[r-2][COL['ATIPICO']]:
                ws.cell(row=r, column=df_original_final.columns.get_loc(COL['CANT_CALCULADA']) + 1).fill = fill_anomalia
                
        # 7.2. Formato de Encabezados y Números
        columnas_formato_excel = [COL['CANTIDAD_BASE'], COL['CANT_CALCULADA'], COL['PESO_NETO'], COL['SECUENCIA'],
                                  'ValPref', 'ValPref1', 'ValPref2', COL['MANO_OBRA'], 'ValPref3', 'ValPref4', 'ValPref5',
                                  COL['NRO_PERSONAS'], COL['NRO_MAQUINAS'], COL['PORCENTAJE_RECHAZO']] + COL['RESALTAR']
                                  
        indices_formato = {
            df_original_final.columns.get_loc(col) + 1: col for col in columnas_formato_excel if col in df_original_final.columns
        }
        
        for col_idx, col_name in indices_formato.items():
            # Aplicar formato de encabezado a columnas clave
            if col_name in [COL['CANT_CALCULADA'], COL['CANTIDAD_BASE']] + COL['RESALTAR']:
                ws.cell(row=1, column=col_idx).fill = fill_encabezado
                ws.cell(row=1, column=col_idx).font = font_negrita
            
            # Aplicar formato numérico a los datos
            for r in range(2, len(df_original_final) + 2):
                ws.cell(row=r, column=col_idx).number_format = '#,##0'


        # 8. Creación de Hojas Adicionales
        crear_y_guardar_hoja(wb, df_original_final, COL['HOJA_SALIDA_LSMW'], COLUMNAS_OUTPUT['LSMW'], fill_encabezado, font_negrita, hoja_salida_name=HOJA_SALIDA)
        crear_y_guardar_hoja(wb, df_original_final, COL['HOJA_SALIDA_CAMPOS_USUARIO'], COLUMNAS_OUTPUT['CAMPOS_USUARIO'], fill_encabezado, font_negrita)
        crear_y_guardar_hoja(wb, df_original_final, COL['HOJA_SALIDA_RECHAZO'], COLUMNAS_OUTPUT['RECHAZO'], fill_encabezado, font_negrita)
        
        # 9. Guardar
        output_buffer = io.BytesIO()
        wb.save(output_buffer)
        output_buffer.seek(0)

        return True, output_buffer

    except KeyError as ke:
        return False, f"❌ ERROR CRÍTICO DE ENCABEZADO: Columna no encontrada {ke}. Verifique las hojas y encabezados. Hoja principal: **{config['HOJA_PRINCIPAL']}**."
    except IndexError as ie:
        return False, f"❌ ERROR CRÍTICO DE ÍNDICE: Un índice de columna está fuera de rango. Mensaje: {ie}"
    except ValueError as ve:
        if 'sheetname' in str(ve) or 'Worksheet' in str(ve):
            hojas_requeridas = [config['HOJA_PRINCIPAL'], 'Peso neto', COL['HOJA_SALIDA_SECUENCIAS'], COL['HOJA_MANO_OBRA'], 'Especif y Rutas']
            return False, f"❌ Error de Lectura de Hoja: Una de las hojas clave ({', '.join(hojas_requeridas)}) no se encontró. Mensaje: {ve}"
        return False, f"❌ Ocurrió un error inesperado de valor. Mensaje: {ve}"
    except Exception as e:
        return False, f"❌ Ocurrió un error inesperado. Mensaje: {e}"


# --- INTERFAZ DE STREAMLIT (SIN CAMBIOS) ---

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
            help=f"El archivo debe contener las hojas: **{hoja_principal}**, 'Peso neto', '{COL['HOJA_SALIDA_SECUENCIAS']}' y '{COL['HOJA_MANO_OBRA']}'.",
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

                # Mensaje de instrucción clave para el usuario
                st.warning("⚠️ **ACCIÓN REQUERIDA EN EXCEL:** El cálculo automático está desactivado en el archivo. **Deberá abrir el archivo de Excel y presionar la tecla F9 para activar todas las fórmulas** (especialmente en las hojas 'lsmw' y 'HORNOXX_procesado').")

                # Nombre de archivo de salida
                base_name = file_original.name.rsplit('.', 1)[0]
                file_name_output = f"{base_name.replace(hoja_principal, '')}_{hoja_salida}.xlsx" if hoja_principal in base_name else f"{base_name}_{hoja_salida}.xlsx"

                st.download_button(
                    label="⬇️ Descargar Archivo Procesado",
                    data=resultado,
                    file_name=file_name_output,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True
                )
                st.info(f"El archivo descargado contiene todas las hojas originales más las 4 hojas de reporte: **{hoja_salida}**, '{COL['HOJA_SALIDA_LSMW']}', '{COL['HOJA_SALIDA_CAMPOS_USUARIO']}' y '{COL['HOJA_SALIDA_RECHAZO']}'.")
            else:
                st.error("❌ Error en el Proceso")
                st.warning(resultado)
                st.write("Verifique el formato de las hojas y los nombres de las columnas en sus archivos.")

if __name__ == "__main__":
    main()

