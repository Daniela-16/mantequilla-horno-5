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
COL = {
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
    'CANTIDAD_BASE': 'Cantidad base',
    'OP': 'Op.',
    'LINEA': 'Linea',
    'CLAVE_EXTERNA': 'MaterialHorno',
    'CANT_EXTERNA': 'CantidadBaseXHora',
    'HOJA_SALIDA_SECUENCIAS': 'Secuencias',
    'HOJA_SALIDA_LSMW': 'lsmw',
    'HOJA_SALIDA_CAMPOS_USUARIO': 'campos de usuario',
    'HOJA_SALIDA_RECHAZO': '% de rechazo',
    'HOJA_MANO_OBRA': 'Mano de obra',
    'RESALTAR': ['Mano de obra', 'suma valores', 'Cant_Manual', 'Cant_Maquinas']
}

HORNOS_CONFIG = {f'HORNO {i}': {'HOJA_PRINCIPAL': f'HORNO {i}', 'HOJA_SALIDA': f'HORNO{i}_procesado'} for i in range(1, 13)}
HORNOS_CONFIG['HORNO 18'] = {'HOJA_PRINCIPAL': 'HORNO 18', 'HOJA_SALIDA': 'HORNO18_procesado'}

IDX = {
    'MATERIAL': 2, 'GRPLF': 4, 'CANTIDAD_BASE_LEIDA': 6, 'PSTTBJO': 18,
    'MATERIAL_PN': 0, 'RECHAZO_EXTERNA': 28, 'PESO_NETO_VALOR': 2,
}

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

COLUMNAS_A_SUMAR = ['ValPref', 'ValPref1', COL['MANO_OBRA'], 'ValPref3']

# --- 2. FUNCIONES DE LÓGICA ---

def _obtener_nombre_columna(cols: list, idx: int, default_name: str) -> str:
    return cols[idx] if idx < len(cols) else default_name

def _mapear_df(df_origen, df_mapa, col_clave_origen, col_clave_mapa, col_valor_mapa, col_destino, keep_mode='first'):
    mapa_series = (df_mapa.sort_values(by=col_valor_mapa, ascending=(keep_mode == 'first'), na_position='last')
                   .drop_duplicates(subset=[col_clave_mapa], keep=keep_mode).set_index(col_clave_mapa)[col_valor_mapa])
    df_origen[col_destino] = df_origen[col_clave_origen].map(mapa_series)

def detectar_y_marcar_cantidad_atipica(grupo: pd.DataFrame) -> pd.Series:
    valores_no_nan = grupo[COL['CANT_CALCULADA']].dropna()
    if valores_no_nan.empty: return pd.Series(False, index=grupo.index)
    moda = Counter(valores_no_nan).most_common(1)[0][0]
    return grupo[COL['CANT_CALCULADA']] != moda

def filtrar_operaciones_impares_desde_31(df: pd.DataFrame) -> pd.DataFrame:
    if COL['OP'] not in df.columns: return pd.DataFrame()
    op_num = pd.to_numeric(df[COL['OP']].astype(str).str.strip(), errors='coerce')
    return df[(op_num.notna()) & (op_num >= 31) & (op_num % 2 != 0)]

def obtener_secuencia(puesto_trabajo: str, df_secuencias: pd.DataFrame) -> Union[int, float]:
    pst_str = str(puesto_trabajo).strip()
    try:
        for col_idx in range(df_secuencias.shape[1]):
            if pst_str in df_secuencias.iloc[:, col_idx].dropna().astype(str).str.strip().values:
                return col_idx + 1
    except: return np.nan
    return np.nan

# --- 3. CREACIÓN DE HOJAS VINCULADAS ---

def crear_y_guardar_hoja(wb, df_base, nombre_hoja, columnas_destino, fill_encabezado, font_negrita, hoja_salida_name):
    """Crea una hoja vinculada a la principal por fórmulas, con campos determinados."""
    df_temp = filtrar_operaciones_impares_desde_31(df_base) if nombre_hoja == COL['HOJA_SALIDA_CAMPOS_USUARIO'] else df_base.copy()
    
    if nombre_hoja in wb.sheetnames: del wb[nombre_hoja]
    ws = wb.create_sheet(nombre_hoja)
    ws.append(columnas_destino)

    if not df_temp.empty:
        # Obtener letras de columnas en la hoja procesada
        mapeo_letras = {col: get_column_letter(list(df_base.columns).index(col) + 1) 
                        for col in columnas_destino if col in df_base.columns}

        for r_idx, (df_idx, _) in enumerate(df_temp.iterrows(), start=2):
            fila_orig = df_idx + 2
            for c_idx, col_name in enumerate(columnas_destino, start=1):
                # 1. Campos fijos (Indicador y clase de control)
                if nombre_hoja == COL['HOJA_SALIDA_CAMPOS_USUARIO']:
                    if col_name == 'Indicador': ws.cell(row=r_idx, column=c_idx, value='x'); continue
                    if col_name == 'clase de control': ws.cell(row=r_idx, column=c_idx, value='ZPP0006'); continue

                # 2. Campo vacío solicitado (% rechazo anterior)
                if col_name == '% rechazo anterior':
                    ws.cell(row=r_idx, column=c_idx, value=None)
                    continue

                # 3. Diferencia específica en hoja de rechazo
                if nombre_hoja == COL['HOJA_SALIDA_RECHAZO'] and col_name == COL['DIFERENCIA']:
                    l_rechazo = get_column_letter(columnas_destino.index(COL['PORCENTAJE_RECHAZO']) + 1)
                    l_anterior = get_column_letter(columnas_destino.index('% rechazo anterior') + 1)
                    ws.cell(row=r_idx, column=c_idx, value=f"={l_rechazo}{r_idx}-{l_anterior}{r_idx}")
                    continue

                # 4. Vinculación general por fórmula
                if col_name in mapeo_letras:
                    ref = f"'{hoja_salida_name}'!{mapeo_letras[col_name]}{fila_orig}"
                    ws.cell(row=r_idx, column=c_idx, value=f"=IF({ref}=0,\"\",{ref})")

    # Formato de encabezados original
    for c_idx, col in enumerate(columnas_destino, start=1):
        if col in COL['RESALTAR'] or col in [COL['CANT_CALCULADA'], COL['CANTIDAD_BASE']]:
            ws.cell(row=1, column=c_idx).fill = fill_encabezado
            ws.cell(row=1, column=c_idx).font = font_negrita

# --- 4. PROCESO PRINCIPAL ---

def automatizacion_final_diferencia_reforzada(file_original, file_info_externa, nombre_horno):
    config = HORNOS_CONFIG[nombre_horno]
    HOJA_SALIDA = config['HOJA_SALIDA']

    try:
        # Carga
        df_original = pd.read_excel(file_original, sheet_name=config['HOJA_PRINCIPAL'])
        file_original.seek(0)
        df_peso_neto = pd.read_excel(file_original, sheet_name='Peso neto')
        file_original.seek(0)
        df_secuencias = pd.read_excel(file_original, sheet_name=COL['HOJA_SALIDA_SECUENCIAS'])
        file_original.seek(0)
        df_mano_obra = pd.read_excel(file_original, sheet_name=COL['HOJA_MANO_OBRA'], header=None)
        file_original.seek(0)
        df_externo = pd.read_excel(file_info_externa, sheet_name='Especif y Rutas')
        
        # Nombres de columnas dinámicos
        cols_orig = pd.read_excel(file_original, sheet_name=config['HOJA_PRINCIPAL'], nrows=0).columns.tolist()
        col_rechazo_ext = pd.read_excel(file_info_externa, sheet_name='Especif y Rutas', nrows=0).columns[IDX['RECHAZO_EXTERNA']]
        
        df_original.rename(columns={cols_orig[IDX['CANTIDAD_BASE_LEIDA']]: COL['CANTIDAD_BASE'], 
                                    cols_orig[IDX['MATERIAL']]: 'Material',
                                    cols_orig[IDX['PSTTBJO']]: 'PstoTbjo',
                                    cols_orig[IDX['GRPLF']]: 'GrPlf'}, inplace=True)

        # Cálculos base
        limpiar = lambda s: s.astype(str).str.strip().str.replace(r'\W+', '', regex=True)
        df_original[COL['CLAVE_BUSQUEDA']] = limpiar(df_original['Material']) + limpiar(df_original['GrPlf']) + limpiar(df_original['PstoTbjo'])
        df_externo[COL['CLAVE_EXTERNA']] = limpiar(df_externo[COL['CLAVE_EXTERNA']])
        
        df_original[COL['SECUENCIA']] = [obtener_secuencia(x, df_secuencias) for x in df_original['PstoTbjo']]
        
        map_cant = df_externo.sort_values(by=COL['CANT_EXTERNA'], ascending=False).drop_duplicates(COL['CLAVE_EXTERNA']).set_index(COL['CLAVE_EXTERNA'])[COL['CANT_EXTERNA']]
        df_original[COL['CANT_CALCULADA']] = df_original[COL['CLAVE_BUSQUEDA']].map(map_cant)
        _mapear_df(df_original, df_externo, COL['CLAVE_BUSQUEDA'], COL['CLAVE_EXTERNA'], col_rechazo_ext, COL['PORCENTAJE_RECHAZO'])
        
        # Mano de Obra
        idx_1 = df_original[COL['OP']].astype(str).str.strip().str.endswith('1')
        map_mo = df_mano_obra.astype(str).drop_duplicates(subset=[0]).set_index(0)
        df_original.loc[idx_1, COL['MANO_OBRA']] = df_original.loc[idx_1, 'PstoTbjo'].map(lambda x: float(map_mo.loc[x, 2])*60 if x in map_mo.index else np.nan)
        df_original.loc[idx_1, COL['NRO_PERSONAS']] = df_original.loc[idx_1, 'PstoTbjo'].map(lambda x: map_mo.loc[x, 4] if x in map_mo.index else np.nan)
        df_original.loc[idx_1, COL['NRO_MAQUINAS']] = df_original.loc[idx_1, 'PstoTbjo'].map(lambda x: map_mo.loc[x, 3] if x in map_mo.index else np.nan)

        # Atípicos
        atipicos = df_original.groupby([pd.read_excel(file_original, sheet_name='Peso neto').columns[IDX['PESO_NETO_VALOR']], COL['SECUENCIA']], dropna=False).apply(detectar_y_marcar_cantidad_atipica)
        df_original[COL['ATIPICO']] = pd.Series(atipicos.values, index=atipicos.index.get_level_values(-1)).reindex(df_original.index, fill_value=False)

        # Preparar Hoja Procesada
        df_final = df_original.reindex(columns=[c for c in FINAL_COL_ORDER if c in df_original.columns])
        wb = load_workbook(file_original)
        f_anomalia = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid') # Naranja original
        f_encabezado = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid') # Azul original
        font_b = Font(bold=True)

        if HOJA_SALIDA in wb.sheetnames: del wb[HOJA_SALIDA]
        ws_p = wb.create_sheet(HOJA_SALIDA)
        for r in dataframe_to_rows(df_final, index=False, header=True): ws_p.append(r)

        # Fórmulas y resaltados en Procesada
        c_dif = df_final.columns.get_loc(COL['DIFERENCIA']) + 1
        c_base = get_column_letter(df_final.columns.get_loc(COL['CANTIDAD_BASE']) + 1)
        c_calc = get_column_letter(df_final.columns.get_loc(COL['CANT_CALCULADA']) + 1)
        c_sum_idx = df_final.columns.get_loc(COL['SUMA_VALORES']) + 1
        c_sum_ls = [get_column_letter(df_final.columns.get_loc(c) + 1) for c in COLUMNAS_A_SUMAR if c in df_final.columns]

        for r in range(2, len(df_final) + 2):
            ws_p.cell(row=r, column=c_dif, value=f'=ROUNDDOWN({c_base}{r}, 0) - {c_calc}{r}')
            s_exp = f'SUM({",".join([f"{l}{r}" for l in c_sum_ls])})'
            ws_p.cell(row=r, column=c_sum_idx, value=f'=IF({s_exp}=0,"",{s_exp})')
            if df_original.iloc[r-2][COL['ATIPICO']]:
                ws_p.cell(row=r, column=df_final.columns.get_loc(COL['CANT_CALCULADA']) + 1).fill = f_anomalia

        # Generar Hojas Vinculadas
        crear_y_guardar_hoja(wb, df_final, COL['HOJA_SALIDA_LSMW'], COLUMNAS_OUTPUT['LSMW'], f_encabezado, font_b, HOJA_SALIDA)
        crear_y_guardar_hoja(wb, df_final, COL['HOJA_SALIDA_CAMPOS_USUARIO'], COLUMNAS_OUTPUT['CAMPOS_USUARIO'], f_encabezado, font_b, HOJA_SALIDA)
        crear_y_guardar_hoja(wb, df_final, COL['HOJA_SALIDA_RECHAZO'], COLUMNAS_OUTPUT['RECHAZO'], f_encabezado, font_b, HOJA_SALIDA)

        res_buf = io.BytesIO()
        wb.save(res_buf)
        res_buf.seek(0)
        return True, res_buf
    except Exception as e: return False, str(e)

# --- 5. INTERFAZ STREAMLIT ---
def main():
    st.title("⚙️ Automatización Hornos")
    horno = st.radio("Seleccione Horno", list(HORNOS_CONFIG.keys()), index=4, horizontal=True)
    f1, f2 = st.file_uploader("Base", type=['xlsx']), st.file_uploader("Externo", type=['xlsx', 'xlsb'])
    if st.button("PROCESAR"):
        if f1 and f2:
            ok, res = automatizacion_final_diferencia_reforzada(io.BytesIO(f1.read()), io.BytesIO(f2.read()), horno)
            if ok: 
                st.success("Completado.")
                st.download_button("Descargar", res, f"Reporte_{horno}.xlsx")
            else: st.error(res)

if __name__ == "__main__": main()





