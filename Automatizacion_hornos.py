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
    'RESALTAR': ['Mano de obra', 'suma valores', 'Cant_Manual', 'Cant_Maquinas','Cant. base calculada',]
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

# --- 2. FUNCIONES DE LÓGICA ABSTRAÍDA ---

def _obtener_nombre_columna(cols: list, idx: int, default_name: str) -> str:
    return cols[idx] if idx < len(cols) else default_name

def _mapear_df(df_origen: pd.DataFrame, df_mapa: pd.DataFrame, col_clave_origen: str, col_clave_mapa: str, col_valor_mapa: str, col_destino: str, keep_mode: str = 'first'):
    mapa_series = (
        df_mapa.sort_values(by=col_valor_mapa, ascending=(keep_mode == 'first'), na_position='last')
        .drop_duplicates(subset=[col_clave_mapa], keep=keep_mode)
        .set_index(col_clave_mapa)[col_valor_mapa]
    )
    df_origen[col_destino] = df_origen[col_clave_origen].map(mapa_series)

def detectar_y_marcar_cantidad_atipica(grupo: pd.DataFrame) -> pd.Series:
    valores_no_nan = grupo[COL['CANT_CALCULADA']].dropna()
    if valores_no_nan.empty:
        return pd.Series(False, index=grupo.index)
    moda = Counter(valores_no_nan).most_common(1)[0][0]
    return grupo[COL['CANT_CALCULADA']] != moda

def filtrar_operaciones_impares_desde_31(df: pd.DataFrame) -> pd.DataFrame:
    if COL['OP'] not in df.columns: return pd.DataFrame()
    df_temp = df.copy()
    op_num = pd.to_numeric(df_temp[COL['OP']].astype(str).str.strip(), errors='coerce')
    condicion = (op_num.notna()) & (op_num >= 31) & (op_num % 2 != 0)
    return df_temp[condicion]

def obtener_secuencia(puesto_trabajo: str, df_secuencias: pd.DataFrame) -> Union[int, float]:
    psttbjo_str = str(puesto_trabajo).strip()
    try:
        for col_idx in range(df_secuencias.shape[1]):
            col_data = df_secuencias.iloc[:, col_idx].dropna()
            if psttbjo_str in set(col_data.astype(str).str.strip()):
                return col_idx + 1
    except Exception: return np.nan
    return np.nan

# --- 3. FUNCIÓN DE CARGA Y LIMPIEZA ---

def cargar_y_limpiar_datos(file_original: io.BytesIO, file_info_externa: io.BytesIO, nombre_horno: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, str]]:
    config = HORNOS_CONFIG[nombre_horno]
    hoja_principal = config['HOJA_PRINCIPAL']

    cols_original = pd.read_excel(file_original, sheet_name=hoja_principal, nrows=0).columns.tolist()
    file_original.seek(0)
    cols_pn = pd.read_excel(file_original, sheet_name='Peso neto', nrows=0).columns.tolist()
    file_original.seek(0)
    cols_externo = pd.read_excel(file_info_externa, sheet_name='Especif y Rutas', nrows=0).columns.tolist()
    file_info_externa.seek(0)
    
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

    df_original = pd.read_excel(file_original, sheet_name=hoja_principal, dtype={col_names['cant_base_leida']: str})
    file_original.seek(0)
    df_peso_neto = pd.read_excel(file_original, sheet_name='Peso neto')
    file_original.seek(0)
    df_secuencias = pd.read_excel(file_original, sheet_name=COL['HOJA_SALIDA_SECUENCIAS'])
    file_original.seek(0)
    df_mano_obra = pd.read_excel(file_original, sheet_name=COL['HOJA_MANO_OBRA'], header=None, usecols=range(5), names=range(5))
    file_original.seek(0)
    df_externo = pd.read_excel(file_info_externa, sheet_name='Especif y Rutas', usecols=[COL['CLAVE_EXTERNA'], COL['CANT_EXTERNA'], col_names['nombre_col_rechazo_externa']])
    file_info_externa.seek(0)
    
    rename_map = {col_names['cant_base_leida']: COL['CANTIDAD_BASE'], col_names['material']: 'Material', col_names['psttbjo']: 'PstoTbjo', col_names['grplf']: 'GrPlf'}
    df_original.rename(columns={k: v for k, v in rename_map.items() if k in df_original.columns}, inplace=True)
    
    return df_original, df_externo, df_peso_neto, df_secuencias, df_mano_obra, col_names

# --- 4. FUNCIÓN DE CREACIÓN DE HOJAS DINÁMICAS ---

def crear_y_guardar_hoja(wb, df_base_full: pd.DataFrame, nombre_hoja: str, columnas_destino: list, fill_encabezado: PatternFill, font_negrita: Font, hoja_salida_name: str):
    """Crea una hoja donde los datos son FÓRMULAS que apuntan a la hoja principal procesada."""
    
    if nombre_hoja == COL['HOJA_SALIDA_CAMPOS_USUARIO']:
        df_filtrado = filtrar_operaciones_impares_desde_31(df_base_full)
    else:
        df_filtrado = df_base_full.copy()

    if nombre_hoja in wb.sheetnames: del wb[nombre_hoja]
    ws = wb.create_sheet(nombre_hoja)
    ws.append(columnas_destino)
    
    col_to_letter = {col: get_column_letter(i+1) for i, col in enumerate(df_base_full.columns)}

    for r_idx, (orig_idx, _) in enumerate(df_filtrado.iterrows(), start=2):
        for c_idx, col_name in enumerate(columnas_destino, start=1):
            # 1. Valores Fijos
            if nombre_hoja == COL['HOJA_SALIDA_CAMPOS_USUARIO'] and col_name == 'Indicador':
                ws.cell(row=r_idx, column=c_idx, value='x')
            elif nombre_hoja == COL['HOJA_SALIDA_CAMPOS_USUARIO'] and col_name == 'clase de control':
                ws.cell(row=r_idx, column=c_idx, value='ZPP0006')
            elif nombre_hoja == COL['HOJA_SALIDA_RECHAZO'] and col_name == '% rechazo anterior':
                ws.cell(row=r_idx, column=c_idx, value=0)
            
            # 2. Diferencia en Hoja Rechazo
            elif nombre_hoja == COL['HOJA_SALIDA_RECHAZO'] and col_name == COL['DIFERENCIA']:
                col_act_idx = columnas_destino.index(COL['PORCENTAJE_RECHAZO']) + 1
                col_ant_idx = columnas_destino.index('% rechazo anterior') + 1
                ws.cell(row=r_idx, column=c_idx, value=f'={get_column_letter(col_act_idx)}{r_idx}-{get_column_letter(col_ant_idx)}{r_idx}')
            
            # 3. Vínculo a Hoja Principal
            elif col_name in df_base_full.columns:
                letra_orig = col_to_letter[col_name]
                fila_orig = df_base_full.index.get_loc(orig_idx) + 2
                formula = f"=IF('{hoja_salida_name}'!{letra_orig}{fila_orig}=0,\"\",'{hoja_salida_name}'!{letra_orig}{fila_orig})"
                ws.cell(row=r_idx, column=c_idx, value=formula)

    for c_idx, col_name in enumerate(columnas_destino, start=1):
        if col_name in COL['RESALTAR']:
            ws.cell(row=1, column=c_idx).fill = fill_encabezado
            ws.cell(row=1, column=c_idx).font = font_negrita

# --- 5. LÓGICA DE PROCESAMIENTO ---

def automatizacion_final_diferencia_reforzada(file_original: io.BytesIO, file_info_externa: io.BytesIO, nombre_horno: str) -> Tuple[bool, Union[str, io.BytesIO]]:
    config = HORNOS_CONFIG[nombre_horno]
    HOJA_SALIDA = config['HOJA_SALIDA']

    try:
        st.subheader(f"Preparando datos para **{nombre_horno}**... 📊")
        df_original, df_externo, df_peso_neto, df_secuencias, df_mano_obra, col_names = cargar_y_limpiar_datos(file_original, file_info_externa, nombre_horno)
        
        limpiar = lambda s: s.astype(str).str.strip().str.replace(r'\W+', '', regex=True)
        df_original[COL['CLAVE_BUSQUEDA']] = limpiar(df_original['Material']) + limpiar(df_original['GrPlf']) + limpiar(df_original['PstoTbjo'])
        df_externo[COL['CLAVE_EXTERNA']] = limpiar(df_externo[COL['CLAVE_EXTERNA']])
        
        col_sec = 'PstoTbjo'
        if COL['LINEA'] in df_original.columns and df_original[COL['LINEA']].astype(str).str.strip().str.len().gt(0).any():
            l, p = df_original[COL['LINEA']].astype(str).str.strip(), df_original['PstoTbjo'].astype(str).str.strip()
            df_original['PstoTbjo_Concat'] = np.where(l.str.lower().isin(['nan', 'none', '', '-']), p, p + l)
            col_sec = 'PstoTbjo_Concat'
            
        df_original[COL['SECUENCIA']] = [obtener_secuencia(x, df_secuencias) for x in df_original[col_sec]]

        map_max = df_externo.sort_values(by=COL['CANT_EXTERNA'], ascending=False).drop_duplicates(subset=[COL['CLAVE_EXTERNA']], keep='first').set_index(COL['CLAVE_EXTERNA'])[COL['CANT_EXTERNA']]
        df_original[COL['CANT_CALCULADA']] = df_original[COL['CLAVE_BUSQUEDA']].map(map_max)
        _mapear_df(df_original, df_externo, COL['CLAVE_BUSQUEDA'], COL['CLAVE_EXTERNA'], col_names['nombre_col_rechazo_externa'], COL['PORCENTAJE_RECHAZO'])
        _mapear_df(df_original, df_peso_neto, 'Material', col_names['material_pn'], col_names['peso_neto_valor'], COL['PESO_NETO'])

        df_mano_obra['Calc_MO'] = pd.to_numeric(df_mano_obra[2], errors='coerce') * 60
        idx1 = df_original[COL['OP']].astype(str).str.strip().str.endswith('1')
        p_f = df_original.loc[idx1, 'PstoTbjo'].astype(str).str.strip()
        mo_m = df_mano_obra.drop_duplicates(subset=[0]).set_index(0)
        df_original.loc[idx1, COL['MANO_OBRA']] = p_f.map(mo_m['Calc_MO'])
        df_original.loc[idx1, COL['NRO_PERSONAS']] = p_f.map(mo_m[4])
        df_original.loc[idx1, COL['NRO_MAQUINAS']] = p_f.map(mo_m[3])
        
        df_original[COL['CANTIDAD_BASE']] = np.trunc(pd.to_numeric(df_original[COL['CANTIDAD_BASE']].astype(str).str.replace(',', '.'), errors='coerce'))
        df_original[COL['SUMA_VALORES']] = np.nan
        df_original[COL['DIFERENCIA']] = np.nan

        atip = df_original.groupby([COL['PESO_NETO'], COL['SECUENCIA']], dropna=True).apply(detectar_y_marcar_cantidad_atipica)
        df_original[COL['ATIPICO']] = pd.Series(atip.values, index=atip.index.get_level_values(-1)).reindex(df_original.index, fill_value=False)

        df_original_final = df_original.reindex(columns=[c for c in FINAL_COL_ORDER if c in df_original.columns])

        file_original.seek(0)
        wb = load_workbook(file_original)
        f_anom, f_enc, f_bold = PatternFill(start_color='FFA500', fill_type='solid'), PatternFill(start_color='DDEBF7', fill_type='solid'), Font(bold=True)
        
        if HOJA_SALIDA in wb.sheetnames: del wb[HOJA_SALIDA]
        ws = wb.create_sheet(HOJA_SALIDA)
        for row in dataframe_to_rows(df_original_final, header=True, index=False): ws.append(row)
        for c_idx, col_name in enumerate(df_original_final.columns, start=1):
            if col_name in COL['RESALTAR']:
                ws.cell(row=1, column=c_idx).fill = f_enc
                ws.cell(row=1, column=c_idx).font = f_bold

        c_dif = df_original_final.columns.get_loc(COL['DIFERENCIA']) + 1
        l_base = get_column_letter(df_original_final.columns.get_loc(COL['CANTIDAD_BASE']) + 1)
        l_calc = get_column_letter(df_original_final.columns.get_loc(COL['CANT_CALCULADA']) + 1)
        c_sum = df_original_final.columns.get_loc(COL['SUMA_VALORES']) + 1
        l_sums = [get_column_letter(df_original_final.columns.get_loc(c) + 1) for c in COLUMNAS_A_SUMAR if c in df_original_final.columns]

        for r in range(2, len(df_original_final) + 2):
            ws.cell(row=r, column=c_dif, value=f'=ROUNDDOWN({l_base}{r}, 0) - {l_calc}{r}').number_format = '#,##0.00'
            s_exp = f'SUM({",".join([f"{l}{r}" for l in l_sums])})'
            ws.cell(row=r, column=c_sum, value=f'=IF({s_exp}=0,"",{s_exp})').number_format = '#,##0.00'
            if df_original.iloc[r-2][COL['ATIPICO']]:
                ws.cell(row=r, column=df_original_final.columns.get_loc(COL['CANT_CALCULADA']) + 1).fill = f_anom
            for c_idx in range(1, len(df_original_final.columns) + 1):
                ws.cell(row=r, column=c_idx).number_format = '#,##0'

        crear_y_guardar_hoja(wb, df_original_final, COL['HOJA_SALIDA_LSMW'], COLUMNAS_OUTPUT['LSMW'], f_enc, f_bold, HOJA_SALIDA)
        crear_y_guardar_hoja(wb, df_original_final, COL['HOJA_SALIDA_CAMPOS_USUARIO'], COLUMNAS_OUTPUT['CAMPOS_USUARIO'], f_enc, f_bold, HOJA_SALIDA)
        crear_y_guardar_hoja(wb, df_original_final, COL['HOJA_SALIDA_RECHAZO'], COLUMNAS_OUTPUT['RECHAZO'], f_enc, f_bold, HOJA_SALIDA)
        
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return True, out
    except Exception as e: return False, str(e)

# --- 6. INTERFAZ ---

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

                #Nombre del archivo
                full_name = file_original.name.rsplit('.', 1)[0]
                
                # Dividir el nombre donde aparezca el 1er guion
                # El .strip() es para eliminar espacios sobrantes si los hubiera
                base_name = full_name.split('_', 1)[0].strip()
                
                # Crear el nombre de salida
                file_name_output = f"{base_name}_{hoja_salida}.xlsx"
                
               
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










