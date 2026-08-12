# -*- coding: utf-8 -*-
"""
Created on Wed Aug 12 15:57:42 2026

@author: acer
"""

import streamlit as st
import pandas as pd
import datetime
import re
from supabase import create_client, Client

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="MTC Southbound - Validador e Ingesta", layout="wide")
st.title("🚢 Validador y Carga MTC Southbound (Supabase)")

# --- CONEXIÓN A SUPABASE ---
@st.cache_resource
def init_supabase() -> Client:
    # Llama a las variables estrictamente desde Streamlit Secrets
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = init_supabase()

# --- FUNCIONES CORE ---
def clean_date(val):
    if pd.isna(val): return None
    if isinstance(val, str) and val.strip().lower() in ['nd', 'n/a', '#n/a', 'na', '-', '']: return None
    try:
        dt = pd.to_datetime(val)
        return dt.date() if not pd.isna(dt) else None
    except: return None

def clean_val(val):
    if pd.isna(val): return None
    if isinstance(val, str) and val.strip().lower() in ['nd', 'n/a', '#n/a', 'na', '-', '']: return None
    if isinstance(val, pd.Timestamp): return val.to_pydatetime()
    return val

def get_date_from_val(val):
    if val is None: return None
    if isinstance(val, (datetime.datetime, datetime.date)): return val.date() if hasattr(val, 'date') else val
    try:
        dt = pd.to_datetime(val)
        return dt.date() if not pd.isna(dt) else None
    except: return None

# --- ESTILOS VISUALES PARA STREAMLIT ---
def apply_status_color(row):
    """Aplica colores a la fila basados en las discrepancias."""
    styles = [''] * len(row)
    
    try:
        pickup_idx = row.index.get_loc('PICK UP DATE')
        pickup_diff_idx = row.index.get_loc('BD GATE OUT (Diff)')
        return_idx = row.index.get_loc('TERMINAL EMPTY RETURN VALIDATION ')
        return_diff_idx = row.index.get_loc('BD EMPTY TERM (Diff)')
        
        # Lógica Pick Up
        if pd.isna(row['PICK UP DATE']) or pd.isna(row.get('Base_Gate_Out')):
            styles[pickup_idx] = 'background-color: #e0e0e0; color: black' # Gris
        elif not pd.isna(row['BD GATE OUT (Diff)']):
            styles[pickup_idx] = 'background-color: #fffacd; color: black' # Amarillo
            styles[pickup_diff_idx] = 'background-color: #fffacd; color: black'
        else:
            styles[pickup_idx] = 'background-color: #e6ffe6; color: black' # Verde
            
        # Lógica Return Empty
        if pd.isna(row['TERMINAL EMPTY RETURN VALIDATION ']) or pd.isna(row.get('Base_Empty_Term')):
            styles[return_idx] = 'background-color: #e0e0e0; color: black' # Gris
        elif not pd.isna(row['BD EMPTY TERM (Diff)']):
            styles[return_idx] = 'background-color: #fffacd; color: black' # Amarillo
            styles[return_diff_idx] = 'background-color: #fffacd; color: black'
        else:
            styles[return_idx] = 'background-color: #e6ffe6; color: black' # Verde
            
    except KeyError:
        pass # Ignorar si las columnas no están
        
    return styles

# --- INTERFAZ PRINCIPAL ---
st.markdown("### 1. Cargar Base de Datos Principal")
base_file = st.file_uploader("Sube 'Nueva base de datos AMTC (6).xlsx' (Para cruce de datos)", type=['xlsx'])

base_dict = {}
if base_file:
    with st.spinner('Cargando diccionario base...'):
        df_base = pd.read_excel(base_file, sheet_name='TJ-LB')
        for index, row in df_base.iterrows():
            container = str(row.iloc[0]).strip()
            if container == 'nan' or not container: continue
            base_dict[container] = {
                'gate_out': clean_date(row.get('GATE OUT LB')),
                'empty_term': clean_date(row.get('Empty Termination (EI)')),
                'amtc_delivery': clean_val(row.get('AMTC Delivery (Loaded)')),
                'amtc_unload': clean_val(row.get('AMTC Unload Notification (Empty)'))
            }
    st.success(f"Diccionario base cargado con {len(base_dict)} contenedores.")

st.divider()

st.markdown("### 2. Procesar Reportes Semanales (MTC - Southbound Report)")
report_files = st.file_uploader("Sube los reportes a procesar", type=['xlsx'], accept_multiple_files=True)

if report_files and base_dict:
    for file in report_files:
        st.subheader(f"📊 Procesando: {file.name}")
        
        match = re.search(r'(wk-|week\s|wk)(\d{2})', file.name, re.IGNORECASE)
        semana = match.group(2) if match else "XX"
        
        try:
            # 1. Leer y Limpiar
            df = pd.read_excel(file)
            if 'CONTAINER' not in [str(c).upper() for c in df.columns]:
                df = pd.read_excel(file, header=4)
                
            df.columns = [str(c).strip().upper() for c in df.columns]
            
            if 'CONTAINER' not in df.columns:
                st.error("No se encontró la columna CONTAINER")
                continue
                
            df['CONTAINER'] = df['CONTAINER'].astype(str).str.strip()
            df = df[df['CONTAINER'].str.len() == 11]
            
            df = df.map(clean_val)
            
            # 2. Aplicar Lógica de Cruce y Discrepancias
            df['BD GATE OUT (Diff)'] = None
            df['AMTC Delivery (Loaded)'] = None
            df['AMTC Unload Notification (Empty)'] = None
            df['BD EMPTY TERM (Diff)'] = None
            
            df['Base_Gate_Out'] = None
            df['Base_Empty_Term'] = None
            
            return_col_name = 'TERMINAL RETURN EMPTY' if 'TERMINAL RETURN EMPTY' in df.columns else 'RETURN EMPTY DATE'
            if return_col_name not in df.columns and 'TERMINAL EMPTY RETURN VALIDATION' in df.columns:
                 return_col_name = 'TERMINAL EMPTY RETURN VALIDATION'
            
            for index, row in df.iterrows():
                container = row['CONTAINER']
                if container in base_dict:
                    base_data = base_dict[container]
                    
                    df.at[index, 'Base_Gate_Out'] = base_data['gate_out']
                    df.at[index, 'Base_Empty_Term'] = base_data['empty_term']
                    df.at[index, 'AMTC Delivery (Loaded)'] = base_data['amtc_delivery']
                    df.at[index, 'AMTC Unload Notification (Empty)'] = base_data['amtc_unload']
                    
                    # Lógica Pick Up
                    report_pickup = get_date_from_val(row.get('PICK UP DATE'))
                    if report_pickup and base_data['gate_out'] and report_pickup != base_data['gate_out']:
                        df.at[index, 'BD GATE OUT (Diff)'] = base_data['gate_out']
                        
                    # Lógica Return
                    report_return = get_date_from_val(row.get(return_col_name))
                    if report_return and base_data['empty_term'] and report_return != base_data['empty_term']:
                        df.at[index, 'BD EMPTY TERM (Diff)'] = base_data['empty_term']

            # 3. Mapear al Schema de Supabase
            column_mapping = {
                'CONTAINER': 'CONTAINER',
                'PICK UP DATE': 'PICK UP DATE',
                'AMTC Delivery (Loaded)': 'MTC MX DELIVERY',
                'AMTC Unload Notification (Empty)': 'MTC EMPTY NOTIFICATION',
                return_col_name: 'TERMINAL EMPTY RETURN VALIDATION ',
                'CHASIS DAYS INBOUND': 'CHASIS DAYS INBOUND',
                'TOTAL': 'TOTAL'
            }
            
            df_display = df.rename(columns=column_mapping)
            
            # 4. Visualización en Streamlit con lógica de colores
            cols_to_show = ['CONTAINER', 'PICK UP DATE', 'BD GATE OUT (Diff)', 'MTC MX DELIVERY', 'MTC EMPTY NOTIFICATION', 'TERMINAL EMPTY RETURN VALIDATION ', 'BD EMPTY TERM (Diff)']
            df_styled = df_display[cols_to_show].style.apply(apply_status_color, axis=1)
            
            st.markdown("#### Revisión de Discrepancias (Vista Previa)")
            st.dataframe(df_styled, use_container_width=True)
            
            # 5. Carga a Supabase
            if st.button(f"Subir datos verificados (WK-{semana})", type="primary", key=f"up_{file.name}"):
                with st.spinner("Sincronizando con Supabase..."):
                    try:
                        # Limpiar variables temporales antes de subir
                        df_upload = df_display.drop(columns=['Base_Gate_Out', 'Base_Empty_Term', 'BD GATE OUT (Diff)', 'BD EMPTY TERM (Diff)'], errors='ignore')
                        
                        # Convertir fechas a formato ISO para Supabase
                        for col in df_upload.select_dtypes(include=['datetime64', 'datetimetz']).columns:
                             df_upload[col] = df_upload[col].dt.strftime('%Y-%m-%d %H:%M:%S')
                             
                        records = df_upload.where(pd.notnull(df_upload), None).to_dict(orient='records')
                        
                        response = supabase.table('mtc_southbound').insert(records).execute()
                        st.success(f"¡Éxito! {len(records)} registros fueron guardados en la base de datos.")
                    except Exception as e:
                        st.error(f"Fallo en la conexión: {e}")
                        
        except Exception as e:
            st.error(f"Error procesando el archivo: {e}")

elif not base_dict:
    st.info("Sube el archivo base primero para habilitar la comparación.")
