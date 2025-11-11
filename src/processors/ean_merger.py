import os
import pandas as pd
import streamlit as st
from src.utils.data_utils import fix_if_date

def merge_ean_data(df_base, ean_file):
    """Mescla dados de EAN do arquivo externo com o DataFrame base"""
    try:
        file_extension = os.path.splitext(ean_file.name)[1].lower()
        if file_extension in ['.xlsx', '.xls']:
            df_ean = pd.read_excel(ean_file, dtype={'ean': str})
        elif file_extension == '.csv':
            df_ean = pd.read_csv(ean_file, sep=';', dtype={'ean': str})
        else:
            st.error("Formato de arquivo de EANs não suportado. Use xlsx, xls ou csv.")
            return df_base
        
        df_ean = df_ean.rename(columns={'CÓDIGO PRODUTO': 'código', 'CÓDIGO EAN': 'ean'})
        if 'código' in df_ean.columns:
            df_ean['código'] = df_ean['código'].apply(fix_if_date)
        if 'ean' in df_ean.columns:
            df_ean['ean'] = df_ean['ean'].apply(fix_if_date)
        
        df_base['código'] = df_base['código'].astype(str).str.strip().str.replace('-', '')
        df_ean['código'] = df_ean['código'].astype(str).str.strip().str.replace('-', '')
        
        expanded_rows = []
        for _, row in df_base.iterrows():
            new_row = row.copy()
            encarte_ean = str(new_row['ean']).strip() if not pd.isna(new_row['ean']) else ""
            encarte_ean = encarte_ean.replace('/', ';')
            encarte_ean_list = [e.strip() for e in encarte_ean.split(';') if e.strip() and e != 'nan']
            
            matching_eans = df_ean[df_ean['código'] == new_row['código']]['ean']
            
            if not matching_eans.empty:
                ean_list = []
                for ean in matching_eans:
                    if not pd.isna(ean) and str(ean).strip() and str(ean).strip() != 'nan':
                        eans = str(ean).strip().replace('/', ';').split(';')
                        ean_list.extend([e.strip() for e in eans if e.strip()])
                ean_list.extend(encarte_ean_list)
                ean_list = list(dict.fromkeys(ean_list))
                new_row['ean'] = ';'.join(ean_list) if ean_list else encarte_ean
            else:
                new_row['ean'] = ';'.join(encarte_ean_list) if encarte_ean_list else encarte_ean
            
            expanded_rows.append(new_row)
        
        df_updated = pd.DataFrame(expanded_rows)
        return df_updated
    except Exception as e:
        st.error(f"Erro ao mesclar dados de EAN: {e}")
        return df_base