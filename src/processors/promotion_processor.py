import os
import pandas as pd
import streamlit as st
from io import BytesIO
import warnings

from src.config.config_loader import load_config
from src.processors.header_detector import detect_header_with_scoring
from src.processors.ean_merger import merge_ean_data
from src.processors.dataframe_builder import build_final_dataframe
from src.processors.excel_exporter import export_to_excel
from src.utils.data_utils import fix_if_date, clean_price_value
from src.utils.link_loader import load_links_json
from src.utils.file_utils import get_unique_filename

# Suprime avisos específicos do openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

def process_promotions(uploaded_file, ean_file, link_file, use_default_url, start_date, end_date, temp_dir, use_ean_file, use_link_file, apply_name_correction, sheet_name):
    """Função principal para processar as promoções"""
    
    # Carregar configurações
    required_columns, buyer_carrossel_map, product_name_corrections = load_config()
    if required_columns is None or buyer_carrossel_map is None or product_name_corrections is None:
        st.stop()
    
    profiles = ["GERAL/PREMIUM", "GERAL", "PREMIUM"]
    store_mapping = {
        "GERAL": "4368-4363-4362-4357-4360-4356-4370-4359-4372-4353-4371-4365-4369-4361-4366-4354-4355-4364",
        "PREMIUM": "4373-4358-4367-5839",
        "GERAL/PREMIUM": "4368-4363-4362-4357-4360-4356-4370-4359-4372-4353-4371-4365-4369-4361-4366-4354-4355-4364-4373-4358-4367-5839"
    }

    file_extension = os.path.splitext(uploaded_file.name)[1].lower()
    try:
        if file_extension in ['.xlsx', '.xls']:
            temp_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, dtype={'ean': str})
        elif file_extension == '.csv':
            temp_df = pd.read_csv(uploaded_file, sep=';', header=None, dtype={'ean': str})
        else:
            st.error("Formato de arquivo base não suportado. Use xlsx, xls ou csv.")
            return []
    except Exception as e:
        st.error(f"Erro ao ler o arquivo base: {e}")
        return []

    header_row, errors = detect_header_with_scoring(temp_df, required_columns)
    if errors:
        for msg in errors:
            st.error(msg)
        return []

    try:
        if file_extension in ['.xlsx', '.xls']:
            df_base = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row, dtype={'ean': str})
        else:
            df_base = pd.read_csv(uploaded_file, sep=';', header=header_row, dtype={'ean': str})
    except Exception as e:
        st.error(f"Erro ao ler o arquivo com o cabeçalho detectado: {e}")
        return []

    df_base.columns = df_base.columns.str.strip().str.replace(r'\s+', ' ', regex=True).str.lower()
    
    if 'código' in df_base.columns:
        df_base['código'] = df_base['código'].apply(fix_if_date)
    if 'ean' in df_base.columns:
        df_base['ean'] = df_base['ean'].apply(fix_if_date)

    df_base['ean_original_encarte'] = df_base['ean']
    df_base["preço de:"] = df_base["preço de:"].apply(clean_price_value)
    df_base["preço por:"] = df_base["preço por:"].apply(clean_price_value)
    df_base["copied_preço_de"] = False
    df_base["copied_preço_por"] = False

    # Copiar preços de linhas anteriores quando necessário
    for i in range(len(df_base)):
        if pd.isna(df_base.iloc[i]["preço de:"]) and i > 0:
            current_ean = str(df_base.iloc[i]["ean"]) if not pd.isna(df_base.iloc[i]["ean"]) else ""
            prev_ean = str(df_base.iloc[i-1]["ean"]) if not pd.isna(df_base.iloc[i-1]["ean"]) else ""
            if current_ean[:7] == prev_ean[:7] and not pd.isna(df_base.iloc[i-1]["preço de:"]):
                df_base.iloc[i, df_base.columns.get_loc("preço de:")] = df_base.iloc[i-1]["preço de:"]
                df_base.iloc[i, df_base.columns.get_loc("copied_preço_de")] = True
        
        if pd.isna(df_base.iloc[i]["preço por:"]) and i > 0:
            current_ean = str(df_base.iloc[i]["ean"]) if not pd.isna(df_base.iloc[i]["ean"]) else ""
            prev_ean = str(df_base.iloc[i-1]["ean"]) if not pd.isna(df_base.iloc[i-1]["ean"]) else ""
            if current_ean[:7] == prev_ean[:7] and not pd.isna(df_base.iloc[i-1]["preço por:"]):
                df_base.iloc[i, df_base.columns.get_loc("preço por:")] = df_base.iloc[i-1]["preço por:"]
                df_base.iloc[i, df_base.columns.get_loc("copied_preço_por")] = True

    if use_ean_file and ean_file:
        df_base = merge_ean_data(df_base, ean_file)

    df_base['perfil de loja'] = df_base['perfil de loja'].ffill()
    df_base['tipo ação'] = df_base['tipo ação'].ffill()
    df_filtered = df_base[df_base["tipo ação"].str.contains("CRM", case=False, na=False)]

    os.makedirs(temp_dir, exist_ok=True)
    output_files = []

    link_map = {}
    if use_link_file:
        if use_default_url:
            try:
                link_map = load_links_json("data/default_url.json")
            except FileNotFoundError:
                st.error("Repositório de links não encontrado no diretório do projeto.")
        elif link_file:
            link_map = load_links_json(link_file)

    for profile in profiles:
        df_profile = df_filtered[df_filtered["perfil de loja"] == profile].copy()
        if df_profile.empty:
            st.warning(f"Nenhuma linha encontrada para o perfil {profile}. Pulando geração.")
            continue
        
        df_final = build_final_dataframe(
            df_profile, profile, start_date, end_date, store_mapping, 
            apply_name_correction, link_map, buyer_carrossel_map, product_name_corrections
        )
        if df_final is None or df_final.empty:
            st.warning(f"O DataFrame final do perfil {profile} está vazio. Pulando exportação.")
            continue

        filename = f"promo_{profile.replace('/', '_')}_CRM.xlsx"
        filepath = os.path.join(temp_dir, filename)
        filepath = get_unique_filename(filepath)

        export_to_excel(df_final, filepath)

        with open(filepath, "rb") as f:
            output = BytesIO(f.read())
            output.seek(0)

        output_files.append((filename, output))
        st.success(f"✅ Arquivo gerado: {filename}")

    return output_files