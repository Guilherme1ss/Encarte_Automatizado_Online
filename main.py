import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
import numpy as np
import os
import string
import warnings
import unicodedata
import tempfile
from pathlib import Path
import re

# Suprime avisos específicos do openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

def get_unique_filename(path):
    """Recebe um caminho de arquivo e retorna um nome único no mesmo diretório."""
    base, ext = os.path.splitext(path)
    counter = 1
    new_path = path
    while os.path.exists(new_path):
        new_path = f"{base} ({counter}){ext}"
        counter += 1
    return new_path

# --- Função para limpar e converter preços ---
def clean_price_value(value):
    """Limpa string de preço e converte para float, retornando None se inválido."""
    try:
        text = str(value).replace("R$", "").replace(",", ".").strip()
        return float(text)
    except:
        return None

# --- Função para normalização de texto para matching ---
translator = str.maketrans('', '', string.punctuation)
def normalize_text(text):
    """Normaliza texto: lowercase, remove acentos e pontuação."""
    if pd.isna(text):
        return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    text = text.translate(translator)
    return text

# --- Mapeamento de compradores para carrossel ---
buyer_carrossel_map = {
    normalize_text("tatiane santos"): "12202 - Pereciveis",
    normalize_text("irlene"): "12202 - Pereciveis",
    normalize_text("amara"): "12205 - Mercearia Salgada",
    normalize_text("brenda"): "12205 - Mercearia Salgada",
    normalize_text("ana paula"): "12204 - Mercearia Doce",
    normalize_text("nilcélia"): "12204 - Mercearia Doce",
    normalize_text("natalia"): "12208 - Perfumaria",
    normalize_text("sonia"): "12208 - Perfumaria",
    normalize_text("neci"): "12206 - Bebidas",
    normalize_text("joice"): "12206 - Bebidas",
    normalize_text("vanessa"): "12212 - Itens Essenciais",
    normalize_text("elton"): "12212 - Itens Essenciais",
    normalize_text("carina"): "12212 - Itens Essenciais",
    normalize_text("mariana"): "12207 - Limpeza",
    normalize_text("simone"): "12207 - Limpeza",
}

def get_carrossel_value(normalized_buyer, mapping):
    """Retorna o valor do carrossel se chave estiver contida no texto do comprador."""
    if not normalized_buyer:
        return ''
    for key, value in mapping.items():
        if key in normalized_buyer:
            return value
    return ''

# --- Dicionário de correção de nomes de produtos ---
product_name_corrections = {
    r'\bcafe\b': 'CAFÉ',
    r'\bpo\b': 'PÓ',
    r'\bpao\b': 'PÃO',
    r'\bleite ferm\b': 'LEITE FERMENTADO',
    r'\bdesinf\b': 'DESINFETANTE',
    r'\bsabao\b': 'SABÃO',
    r'\bsab barra\b': 'SABONETE EM BARRA',
    r'\bcrm pent\b': "CREME DE PENTEAR",
    r'\bsta clara\b': 'SANTA CLARA',
    r'\bvinho tto\b': 'VINHO TINTO',
    r'\bacucar\b': 'AÇÚCAR',
    r'\bqjo\b': 'QUEIJO',
    r'\bparmesão\b': 'PARMESÃO',
    r'\bfile\b': 'FILÉ',
    r'\bhamb\b': 'HAMBÚRGER',
    r'\bfgo\b': 'FRANGO',
    r'\bespag\b': 'ESPAGUETE',
    r'\blacteo\b': 'LÁCTEO',
    r'\bhig\b': 'HIGIÊNICO',
    r'\bracao\b': 'RAÇÃO',
    r'\bhidrat corp\b': 'HIDRATANTE CORPORAL',
    r'\bprot diario\b': 'PROTETOR DIÁRIO',
    r'\bmarata\b': 'MARATÁ',
    r'\balgodao\b': 'ALGODÃO',
    r'\bype\b': 'YPÊ',
    r'\brefrig\b': 'REFRIGERANTE',

    # Adicione mais correções conforme necessário
}

def correct_product_name(name):
    """Corrige o nome do produto com base no dicionário de correções, retornando em maiúsculo."""
    if pd.isna(name):
        return ""
    corrected_name = str(name).strip()
    for pattern, replacement in product_name_corrections.items():
        corrected_name = re.sub(pattern, replacement, corrected_name, flags=re.IGNORECASE)
    return corrected_name.upper()

def remove_suffix(text):
    """Remove sufixos como _sell out, _faturamento e tudo que vier depois."""
    if pd.isna(text):
        return ""
    # Lista de palavras-chave que indicam sufixos a remover
    keywords = ['sell out', 'faturamento', 'sell in']
    pattern = r'_(' + '|'.join(keywords) + r').*$'
    return re.sub(pattern, '', str(text), flags=re.IGNORECASE).strip()


# Função única para classificar EAN
def classify_ean(ean_str):
    """
    Classifica o EAN retornando uma tupla (tipo_codigo, unidade).
    Regras:
      - Se tinha "/" originalmente → Interno, Quilograma
      - Se todos os códigos < 13 dígitos → Interno, Quilograma
      - Caso contrário → EAN, Unidade
    """
    if not ean_str or pd.isna(ean_str) or not str(ean_str).strip():
        return ("EAN", "Unidade")

    ean_str = str(ean_str)

    # Flag para saber se originalmente havia barra
    had_slash = "/" in ean_str

    # Normalizar: trocar "/" por ";"
    ean_str = ean_str.replace("/", ";")

    # Agora separar em lista
    eans = [e.strip() for e in ean_str.split(';') if e.strip()]
    if not eans:
        return ("EAN", "Unidade")

    lens = [len(e) for e in eans]

    if had_slash or all(l < 13 for l in lens):
        return ("Interno", "Quilograma")
    else:
        return ("EAN", "Unidade")

def get_code_type(ean):
    if pd.isna(ean) or not str(ean).strip():
        return 'EAN'
    ean_str = str(ean)

    if "/" in ean_str:
        return 'Interno'
    eans = [e.strip() for e in ean_str.split(';') if e.strip()]
    if not eans:
        return 'EAN'
    lens = [len(e) for e in eans]
    if all(l < 13 for l in lens):
        return 'Interno'
    else:
        return 'EAN'

# Função principal de montagem do DataFrame
def build_final_dataframe(filtered_df, profile, start_date, end_date, store_map, apply_name_correction):
    df_copy = filtered_df.copy()

      # Primeiro remove sufixos indesejados
    df_copy['descrição do item'] = df_copy['descrição do item'].apply(remove_suffix)

    # Aplicar correção de nomes de produtos se habilitado
    if apply_name_correction:
        df_copy['descrição do item'] = df_copy['descrição do item'].apply(correct_product_name)
    else:
        df_copy['descrição do item'] = df_copy['descrição do item'].apply(
            lambda x: str(x).strip().upper() if not pd.isna(x) else ""
        )

    # Substituir "/" por ";" na coluna ean
    df_copy['ean'] = df_copy['ean'].astype(str).str.replace("/", ";", regex=False)

    # Aplicar função única para tipo de código e unidade
    df_copy[['final_code_type', 'final_unit']] = df_copy['ean'].apply(
        lambda x: pd.Series(classify_ean(x))
    )

    # Normalizar a coluna 'comprador' para mapear o carrossel
    if 'comprador' in df_copy.columns:
        df_copy['comprador_normalized'] = df_copy['comprador'].apply(normalize_text)
    else:
        df_copy['comprador_normalized'] = ''
        st.warning("Coluna 'comprador' não encontrada. 'Carrossel' ficará vazia.")

    df_copy['final_carrossel'] = df_copy['comprador_normalized'].apply(
        lambda x: get_carrossel_value(x, buyer_carrossel_map)
    )

    # Monta DataFrame com as colunas esperadas
    return pd.DataFrame({
        "Nome": df_copy["descrição do item"],
        "Carrossel": df_copy["final_carrossel"],
        "Check-In": "Não",
        "Preço": df_copy["preço de:"],
        "Preço promocional": df_copy["preço por:"],
        "Limite por cliente": 0,
        "Dias para Resgate após ativação": 7,
        "Unidade": df_copy["final_unit"],
        "Não exigir ativação no App": "Ativação automática",
        "Ativar em": start_date.strftime("%d/%m/%Y %H:%M"),
        "Inativar em": end_date.strftime("%d/%m/%Y %H:%M"),
        "URL da imagem": "",
        "Tipo do código": df_copy["final_code_type"],
        "Códigos dos produtos": df_copy["ean"],
        "Tipo Promocional": "De / por",
        "Sobrescrever lojas": "Sim",
        "Lojas": store_map[profile]
    })

# --- Função para mesclar EANs do arquivo ---
def merge_ean_data(df_base, ean_file):
    """Mescla os EANs do arquivo com a tabela base usando a coluna CÓDIGO."""
    try:
        # Determinar o tipo de arquivo pela extensão
        file_extension = os.path.splitext(ean_file.name)[1].lower()
        if file_extension in ['.xlsx', '.xls']:
            df_ean = pd.read_excel(ean_file)
        elif file_extension == '.csv':
            df_ean = pd.read_csv(ean_file, sep=';')
        else:
            st.error("Formato de arquivo de EANs não suportado. Use xlsx, xls ou csv.")
            return df_base
        # Renomear colunas para consistência
        df_ean = df_ean.rename(columns={'CÓDIGO PRODUTO': 'código', 'CÓDIGO EAN': 'ean'})
        # Normalizar a coluna 'código' em ambos os DataFrames, removendo hífens
        df_base['código'] = df_base['código'].astype(str).str.strip().str.replace('-', '')
        df_ean['código'] = df_ean['código'].astype(str).str.strip().str.replace('-', '')
        # Criar uma lista para armazenar as novas linhas
        expanded_rows = []
        for _, row in df_base.iterrows():
            # Encontrar EANs correspondentes
            matching_eans = df_ean[df_ean['código'] == row['código']]['ean']
            new_row = row.copy()
            if not matching_eans.empty:
                # Usar a string concatenada completa
                new_row['ean'] = matching_eans.iloc[0].strip()
            expanded_rows.append(new_row)
        # Criar novo DataFrame com as linhas atualizadas
        df_updated = pd.DataFrame(expanded_rows)
        return df_updated
    except Exception as e:
        st.error(f"Erro ao mesclar dados de EAN: {e}")
        return df_base

# --- Função para listar planilhas disponíveis ---
def list_sheets(uploaded_file):
    """Retorna a lista de planilhas disponíveis no arquivo ou None para CSV."""
    file_extension = os.path.splitext(uploaded_file.name)[1].lower()
    try:
        if file_extension in ['.xlsx', '.xls']:
            # Ler o arquivo Excel sem carregar os dados imediatamente
            xl = pd.ExcelFile(uploaded_file)
            return xl.sheet_names
        elif file_extension == '.csv':
            # CSV não tem planilhas, retornar nome genérico
            return ["Planilha CSV"]
        else:
            st.error("Formato de arquivo não suportado. Use xlsx, xls ou csv.")
            return []
    except Exception as e:
        st.error(f"Erro ao listar planilhas: {e}")
        return []

# --- Função principal para processar a planilha ---
def process_promotions(uploaded_file, ean_file, start_date, end_date, temp_dir, use_ean_file, apply_name_correction, sheet_name):
    profiles = ["GERAL/PREMIUM", "GERAL", "PREMIUM"]
    store_mapping = {
        "GERAL": "4368-4363-4362-4357-4360-4356-4370-4359-4372-4353-4371-4365-4369-4361-4366-4354-4355-4364",
        "PREMIUM": "4373-4358-4367",
        "GERAL/PREMIUM": "4368-4363-4362-4357-4360-4356-4370-4359-4372-4353-4371-4365-4369-4361-4366-4354-4355-4364-4373-4358-4367"
    }

    # Determinar o tipo de arquivo pela extensão
    file_extension = os.path.splitext(uploaded_file.name)[1].lower()
    try:
        if file_extension in ['.xlsx', '.xls']:
            df_base = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=4)
        elif file_extension == '.csv':
            df_base = pd.read_csv(uploaded_file, sep=';', header=4)
        else:
            st.error("Formato de arquivo base não suportado. Use xlsx, xls ou csv.")
            return []
    except Exception as e:
        st.error(f"Erro ao ler o arquivo base: {e}")
        return []

    # Limpar colunas
    df_base.columns = df_base.columns.str.strip().str.replace(r'\s+', ' ', regex=True).str.lower()
    
    # Mesclar com arquivo de EANs, se fornecido e habilitado
    if use_ean_file and ean_file:
        df_base = merge_ean_data(df_base, ean_file)

    # Preencher valores mesclados
    df_base['perfil de loja'] = df_base['perfil de loja'].ffill()
    df_base['tipo ação'] = df_base['tipo ação'].ffill()
    # Filtrar linhas com "CRM" no 'tipo ação'
    df_filtered = df_base[df_base["tipo ação"].str.contains("CRM", case=False, na=False)]

    # Garante pasta temporária existe
    os.makedirs(temp_dir, exist_ok=True)
    output_files = []

    for profile in profiles:
        df_profile = df_filtered[df_filtered["perfil de loja"] == profile].copy()
        # Preencher preços mesclados
        df_profile["preço de:"] = df_profile["preço de:"].ffill()
        df_profile["preço por:"] = df_profile["preço por:"].ffill()
        # Limpar preços
        df_profile["preço de:"] = df_profile["preço de:"].apply(clean_price_value)
        df_profile["preço por:"] = df_profile["preço por:"].apply(clean_price_value)
        # Montar DataFrame final
        df_final = build_final_dataframe(df_profile, profile, start_date, end_date, store_mapping, apply_name_correction)
        # Salvar arquivo Excel
        filename = f"promo_{profile.replace('/', '_')}_CRM.xlsx"
        filepath = os.path.join(temp_dir, filename)
        filepath = get_unique_filename(filepath)
        df_final.to_excel(filepath, index=False)
        output_files.append((filename, filepath))
        st.success(f"✅ Arquivo gerado: {filename}")

    return output_files

# --- Interface Streamlit ---
st.title("Processador de Promoções CRM")
st.write("Faça upload da planilha de promoções (xlsx, xls ou csv) e, opcionalmente, um arquivo com EANs (xlsx, xls ou csv). Selecione as datas do encarte e a planilha desejada.")

# Criar diretório temporário
temp_dir = tempfile.mkdtemp()

# Definir datas padrão
default_start = datetime.today()
default_end = datetime.today() + timedelta(days=7)

# Inputs de data
col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Data de Início do Encarte", value=default_start, format="DD/MM/YYYY")
with col2:
    end_date = st.date_input("Data de Fim do Encarte", value=default_end, format="DD/MM/YYYY")

# Verificar validade das datas
if end_date < start_date:
    st.error("A data de fim não pode ser anterior à data de início.")
else:
    # Checkbox para correção de nomes
    apply_name_correction = st.checkbox("Aplicar correção de nomes de produtos", value=False)

    # Checkbox para arquivo de EANs
    use_ean_file = st.checkbox("Usar arquivo de EANs", value=False)

    # Upload da planilha base
    uploaded_file = st.file_uploader("Selecione o arquivo de ENCARTE CONSOLIDADO", type=["xlsx", "xls", "csv"])
    
    # Seleção de planilha
    selected_sheet = None
    if uploaded_file:
        sheet_names = list_sheets(uploaded_file)
        if sheet_names:
            st.write("Selecione a planilha para processar:")
            selected_sheet = st.selectbox("Planilhas disponíveis", sheet_names)
        else:
            st.error("Nenhuma planilha encontrada no arquivo.")
    
    # Upload opcional do arquivo de EANs, mostrado apenas se o checkbox estiver marcado
    ean_file = None
    if use_ean_file:
        ean_file = st.file_uploader("Selecione o arquivo de EANs (opcional)", type=["xlsx", "xls", "csv"])

    if uploaded_file and selected_sheet:
        if st.button("Processar Promoções"):
            with st.spinner("Processando..."):
                # Converter datas para datetime
                start_date = datetime.combine(start_date, time(0, 0))
                end_date = datetime.combine(end_date, time(23, 59))
                # Processar
                output_files = process_promotions(uploaded_file, ean_file, start_date, end_date, temp_dir, use_ean_file, apply_name_correction, selected_sheet)
                # Oferecer download dos arquivos gerados
                for filename, filepath in output_files:
                    with open(filepath, "rb") as f:
                        st.download_button(
                            label=f"Baixar {filename}",
                            data=f,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )