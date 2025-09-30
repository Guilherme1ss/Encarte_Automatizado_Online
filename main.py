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
import openpyxl
from openpyxl.styles import PatternFill
import math
from io import BytesIO

# Suprime avisos específicos do openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

# Função para corrigir códigos interpretados como datas
def fix_if_date(value):
    if pd.isna(value):
        return value
    if isinstance(value, (datetime, pd.Timestamp)):
        # Converte data para 'ano-mês' (ex: 2050-8, sem zero à esquerda no mês)
        return f"{value.year}-{value.month}"
    else:
        return str(value)

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
    r'\bfeijao\b': 'FEIJÃO',
    r'\bleite ferm\b': 'LEITE FERMENTADO',
    r'\bdesinf\b': 'DESINFETANTE',
    r'\bsabao\b': 'SABÃO',
    r'\bsab barra\b': 'SABONETE EM BARRA',
    r'\bcrm pent\b': "CREME DE PENTEAR",
    r'\bsta clara\b': 'SANTA CLARA',
    r'\bvinho bco\b': 'VINHO BRANCO',
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
    r'\bamac roupa\b': 'AMACIANTE DE ROUPA',
    r'\bdesod aer\b': 'DESODORANTE AER',
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

def classify_ean(ean_str):
    """
    Classifica o EAN retornando uma tupla (tipo_codigo, unidade).
    Regras:
      - Se tinha "/" originalmente → Interno, Quilograma
      - Se todos os códigos < 12 dígitos → Interno, Quilograma
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

    if had_slash or all(l < 12 for l in lens):
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
    if all(l < 12 for l in lens):
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

    # Verificar se há linhas válidas após o processamento
    if df_copy.empty:
        st.warning(f"Nenhuma linha válida encontrada para o perfil {profile}. O arquivo não será gerado.")
        return None

    # Substituir "/" por ";" na coluna ean
    df_copy['ean'] = df_copy['ean'].astype(str).str.replace("/", ";", regex=False)

    # Aplicar função única para tipo de código e unidade
    df_copy[['final_code_type', 'final_unit']] = df_copy['ean'].apply(
        lambda x: pd.Series(classify_ean(x))
    )

    # Lista de possíveis nomes de coluna
    possible_buyer_names = ['comprador', 'compradora', 'compradores', 'compradoras']

    # Verifica qual delas existe no DataFrame
    col_name = next((col for col in possible_buyer_names if col in df_copy.columns), None)

    if col_name:
        df_copy['comprador_normalized'] = df_copy[col_name].apply(normalize_text)
    else:
        df_copy['comprador_normalized'] = ''
        st.warning("Nenhuma coluna de comprador encontrada. 'Carrossel' ficará vazio.")

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
    """Mescla os EANs do arquivo com a tabela base usando a coluna CÓDIGO, combinando EANs do encarte e do arquivo."""
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
        
        # Aplicar correção para códigos interpretados como datas
        if 'código' in df_ean.columns:
            df_ean['código'] = df_ean['código'].apply(fix_if_date)
        if 'ean' in df_ean.columns:
            df_ean['ean'] = df_ean['ean'].apply(fix_if_date)
        
        # Normalizar a coluna 'código' em ambos os DataFrames, removendo hífens
        df_base['código'] = df_base['código'].astype(str).str.strip().str.replace('-', '')
        df_ean['código'] = df_ean['código'].astype(str).str.strip().str.replace('-', '')
        
        # Criar uma lista para armazenar as novas linhas
        expanded_rows = []
        for _, row in df_base.iterrows():
            new_row = row.copy()
            # Obter o EAN do encarte consolidado
            encarte_ean = str(new_row['ean']).strip() if not pd.isna(new_row['ean']) else ""
            # Normalizar EAN do encarte, substituindo '/' por ';'
            encarte_ean = encarte_ean.replace('/', ';')
            # Converter EAN do encarte em lista
            encarte_ean_list = [e.strip() for e in encarte_ean.split(';') if e.strip() and e != 'nan']
            
            # Encontrar EANs correspondentes no arquivo de EANs
            matching_eans = df_ean[df_ean['código'] == new_row['código']]['ean']
            
            if not matching_eans.empty:
                # Converter EANs do arquivo para uma lista, removendo valores inválidos
                ean_list = []
                for ean in matching_eans:
                    if not pd.isna(ean) and str(ean).strip() and str(ean).strip() != 'nan':
                        # Dividir EANs do arquivo por ';' ou '/' e normalizar
                        eans = str(ean).strip().replace('/', ';').split(';')
                        ean_list.extend([e.strip() for e in eans if e.strip()])
                # Adicionar EANs do encarte à lista
                ean_list.extend(encarte_ean_list)
                # Remover duplicatas e concatenar com ';'
                ean_list = list(dict.fromkeys(ean_list))  # Remove duplicatas mantendo a ordem
                new_row['ean'] = ';'.join(ean_list) if ean_list else encarte_ean
            else:
                # Se não houver EANs no arquivo, manter o EAN do encarte (já normalizado)
                new_row['ean'] = ';'.join(encarte_ean_list) if encarte_ean_list else encarte_ean
            
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
    
    # Aplicar correção para códigos interpretados como datas (após normalizar colunas)
    if 'código' in df_base.columns:
        df_base['código'] = df_base['código'].apply(fix_if_date)
    if 'ean' in df_base.columns:
        df_base['ean'] = df_base['ean'].apply(fix_if_date)
        
    # Inicializar colunas de preços limpos e marcadores de cópia
    df_base["preço de:"] = df_base["preço de:"].apply(clean_price_value)
    df_base["preço por:"] = df_base["preço por:"].apply(clean_price_value)
    df_base["copied_preço_de"] = False
    df_base["copied_preço_por"] = False

    # Processar preços vazios com base nos 7 primeiros dígitos do EAN
    for i in range(len(df_base)):
        if pd.isna(df_base.iloc[i]["preço de:"]) and i > 0:
            # Obter EANs da linha atual e anterior
            current_ean = str(df_base.iloc[i]["ean"]) if not pd.isna(df_base.iloc[i]["ean"]) else ""
            prev_ean = str(df_base.iloc[i-1]["ean"]) if not pd.isna(df_base.iloc[i-1]["ean"]) else ""
            # Comparar os 7 primeiros dígitos
            if current_ean[:7] == prev_ean[:7] and not pd.isna(df_base.iloc[i-1]["preço de:"]):
                df_base.iloc[i, df_base.columns.get_loc("preço de:")] = df_base.iloc[i-1]["preço de:"]
                df_base.iloc[i, df_base.columns.get_loc("copied_preço_de")] = True
        
        if pd.isna(df_base.iloc[i]["preço por:"]) and i > 0:
            current_ean = str(df_base.iloc[i]["ean"]) if not pd.isna(df_base.iloc[i]["ean"]) else ""
            prev_ean = str(df_base.iloc[i-1]["ean"]) if not pd.isna(df_base.iloc[i-1]["ean"]) else ""
            if current_ean[:7] == prev_ean[:7] and not pd.isna(df_base.iloc[i-1]["preço por:"]):
                df_base.iloc[i, df_base.columns.get_loc("preço por:")] = df_base.iloc[i-1]["preço por:"]
                df_base.iloc[i, df_base.columns.get_loc("copied_preço_por")] = True

    # Mesclar com arquivo de EANs, se fornecido e habilitado
    if use_ean_file and ean_file:
        df_base = merge_ean_data(df_base, ean_file)

    # Preencher valores mesclados (exceto preços)
    df_base['perfil de loja'] = df_base['perfil de loja'].ffill()
    df_base['tipo ação'] = df_base['tipo ação'].ffill()
    # Filtrar linhas com "CRM" no 'tipo ação'
    df_filtered = df_base[df_base["tipo ação"].str.contains("CRM", case=False, na=False)]

    # Garante pasta temporária existe
    os.makedirs(temp_dir, exist_ok=True)
    output_files = []

    for profile in profiles:
        df_profile = df_filtered[df_filtered["perfil de loja"] == profile].copy()

         # ⚠️ se não tem dados desse perfil, pula
        if df_profile.empty:
            st.warning(f"Nenhuma linha encontrada para o perfil {profile}. Pulando geração.")
            continue
        
        # Montar DataFrame final
        df_final = build_final_dataframe(df_profile, profile, start_date, end_date, store_mapping, apply_name_correction)
        
         # ⚠️ se não voltou nada ou ficou vazio, pula
        if df_final is None or df_final.empty:
            st.warning(f"O DataFrame final do perfil {profile} está vazio. Pulando exportação.")
            continue

        # Salvar arquivo Excel com formatação condicional
        filename = f"promo_{profile.replace('/', '_')}_CRM.xlsx"
        filepath = os.path.join(temp_dir, filename)
        filepath = get_unique_filename(filepath)

        # Salvar o DataFrame diretamente no arquivo
        df_final.to_excel(filepath, index=False, engine="openpyxl")

        # Carregar o arquivo Excel com openpyxl para aplicar formatação
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active

        # Definir preenchimentos
        yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

        # Identificar índices das colunas
        header_values = [cell.value for cell in ws[1]]
        preco_col = header_values.index("Preço") + 1
        preco_promo_col = header_values.index("Preço promocional") + 1
        unidade_col = header_values.index("Unidade") + 1
        tipo_codigo_col = header_values.index("Tipo do código") + 1

        # Iterar nas linhas de dados
        for row_idx in range(2, ws.max_row + 1):
            preco_cell = ws.cell(row=row_idx, column=preco_col)
            preco_promo_cell = ws.cell(row=row_idx, column=preco_promo_col)
            unidade_cell = ws.cell(row=row_idx, column=unidade_col)
            tipo_codigo_cell = ws.cell(row=row_idx, column=tipo_codigo_col)

            # 1) Se preço ou preço promocional está vazio -> vermelho
            if preco_cell.value is None or str(preco_cell.value).strip() in ("", "nan") or (
                isinstance(preco_cell.value, float) and math.isnan(preco_cell.value)
            ):
                preco_cell.fill = red_fill

            if preco_promo_cell.value is None or str(preco_promo_cell.value).strip() in ("", "nan") or (
                isinstance(preco_promo_cell.value, float) and math.isnan(preco_promo_cell.value)
            ):
                preco_promo_cell.fill = red_fill

            # 2) Se unidade = QUILOGRAMA -> amarelo
            if str(unidade_cell.value).strip().upper() == "QUILOGRAMA":
                unidade_cell.fill = yellow_fill

            # 3) Se tipo código = INTERNO -> amarelo
            if str(tipo_codigo_cell.value).strip().upper() == "INTERNO":
                tipo_codigo_cell.fill = yellow_fill

        # Salvar o arquivo formatado
        wb.save(filepath)

        # Ler o arquivo para o buffer de download
        with open(filepath, "rb") as f:
            output = BytesIO(f.read())
            output.seek(0)  # 🔴 importante

        # Guardar na lista de saída
        output_files.append((filename, output))
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
                for filename, output in output_files:
                    st.download_button(
                        label=f"Baixar {filename}",
                        data=output,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )