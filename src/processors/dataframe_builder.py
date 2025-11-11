import pandas as pd
import streamlit as st
from src.utils.text_utils import normalize_text, remove_suffix, correct_product_name
from src.utils.ean_classifier import classify_ean

def get_carrossel_value(normalized_buyer, mapping):
    """Retorna o valor do carrossel se chave estiver contida no texto do comprador."""
    if not normalized_buyer:
        return ''
    for key, value in mapping.items():
        if key in normalized_buyer:
            return value
    return ''

def build_final_dataframe(filtered_df, profile, start_date, end_date, store_map, apply_name_correction, link_map, buyer_carrossel_map, product_name_corrections):
    """Constrói o DataFrame final para exportação"""
    df_copy = filtered_df.copy()
    df_copy['descrição do item'] = df_copy['descrição do item'].apply(remove_suffix)

    if apply_name_correction:
        df_copy['descrição do item'] = df_copy['descrição do item'].apply(
            lambda x: correct_product_name(x, product_name_corrections)
        )
    else:
        df_copy['descrição do item'] = df_copy['descrição do item'].apply(
            lambda x: str(x).strip().upper() if not pd.isna(x) else ""
        )

    if df_copy.empty:
        st.warning(f"Nenhuma linha válida encontrada para o perfil {profile}. O arquivo não será gerado.")
        return None

    df_copy['ean'] = df_copy['ean'].astype(str).str.replace("/", ";", regex=False)
    df_copy[['final_code_type', 'final_unit']] = df_copy['ean_original_encarte'].apply(
        lambda x: pd.Series(classify_ean(x))
    )

    possible_buyer_names = ['comprador', 'compradora', 'compradores', 'compradoras']
    col_name = next((col for col in possible_buyer_names if col in df_copy.columns), None)

    if col_name:
        df_copy['comprador_normalized'] = df_copy[col_name].apply(normalize_text)
    else:
        df_copy['comprador_normalized'] = ''
        st.warning("Nenhuma coluna de comprador encontrada. 'Carrossel' ficará vazio.")

    # Aplica "8142 - Especial" para produtos com "DESTAQUE CRM" em "tipo ação"
    df_copy['final_carrossel'] = df_copy.apply(
        lambda row: "8142 - Especial" if "DESTAQUE CRM" in str(row['tipo ação']).upper() else get_carrossel_value(row['comprador_normalized'], buyer_carrossel_map),
        axis=1
    )

    result_df = pd.DataFrame({
        "Nome": df_copy["descrição do item"],
        "Carrossel": df_copy["final_carrossel"],
        "Check-In": "Não",
        "Preço": df_copy["preço de:"],
        "Preço promocional": df_copy["preço por:"],
        "Limite por cliente": 0,
        "Dias para Resgate após ativação": (end_date.date() - start_date.date()).days + 1,
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

    # Preencher URLs automaticamente com base no JSON de links
    urls = []
    for ean_field in df_copy["ean"]:
        link = ""
        if pd.notna(ean_field):
            for e in str(ean_field).replace("/", ";").split(";"):
                e = e.strip()
                if e and e in link_map:
                    link = link_map[e]
                    break
        urls.append(link)
    result_df["URL da imagem"] = urls

    return result_df