import pandas as pd
from datetime import datetime

def fix_if_date(value):
    """Corrige códigos interpretados como datas"""
    if pd.isna(value):
        return value
    if isinstance(value, (datetime, pd.Timestamp)):
        return f"{value.year}-{value.month}"
    else:
        str_value = str(value)
        if str_value.endswith(".0") and str_value.replace(".", "").isdigit():
            return str_value[:-2]
        return str_value

def clean_price_value(value):
    """Limpa string de preço e converte para float, retornando None se inválido."""
    try:
        text = str(value).replace("R$", "").replace(",", ".").strip()
        return float(text)
    except:
        return None