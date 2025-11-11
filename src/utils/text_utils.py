import string
import unicodedata
import re
import pandas as pd

translator = str.maketrans('', '', string.punctuation)

def normalize_text(text):
    """Normaliza texto: lowercase, remove acentos e pontuação."""
    if pd.isna(text):
        return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    text = text.translate(translator)
    return text

def remove_suffix(text):
    """Remove sufixos como _sell out, _faturamento e tudo que vier depois."""
    if pd.isna(text):
        return ""
    keywords = ['sell out', 'faturamento', 'sell in']
    pattern = r'_(' + '|'.join(keywords) + r').*$'
    return re.sub(pattern, '', str(text), flags=re.IGNORECASE).strip()

def correct_product_name(name, corrections_dict):
    """Corrige o nome do produto com base no dicionário de correções, retornando em maiúsculo."""
    if pd.isna(name):
        return ""
    corrected_name = str(name).strip()
    for pattern, replacement in corrections_dict.items():
        corrected_name = re.sub(pattern, replacement, corrected_name, flags=re.IGNORECASE)
    return corrected_name.upper()