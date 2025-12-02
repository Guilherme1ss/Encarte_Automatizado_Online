import pandas as pd

def classify_ean(ean_str):
    """
    Classifica o EAN retornando uma tupla (tipo_codigo, unidade).
    Regras:
      - Se todos os códigos < 10 dígitos → Interno, Quilograma
      - Caso contrário → EAN, Unidade
    """
    if not ean_str or pd.isna(ean_str) or not str(ean_str).strip():
        return ("EAN", "Unidade")

    ean_str = str(ean_str).replace("/", ";")
    eans = [e.strip() for e in ean_str.split(';') if e.strip()]
    if not eans:
        return ("EAN", "Unidade")
    
    first_ean = eans[0]
    first_len = len(first_ean)

    if first_len < 10:
        return ("Interno", "Quilograma")
    else:
        return ("EAN", "Unidade")

def get_code_type(ean):
    """Retorna o tipo de código baseado no EAN"""
    if pd.isna(ean) or not str(ean).strip():
        return 'EAN'
    
    ean_str = str(ean).replace("/", ";")
    eans = [e.strip() for e in ean_str.split(';') if e.strip()]
    if not eans:
        return 'EAN'
    
    lens = [len(e) for e in eans]
    if all(l < 10 for l in lens):
        return 'Interno'
    else:
        return 'EAN'