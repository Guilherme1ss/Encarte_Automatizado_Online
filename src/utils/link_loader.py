import json
import streamlit as st

def load_links_json(file):
    """Carrega um arquivo JSON com links e retorna um dicionário {EAN: URL}"""
    if not file:
        return {}

    try:
        if isinstance(file, str):  # Caso seja o arquivo padrão
            with open(file, "r", encoding="utf-8") as f:
                data = json.load(f)
        else:  # Caso seja um arquivo enviado
            data = json.load(file)
        ean_to_url = {}
        for item in data:
            url = item.get("url", "").strip()
            if not url:
                continue
            for ean in item.get("eans", []):
                ean_to_url[str(ean).strip()] = url
        return ean_to_url
    except Exception as e:
        st.error(f"Erro ao ler arquivo de links: {e}")
        return {}