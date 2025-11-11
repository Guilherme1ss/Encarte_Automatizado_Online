import json
import streamlit as st

def load_config():
    """Carrega configurações do arquivo JSON"""
    try:
        with open("config.json", "r", encoding="utf-8") as f:
            config = json.load(f)
        return (
            config["required_columns"],
            config["buyer_carrossel_map"],
            config["product_name_corrections"]
        )
    except FileNotFoundError:
        st.error("Arquivo config.json não encontrado. Certifique-se de que ele está no mesmo diretório do script.")
        return None, None, None
    except Exception as e:
        st.error(f"Erro ao carregar config.json: {e}")
        return None, None, None