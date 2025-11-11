import os
import pandas as pd
import streamlit as st

def get_unique_filename(path):
    """Recebe um caminho de arquivo e retorna um nome único no mesmo diretório."""
    base, ext = os.path.splitext(path)
    counter = 1
    new_path = path
    while os.path.exists(new_path):
        new_path = f"{base} ({counter}){ext}"
        counter += 1
    return new_path

def list_sheets(uploaded_file):
    """Lista planilhas disponíveis em um arquivo Excel ou retorna opção padrão para CSV"""
    file_extension = os.path.splitext(uploaded_file.name)[1].lower()
    try:
        if file_extension in ['.xlsx', '.xls']:
            xl = pd.ExcelFile(uploaded_file)
            return xl.sheet_names
        elif file_extension == '.csv':
            return ["Planilha CSV"]
        else:
            st.error("Formato de arquivo não suportado. Use xlsx, xls ou csv.")
            return []
    except Exception as e:
        st.error(f"Erro ao listar planilhas: {e}")
        return []