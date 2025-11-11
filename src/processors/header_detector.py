import pandas as pd
from src.utils.text_utils import normalize_text

def detect_header_with_scoring(df, required_columns):
    """
    Sistema de pontuação para descobrir qual linha é o verdadeiro header.
    Retorna a linha que será usada como header e lista de erros caso alguma coluna obrigatória não seja encontrada.
    """
    max_score = -1
    header_row = None
    row_found = None

    # Limitar iteração às primeiras 20 linhas
    limited_df = df.head(20)
    
    for idx, row in limited_df.iterrows():
        score = 0
        normalized_row = [normalize_text(str(cell)) for cell in row if not pd.isna(cell)]

        for col in required_columns:
            if normalize_text(col) in normalized_row:
                score += 1

        if score > max_score:
            max_score = score
            header_row = idx
            row_found = normalized_row

    # Verificar se nenhuma linha válida foi encontrada
    if header_row is None or not row_found:
        errors = ["❌ Nenhuma linha válida encontrada com colunas obrigatórias nas primeiras 20 linhas. Verifique o arquivo."]
        return None, errors

    # Verificar se todas as obrigatórias estão presentes
    missing_cols = [col for col in required_columns if normalize_text(col) not in row_found]
    if missing_cols:
        errors = [
            f"❌ Coluna obrigatória '{col}' não encontrada na linha {header_row + 1} do arquivo do Encarte Consolidado. Verifique a digitação do título da coluna."
            for col in missing_cols
        ]
        return None, errors

    return header_row, []