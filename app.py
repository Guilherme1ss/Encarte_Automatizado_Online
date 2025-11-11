import streamlit as st
from datetime import datetime, timedelta
import tempfile
from io import BytesIO

from src.processors.promotion_processor import process_promotions
from src.utils.file_utils import list_sheets

st.title("Processador de Promoções CRM")
st.write("Faça upload da planilha de promoções (xlsx, xls ou csv) e, opcionalmente, um arquivo com EANs (xlsx, xls ou csv). Selecione as datas do encarte e a planilha desejada.")

temp_dir = tempfile.mkdtemp()
default_start = datetime.today()
default_end = datetime.today() + timedelta(days=7)

col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Data de Início do Encarte", value=default_start, format="DD/MM/YYYY")
with col2:
    end_date = st.date_input("Data de Fim do Encarte", value=default_end, format="DD/MM/YYYY")

if end_date < start_date:
    st.error("A data de fim não pode ser anterior à data de início.")
else:
    apply_name_correction = st.checkbox("Aplicar correção de nomes de produtos", value=False)
    use_ean_file = st.checkbox("Usar arquivo de EANs", value=False)
    use_link_file = st.checkbox("Usar arquivo JSON de Links", value=False)
    use_default_url = False
    link_file = None
    
    if use_link_file:
        st.write("Escolha a fonte do arquivo de links:")
        link_source = st.radio("Fonte do arquivo JSON", ["Usar repositório de links padrão", "Fazer upload de um arquivo JSON"])
        if link_source == "Usar repositório de links padrão":
            use_default_url = True
        else:
            link_file = st.file_uploader("Selecione o arquivo JSON de Links", type=["json"])

    uploaded_file = st.file_uploader("Selecione o arquivo de ENCARTE CONSOLIDADO", type=["xlsx", "xls", "csv"])
    
    selected_sheet = None
    if uploaded_file:
        sheet_names = list_sheets(uploaded_file)
        if sheet_names:
            st.write("Selecione a planilha para processar:")
            selected_sheet = st.selectbox("Planilhas disponíveis", sheet_names)
        else:
            st.error("Nenhuma planilha encontrada no arquivo.")
    
    ean_file = None
    if use_ean_file:
        ean_file = st.file_uploader("Selecione o arquivo de EANs (opcional)", type=["xlsx", "xls", "csv"])

    if st.button("Processar Promoções"):
        output_files = []
        try:
            with st.spinner("Processando..."):
                start_dt = datetime.combine(start_date, datetime.min.time())
                end_dt = datetime.combine(end_date, datetime.max.time().replace(second=0))
                output_files = process_promotions(
                    uploaded_file, ean_file, link_file, use_default_url,
                    start_dt, end_dt, temp_dir,
                    use_ean_file, use_link_file, apply_name_correction, selected_sheet
                )
        except Exception as e:
            st.error(f"Erro durante o processamento: {e}")

        if output_files:
            for filename, output in output_files:
                st.download_button(
                    label=f"Baixar {filename}",
                    data=output,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("Nenhum arquivo foi gerado. Verifique os dados de entrada.")