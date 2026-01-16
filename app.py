import streamlit as st
import pdfplumber

from services.fatura_mapper import extrair_fatura
from services.excel_writer import escrever_uc_geradora

st.set_page_config(
    page_title="Balanço Energético – Equatorial",
    layout="wide"
)

st.title("🔋 Balanço Energético – Equatorial")

arquivo_pdf = st.file_uploader(
    "Envie a fatura PDF da UC Geradora",
    type=["pdf"]
)

if arquivo_pdf:
    texto = ""
    with pdfplumber.open(arquivo_pdf) as pdf:
        for page in pdf.pages:
            texto += page.extract_text() or ""

    st.subheader("🧪 TEXTO BRUTO EXTRAÍDO DO PDF")
    st.text_area("Conteúdo completo", texto, height=400)

    dados = extrair_fatura(texto)
    st.json(dados)


    if st.button("💾 Gravar no Excel"):
        escrever_uc_geradora("BALANÇO E COMPENSAÇÃO.xlsx", dados)
        st.success("Dados gravados com sucesso no Excel.")

