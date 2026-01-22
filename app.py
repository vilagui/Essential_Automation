import streamlit as st
import pdfplumber
import io
from services.fatura_mapper import extrair_fatura as extrair_B
from services.fatura_mapperA import extrair_fatura as extrair_A
from services.excel_writer import preparar_planilha, salvar_dados_multiplos

st.set_page_config(page_title="Balanço Equatorial Multi-UC", layout="wide")
st.title("⚡ Sistema de Balanço Energético (Multi-UC)")

with st.sidebar:
    st.header("Configuração")
    grupo_selecionado = st.radio("Selecione o Grupo Tarifário:", ["A", "B"])
    qtd_geradoras = st.number_input("Qtd. de UC Geradoras", min_value=1, value=1)
    qtd_beneficiarias = st.number_input("Qtd. de UC Beneficiárias", min_value=0, value=0)

arquivo_excel = st.file_uploader("Envie a planilha modelo", type=["xlsx"])

if arquivo_excel:
    dados_processamento = []
    abas = [f"Geradora {i+1}" for i in range(qtd_geradoras)] + [f"Beneficiária {i+1}" for i in range(qtd_beneficiarias)]
    tabs = st.tabs(abas)
    
    # ... (Lógica de upload nas tabs permanece idêntica)

    if st.button(f"🚀 Processar Grupo {grupo_selecionado}"):
        mapper = extrair_A if grupo_selecionado == "A" else extrair_B
        lista_dados_finais = []
        
        for item in dados_processamento:
            faturas = []
            for pdf_file in item['arquivos']:
                with pdfplumber.open(pdf_file) as pdf:
                    texto = "".join([p.extract_text() or "" for p in pdf.pages])
                faturas.append(mapper(texto))
            lista_dados_finais.append({'tipo': item['tipo'], 'indice': item['indice'], 'dados': faturas})
        
        wb = preparar_planilha(arquivo_excel, qtd_geradoras, qtd_beneficiarias)
        wb_final = salvar_dados_multiplos(wb, lista_dados_finais, grupo_selecionado)
        
        output = io.BytesIO()
        wb_final.save(output)
        st.download_button("📥 Baixar Planilha Final", output.getvalue(), file_name=f"BALANCO_GRUPO_{grupo_selecionado}.xlsx")