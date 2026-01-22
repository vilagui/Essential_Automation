import streamlit as st
import pdfplumber
import os
from services.fatura_mapper import extrair_fatura
from services.excel_writer import preparar_planilha, salvar_dados_multiplos

st.set_page_config(page_title="Balanço Equatorial Multi-UC", layout="wide")

st.title("⚡ Sistema de Balanço Energético (Multi-UC)")
st.subheader("Essencial Energia Eficiente")

# --- BARRA LATERAL PARA CONFIGURAÇÃO ---
with st.sidebar:
    st.header("Configuração")
    # NOVO: Seleção de Grupo Tarifário
    grupo_selecionado = st.radio("Selecione o Grupo Tarifário:", ["A", "B"], help="Grupo A: Alta Tensão (Demanda/Ponta). Grupo B: Baixa Tensão.")
    
    st.markdown("---")
    qtd_geradoras = st.number_input("Qtd. de UC Geradoras", min_value=1, value=1, step=1)
    qtd_beneficiarias = st.number_input("Qtd. de UC Beneficiárias", min_value=0, value=0, step=1)

# --- UPLOAD DA PLANILHA BASE ---
st.subheader("1. Planilha Modelo")
arquivo_excel = st.file_uploader("Envie o arquivo 'BALANÇO E COMPENSAÇÃO.xlsx'", type=["xlsx"])

if arquivo_excel:
    dados_processamento = []
    st.subheader(f"2. Upload das Faturas (Grupo {grupo_selecionado})")
    
    abas_titulos = [f"Geradora {i+1}" for i in range(qtd_geradoras)] + \
                   [f"Beneficiária {i+1}" for i in range(qtd_beneficiarias)]
    tabs = st.tabs(abas_titulos)
    
    idx_tab = 0
    # Loop para Geradoras e Beneficiárias (mesma lógica de upload)
    for tipo, qtd in [('geradora', qtd_geradoras), ('beneficiaria', qtd_beneficiarias)]:
        for i in range(qtd):
            with tabs[idx_tab]:
                st.markdown(f"**{tipo.title()} {i+1}**: Envie os PDFs")
                pdfs = st.file_uploader(f"Faturas - {tipo.title()} {i+1}", type=["pdf"], accept_multiple_files=True, key=f"{tipo}_{i}")
                if pdfs:
                    dados_processamento.append({'tipo': tipo, 'indice': i + 1, 'arquivos': pdfs})
                idx_tab += 1

    # --- BOTÃO DE PROCESSAMENTO ---
    st.markdown("---")
    if st.button(f"🚀 Processar Grupo {grupo_selecionado}"):
        if not dados_processamento:
            st.warning("Envie faturas para pelo menos uma UC.")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
            lista_dados_finais = []
            
            # Passo 1: Extração
            total_items = len(dados_processamento)
            for idx, item in enumerate(dados_processamento):
                status_text.text(f"Lendo PDFs da UC {item['indice']}...")
                faturas_extraidas = []
                for pdf_file in item['arquivos']:
                    texto = ""
                    with pdfplumber.open(pdf_file) as pdf:
                        for page in pdf.pages:
                            texto += page.extract_text() or ""
                    
                    # Passamos o grupo para o mapper se necessário
                    dados = extrair_fatura(texto) 
                    faturas_extraidas.append(dados)
                
                lista_dados_finais.append({
                    'tipo': item['tipo'],
                    'indice': item['indice'],
                    'dados': faturas_extraidas
                })
                progress_bar.progress((idx + 0.5) / total_items)
            
            # Passo 2: Escrita no Excel
            status_text.text("Escrevendo na Planilha...")
            wb_preparado = preparar_planilha(arquivo_excel, qtd_geradoras, qtd_beneficiarias)
            
            # NOVO: Passando o grupo_selecionado para a função de escrita
            wb_final = salvar_dados_multiplos(wb_preparado, lista_dados_finais, grupo_selecionado)
            
            nome_saida = f"BALANCO_GRUPO_{grupo_selecionado}.xlsx"
            wb_final.save(nome_saida)
            
            progress_bar.progress(1.0)
            status_text.text("Concluído!")
            st.success(f"Planilha Grupo {grupo_selecionado} gerada!")
            
            with open(nome_saida, "rb") as f:
                st.download_button("📥 Baixar Planilha Final", f, file_name=nome_saida)