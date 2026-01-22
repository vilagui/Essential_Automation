import streamlit as st
import pdfplumber
import io
import os
# Importação dos mappers específicos para cada grupo
from services.fatura_mapper import extrair_fatura as extrair_B
from services.fatura_mapperA import extrair_fatura as extrair_A
from services.excel_writer import preparar_planilha, salvar_dados_multiplos

st.set_page_config(page_title="Balanço Equatorial Multi-UC", layout="wide")

st.title("⚡ Sistema de Balanço Energético (Multi-UC)")
st.subheader("Essencial Energia Eficiente")

# --- BARRA LATERAL PARA CONFIGURAÇÃO ---
with st.sidebar:
    st.header("⚙️ Configuração")
    
    # Decisor de lógica: Grupo A (Alta Tensão) ou Grupo B (Baixa Tensão)
    grupo_selecionado = st.radio(
        "Selecione o Grupo Tarifário:", 
        ["A", "B"], 
        help="Grupo A: Alta Tensão (Demanda/Ponta). Grupo B: Baixa Tensão (Residencial/Comercial Simples)."
    )
    
    st.markdown("---")
    qtd_geradoras = st.number_input("Qtd. de UC Geradoras", min_value=1, value=1, step=1)
    qtd_beneficiarias = st.number_input("Qtd. de UC Beneficiárias", min_value=0, value=0, step=1)

# --- 1. UPLOAD DA PLANILHA BASE ---
st.subheader("1. Planilha Modelo")
arquivo_excel = st.file_uploader("Envie o arquivo 'BALANÇO E COMPENSAÇÃO.xlsx'", type=["xlsx"])

if arquivo_excel:
    dados_processamento = []
    st.subheader(f"2. Upload das Faturas (Grupo {grupo_selecionado})")
    
    # Criação dinâmica das abas conforme a quantidade configurada
    abas_titulos = [f"Geradora {i+1}" for i in range(qtd_geradoras)] + \
                   [f"Beneficiária {i+1}" for i in range(qtd_beneficiarias)]
    tabs = st.tabs(abas_titulos)
    
    idx_tab = 0
    # --- Interface de Upload para Geradoras ---
    for i in range(qtd_geradoras):
        with tabs[idx_tab]:
            st.markdown(f"**UC Geradora {i+1}**")
            pdfs = st.file_uploader(f"Faturas PDF - Geradora {i+1}", type=["pdf"], accept_multiple_files=True, key=f"ger_{i}")
            if pdfs:
                dados_processamento.append({'tipo': 'geradora', 'indice': i + 1, 'arquivos': pdfs})
            idx_tab += 1

    # --- Interface de Upload para Beneficiárias ---
    for i in range(qtd_beneficiarias):
        with tabs[idx_tab]:
            st.markdown(f"**UC Beneficiária {i+1}**")
            pdfs = st.file_uploader(f"Faturas PDF - Beneficiária {i+1}", type=["pdf"], accept_multiple_files=True, key=f"ben_{i}")
            if pdfs:
                dados_processamento.append({'tipo': 'beneficiaria', 'indice': i + 1, 'arquivos': pdfs})
            idx_tab += 1

    # --- BOTÃO DE PROCESSAMENTO ---
    st.markdown("---")
    if st.button(f"🚀 Processar Todas as UCs - Grupo {grupo_selecionado}"):
        if not dados_processamento:
            st.warning("Por favor, envie faturas para pelo menos uma UC.")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
            lista_dados_finais = []
            
            # Escolha automática do mapper baseado no grupo selecionado
            mapper_func = extrair_A if grupo_selecionado == "A" else extrair_B
            
            # Passo 1: Extração de Dados
            total_uc = len(dados_processamento)
            for idx, item in enumerate(dados_processamento):
                status_text.text(f"Extraindo dados da {item['tipo']} {item['indice']}...")
                faturas_extraidas = []
                
                for pdf_file in item['arquivos']:
                    with pdfplumber.open(pdf_file) as pdf:
                        texto = "".join([page.extract_text() or "" for page in pdf.pages])
                    
                    dados = mapper_func(texto)
                    faturas_extraidas.append(dados)
                
                lista_dados_finais.append({
                    'tipo': item['tipo'],
                    'indice': item['indice'],
                    'dados': faturas_extraidas
                })
                progress_bar.progress((idx + 0.5) / total_uc)
            
            # Passo 2: Escrita no Excel
            try:
                status_text.text("Gerando abas e escrevendo na planilha...")
                wb_preparado = preparar_planilha(arquivo_excel, qtd_geradoras, qtd_beneficiarias)
                
                # Envia o grupo_selecionado para que o writer use as colunas corretas (A ou B)
                wb_final = salvar_dados_multiplos(wb_preparado, lista_dados_finais, grupo_selecionado)
                
                # Salva em memória para disponibilizar o download
                output = io.BytesIO()
                wb_final.save(output)
                output.seek(0)
                
                progress_bar.progress(1.0)
                status_text.text("Processamento concluído!")
                st.success(f"Planilha do Grupo {grupo_selecionado} gerada com sucesso!")
                
                st.download_button(
                    label="📥 Baixar Planilha Consolidada",
                    data=output.getvalue(),
                    file_name=f"BALANCO_ENERGETICO_GRUPO_{grupo_selecionado}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Ocorreu um erro ao processar o Excel: {e}")