import streamlit as st
import pdfplumber
import io
# Importação dos Mappers específicos
from services.fatura_mapper import extrair_fatura as extrair_B
from services.fatura_mapperA import extrair_fatura as extrair_A
# Importação dos Writers específicos
from services.excel_writer import preparar_planilha as prep_B, salvar_dados_multiplos as salvar_B
from services.excel_writterA import preparar_planilha as prep_A, salvar_dados_A as salvar_A

st.set_page_config(page_title="Balanço Multi-UC", layout="wide")

st.title("⚡ Sistema de Balanço Energético")
st.subheader("Essencial Energia Eficiente")

# --- BARRA LATERAL PARA CONFIGURAÇÃO ---
with st.sidebar:
    st.header("⚙️ Configuração")
    
    # 1. Input crucial: Define qual lógica de código o sistema seguirá
    grupo_selecionado = st.radio(
        "Selecione o Grupo Tarifário:", 
        ["A", "B"], 
        help="Grupo A: Alta Tensão (Demanda e Postos Tarifários). Grupo B: Baixa Tensão (Consumo Único)."
    )
    
    st.markdown("---")
    qtd_geradoras = st.number_input("Qtd. de UC Geradoras", min_value=1, value=1, step=1)
    qtd_beneficiarias = st.number_input("Qtd. de UC Beneficiárias", min_value=0, value=0, step=1)

# --- 2. UPLOAD DA PLANILHA BASE ---
st.subheader("1. Planilha Modelo")
tipo_template = "BALANÇO_A.xlsx" if grupo_selecionado == "A" else "BALANÇO_B.xlsx"
arquivo_excel = st.file_uploader(f"Envie o arquivo Excel para o Grupo {grupo_selecionado}", type=["xlsx"])

if arquivo_excel:
    dados_processamento = []
    
    # --- 3. UPLOAD DAS FATURAS ---
    st.subheader(f"2. Upload das Faturas (Grupo {grupo_selecionado})")
    
    abas_titulos = [f"Geradora {i+1}" for i in range(qtd_geradoras)] + \
                   [f"Beneficiária {i+1}" for i in range(qtd_beneficiarias)]
    tabs = st.tabs(abas_titulos)
    
    idx_tab = 0
    # Interface dinâmica para Geradoras
    for i in range(qtd_geradoras):
        with tabs[idx_tab]:
            pdfs = st.file_uploader(f"Faturas - Geradora {i+1}", type=["pdf"], accept_multiple_files=True, key=f"ger_{i}")
            if pdfs:
                dados_processamento.append({'tipo': 'geradora', 'indice': i + 1, 'arquivos': pdfs})
            idx_tab += 1

    # Interface dinâmica para Beneficiárias
    for i in range(qtd_beneficiarias):
        with tabs[idx_tab]:
            pdfs = st.file_uploader(f"Faturas - Beneficiária {i+1}", type=["pdf"], accept_multiple_files=True, key=f"ben_{i}")
            if pdfs:
                dados_processamento.append({'tipo': 'beneficiaria', 'indice': i + 1, 'arquivos': pdfs})
            idx_tab += 1

    # --- 4. PROCESSAMENTO ---
    st.markdown("---")
    if st.button(f"🚀 Processar Balanço Grupo {grupo_selecionado}"):
        if not dados_processamento:
            st.warning("Envie PDFs para pelo menos uma UC.")
        else:
            progresso = st.progress(0)
            status = st.empty()
            lista_dados_finais = []
            
            # ESCOLHA DO MAPPER (LÓGICA DE LEITURA)
            mapper_func = extrair_A if grupo_selecionado == "A" else extrair_B
            
            # Fase 1: Extração
            for idx, item in enumerate(dados_processamento):
                status.text(f"Lendo faturas da {item['tipo']} {item['indice']}...")
                faturas_extraidas = []
                
                for pdf_file in item['arquivos']:
                    with pdfplumber.open(pdf_file) as pdf:
                        texto = "".join([p.extract_text() or "" for p in pdf.pages])
                    
                    # Chama o mapper correto baseado na seleção do radio button
                    faturas_extraidas.append(mapper_func(texto))
                
                lista_dados_finais.append({
                    'tipo': item['tipo'],
                    'indice': item['indice'],
                    'dados': faturas_extraidas
                })
                progresso.progress((idx + 1) / (len(dados_processamento) + 1))

            # Fase 2: Escrita no Excel (LÓGICA DE GRAVAÇÃO)
            status.text("Gravando dados no Excel...")
            try:
                if grupo_selecionado == "A":
                    # Usa o Writer exclusivo para Alta Tensão (Colunas B, C, D, L, M, N)
                    wb = prep_A(arquivo_excel, qtd_geradoras, qtd_beneficiarias)
                    wb_final = salvar_A(wb, lista_dados_finais)
                else:
                    # Usa o Writer original para Baixa Tensão (Consumo Único)
                    wb = prep_B(arquivo_excel, qtd_geradoras, qtd_beneficiarias)
                    wb_final = salvar_B(wb, lista_dados_finais)
                
                # Download em memória
                output = io.BytesIO()
                wb_final.save(output)
                
                progresso.progress(1.0)
                status.success(f"Planilha Grupo {grupo_selecionado} concluída!")
                
                st.download_button(
                    label="📥 Baixar Resultado Final",
                    data=output.getvalue(),
                    file_name=f"BALANCO_COMPENSAÇÃO_GRUPO_{grupo_selecionado}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Erro no processamento do Excel: {e}")