import re

def normalizar_numero_br(valor: str) -> float:
    """Converte padrão brasileiro de números para float."""
    if not valor: return 0.0
    valor = valor.replace(".", "").replace(",", ".")
    try: return float(valor)
    except ValueError: return 0.0

def normalizar_texto(texto: str) -> str:
    """Limpa o texto e remove espaços extras."""
    return " ".join(texto.upper().split())

def extrair_historico_consumo(texto: str) -> list:
    """
    Extrai a tabela de histórico de 9 colunas do Grupo A.
    Mapeia: Demanda (P, FP, RE) e Consumo (P, FP, RE, Reservado).
    """
    historico = []
    # Regex para capturar Mês/Ano e a sequência de valores numéricos
    padrao = r"(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\s*[\/\-]\s*(\d{2})((?:\s+[\d\.,]+){7,9})"
    matches = re.findall(padrao, texto)
    
    for mes, ano, valores_str in matches:
        v = valores_str.strip().split()
        if len(v) >= 7:
            historico.append({
                "mes": mes, 
                "ano": ano,
                # Chaves padronizadas para o excel_writer_A
                "d_p": normalizar_numero_br(v[0]), 
                "d_fp": normalizar_numero_br(v[1]),
                "d_hr": normalizar_numero_br(v[2]), # Demanda Reservada/Reativa
                "c_p": normalizar_numero_br(v[3]), 
                "c_fp": normalizar_numero_br(v[4]),
                "c_hr": normalizar_numero_br(v[6])  # Consumo Horário Reservado
            })
    return historico

def extrair_fatura(texto: str) -> dict:
    """Função principal de extração para faturas de Alta Tensão (Grupo A)."""
    dados = {}
    t_norm = normalizar_texto(texto)
    
    # --- 1. MÊS E ANO ATUAL ---
    m = re.search(r"(\d{7,})\s+(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)/(\d{4})", texto)
    dados["uc"] = m.group(1) if m else ""
    dados["mes"] = m.group(2) if m else ""
    dados["ano"] = int(m.group(3)) if m else 0

    # --- 2. ENDEREÇO ---
    if "ENDEREÇO DE ENTREGA:" in texto:
        trecho = texto.split("ENDEREÇO DE ENTREGA:", 1)[1]
        dados["endereco"] = trecho.split("CEP:", 1)[0].strip()
    else:
        dados["endereco"] = ""

    # --- 3. DATAS ---
    m = re.search(r"(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+\d+\s+\d{2}/\d{2}/\d{4}", texto)
    dados["data_leitura_anterior"] = m.group(1) if m else ""
    dados["data_leitura_atual"] = m.group(2) if m else ""

    # 4. Consumo e Demanda Atuais (Postos Tarifários)
    pats = {
        "c_p": r"ENERGIA ATIVA - KWH PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)",
        "c_fp": r"ENERGIA ATIVA - KWH FORA PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)",
        "c_hr": r"ENERGIA ATIVA - KWH RESERVADO\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)",
        "d_p": r"DEMANDA - KW PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)",
        "d_fp": r"DEMANDA - KW FORA PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)",
        "d_hr": r"DEMANDA - KW RESERVADO\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)"
    }
    for chave, pat in pats.items():
        m = re.search(pat, t_norm)
        dados[chave] = normalizar_numero_br(m.group(1)) if m else 0.0

    # 5. Saldo SCEE (Soma dos Créditos P + FP + HR)
    dados["saldo"] = 0.0
    if "INFORMAÇÕES DO SCEE" in t_norm:
        bloco = t_norm[t_norm.find("INFORMAÇÕES DO SCEE") : t_norm.find("INFORMAÇÕES DO SCEE")+800]
        m_s = re.search(r"SALDO KWH.*?(?=SALDO A EXPIRAR|TOTAL|$)", bloco)
        if m_s:
            # Soma todos os valores de saldo encontrados no trecho
            dados["saldo"] = sum(normalizar_numero_br(v) for v in re.findall(r"[\d\.]*,\d{2}", m_s.group(0)))

    # 6. Histórico e Valor Total
    dados["historico"] = extrair_historico_consumo(t_norm)
    m_val = re.search(r"TOTAL\s+([\d\.]+,\d{2})", t_norm)
    dados["valor_fatura"] = normalizar_numero_br(m_val.group(1)) if m_val else 0.0
    
    # 7. Geração Distribuída (Crédito)
    m_cred = re.search(r"CRÉDITO RECEBIDO.*?([\d\.]+,\d{2})", t_norm)
    dados["credito_recebido"] = normalizar_numero_br(m_cred.group(1)) if m_cred else 0.0

    return dados