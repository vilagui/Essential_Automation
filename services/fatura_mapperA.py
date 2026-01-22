import re

def normalizar_numero_br(valor: str) -> float:
    """Converte strings no formato brasileiro (1.234,56) para float (1234.56)."""
    if not valor:
        return 0.0
    valor = valor.replace(".", "").replace(",", ".")
    try:
        return float(valor)
    except ValueError:
        return 0.0

def normalizar_texto(texto: str) -> str:
    """Padroniza o texto em maiúsculas e remove espaços extras."""
    return " ".join(texto.upper().split())

def extrair_historico_consumo(texto: str) -> list:
    """
    Captura as colunas da tabela de histórico: Mês/Ano, Demanda (P, FP, RE), 
    Consumo Faturado (P, FP, RE) e Horário Reservado (Consumo).
    """
    historico = []
    # Regex flexível para capturar a linha da tabela (Mês/Ano + sequência de valores) 
    padrao = r"(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\s*[\/\-]\s*(\d{2})((?:\s+[\d\.,]+){7,9})"
    
    matches = re.findall(padrao, texto)
    
    for mes, ano, valores_str in matches:
        v = valores_str.strip().split()
        if len(v) >= 7:
            # Mapeamento conforme a tabela de histórico de 9 colunas 
            historico.append({
                "mes": mes,
                "ano": ano,
                "demanda_p": normalizar_numero_br(v[0]),
                "demanda_fp": normalizar_numero_br(v[1]),
                "demanda_hr": normalizar_numero_br(v[2]), # Demanda Reativo Excedente/Reservado
                "consumo_p": normalizar_numero_br(v[3]),
                "consumo_fp": normalizar_numero_br(v[4]),
                "consumo_hr": normalizar_numero_br(v[6])  # Consumo Horário Reservado
            })
    return historico

def extrair_fatura(texto: str) -> dict:
    """Realiza a extração completa dos dados da fatura para o Grupo A."""
    dados = {}
    texto_norm = normalizar_texto(texto)

    # --- 1. MÊS E UC ---
    m_uc_mes = re.search(r"(\d{7,})\s+(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)/(\d{4})", texto_norm)
    dados["uc"] = m_uc_mes.group(1) if m_uc_mes else ""
    dados["mes"] = m_uc_mes.group(2) if m_uc_mes else ""
    dados["ano"] = m_uc_mes.group(3)[2:] if m_uc_mes else "00" # Pega os dois últimos dígitos do ano

    # --- 2. CONSUMO E DEMANDA ATUAL (PONTA, FORA PONTA E RESERVADO) ---
    pats = {
        "c_p": r"ENERGIA ATIVA - KWH PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)",
        "c_fp": r"ENERGIA ATIVA - KWH FORA PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)",
        "c_hr": r"ENERGIA ATIVA - KWH RESERVADO\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)",
        "d_p": r"DEMANDA KW PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)",
        "d_fp": r"DEMANDA KW FORA PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)",
        "d_hr": r"DEMANDA KW RESERVADO\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)"
    }
    for chave, pat in pats.items():
        m = re.search(pat, texto_norm)
        dados[chave] = normalizar_numero_br(m.group(1)) if m else 0.0

    # --- 3. SALDO TOTAL (SOMA P + FP + HR) ---
    # Soma todos os componentes do saldo encontrados no bloco SCEE 
    dados["saldo"] = 0.0
    idx_scee = texto_norm.find("INFORMAÇÕES DO SCEE")
    if idx_scee != -1:
        bloco = texto_norm[idx_scee : idx_scee + 800]
        # Captura o trecho específico do saldo para evitar pegar outros valores do bloco 
        m_saldo = re.search(r"SALDO KWH.*?(?=SALDO A EXPIRAR|TOTAL|$)", bloco)
        if m_saldo:
            valores_encontrados = re.findall(r"[\d\.]*,\d{2}", m_saldo.group(0))
            dados["saldo"] = sum(normalizar_numero_br(v) for v in valores_encontrados)

    # --- 4. HISTÓRICO COMPLETO ---
    dados["historico"] = extrair_historico_consumo(texto_norm)
    
    # --- 5. VALOR TOTAL DA FATURA ---
    m_val = re.search(r"TOTAL\s+([\d\.]+,\d{2})", texto_norm)
    dados["valor_fatura"] = normalizar_numero_br(m_val.group(1)) if m_val else 0.0

    return dados