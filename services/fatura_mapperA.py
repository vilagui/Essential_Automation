import re

def normalizar_numero_br(valor: str) -> float:
    if not valor:
        return 0.0
    valor = valor.replace(".", "").replace(",", ".")
    try:
        return float(valor)
    except ValueError:
        return 0.0

def normalizar_texto(texto: str) -> str:
    return " ".join(texto.upper().split())

def extrair_historico_consumo(texto: str) -> list:
    """
    Extrai as colunas da tabela de histórico: Mês, Demanda (P, FP), 
    Consumo (P, FP) e Horário Reservado[cite: 96].
    """
    historico = []
    # Regex para capturar Mês/Ano e a sequência de valores da tabela [cite: 96]
    padrao = r"(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\s*[\/\-]\s*(\d{2})((?:\s+[\d\.,]+){7,9})"
    matches = re.findall(padrao, texto)
    
    for mes, ano, valores_str in matches:
        v = valores_str.strip().split()
        if len(v) >= 7:
            historico.append({
                "mes": mes,
                "ano": int(ano),
                "demanda_p": normalizar_numero_br(v[0]),
                "demanda_fp": normalizar_numero_br(v[1]),
                "consumo_p": normalizar_numero_br(v[3]),
                "consumo_fp": normalizar_numero_br(v[4]),
                "consumo_hr": normalizar_numero_br(v[6])
            })
    return historico

def extrair_fatura(texto: str) -> dict:
    dados = {}
    texto_norm = normalizar_texto(texto)

    # --- 1. MÊS E UC ---
    m_uc_mes = re.search(r"(\d{7,})\s+(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)/(\d{4})", texto_norm)
    dados["uc"] = m_uc_mes.group(1) if m_uc_mes else ""
    dados["mes"] = m_uc_mes.group(2) if m_uc_mes else ""
    dados["ano"] = int(m_uc_mes.group(3)) if m_uc_mes else 0

    # --- 2. CONSUMO E DEMANDA ATUAL (INDIVIDUAL) ---
    pats = {
        "c_p": r"ENERGIA ATIVA - KWH PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)",
        "c_fp": r"ENERGIA ATIVA - KWH FORA PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)",
        "c_hr": r"ENERGIA ATIVA - KWH RESERVADO\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)",
        "d_p": r"DEMANDA KW PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)",
        "d_fp": r"DEMANDA KW FORA PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)"
    }
    for chave, pat in pats.items():
        m = re.search(pat, texto_norm)
        dados[chave] = normalizar_numero_br(m.group(1)) if m else 0.0

    # --- 3. SALDO TOTAL (Soma P + FP + HR) ---
    # $Saldo_{total} = \sum (P, FP, HR)$ 
    dados["saldo"] = 0.0
    idx_scee = texto_norm.find("INFORMAÇÕES DO SCEE")
    if idx_scee != -1:
        bloco = texto_norm[idx_scee : idx_scee + 800]
        m_saldo = re.search(r"SALDO KWH.*?(?=SALDO A EXPIRAR|TOTAL|$)", bloco)
        if m_saldo:
            valores_encontrados = re.findall(r"[\d\.]*,\d{2}", m_saldo.group(0))
            dados["saldo"] = sum(normalizar_numero_br(v) for v in valores_encontrados)

    # --- 4. HISTÓRICO ---
    dados["historico"] = extrair_historico_consumo(texto_norm)
    
    # --- 5. DADOS BÁSICOS ---
    m_val = re.search(r"TOTAL\s+([\d\.]+,\d{2})", texto_norm)
    dados["valor_fatura"] = normalizar_numero_br(m_val.group(1)) if m_val else 0.0

    return dados