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

# AQUI VAI A PARTE DO HISTORICO (Como tá dando problema pro grupo B e o grupo A é mais complexo) vou tentar fazer depois

def extrair_historico_consumo(texto: str) -> list:
    """
    Captura as colunas: Mês/Ano, Demanda (P, FP, RE), 
    Consumo Faturado (P, FP, RE) e Horário Reservado (Consumo).
    """
    historico = []
    # Regex para capturar a linha da tabela de histórico (image_b84cef.png)
    # Procura Mes/Ano seguido de 7 a 9 blocos numéricos
    padrao = r"(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\s*[\/\-]\s*(\d{2,4})((?:\s+[\d\.,]+){7,9})"
    
    matches = re.findall(padrao, texto)
    
    for mes, ano, valores_str in matches:
        # Divide os valores numéricos capturados no bloco
        v = valores_str.strip().split()
        if len(v) >= 7:
            historico.append({
                "mes": mes,
                "ano": int(ano),
                "demanda_ponta": normalizar_numero_br(v[0]),
                "demanda_fora_ponta": normalizar_numero_br(v[1]),
                "demanda_reativo": normalizar_numero_br(v[2]),
                "consumo_ponta": normalizar_numero_br(v[3]),
                "consumo_fora_ponta": normalizar_numero_br(v[4]),
                "consumo_reativo": normalizar_numero_br(v[5]),
                "reservado_consumo": normalizar_numero_br(v[6])
            })
    return historico

def extrair_fatura(texto: str) -> dict:
    dados = {}
    texto = normalizar_texto(texto)

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

    # --- 4. MEDIDOR ---
    m = re.search(r"(\d{7,}-\d)\s+ENERGIA ATIVA - KWH ÚNICO\s+(\d+)\s+(\d+)", texto)
    if m:
        dados["medidor"] = m.group(1)
        dados["leitura_anterior"] = int(m.group(2))
        dados["leitura_atual"] = int(m.group(3))
    else:
        dados["medidor"] = "" 
        dados["leitura_anterior"] = 0
        dados["leitura_atual"] = 0

    # --- 5. ENERGIA ATIVA (Consumo Atual) ---
    dados["energia_ativa"] = 0.0
    m_ativa_p = re.search(r"ENERGIA ATIVA - KWH PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)", texto)
    m_ativa_fp = re.search(r"ENERGIA ATIVA - KWH FORA PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)", texto)
    m_ativa_hr = re.search(r"ENERGIA ATIVA - KWH RESERVADO\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)", texto)
    m_ativa = m_ativa_p+m_ativa_fp+m_ativa_hr

    if m_ativa:
        dados["energia_ativa"] = normalizar_numero_br(m_ativa.group(1))

    # --- 6. GERAÇÃO, CRÉDITO E SALDO ---

    dados["energia_gerada"] = 0.0
    dados["credito_recebido"] = 0.0
    dados["saldo"] = 0.0

    # Tenta geração na linha
    m_geracao_linha_ponta = re.search(r"ENERGIA GERAÇÃO - KWH PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)", texto)
    m_geracao_linha_foraponta = re.search(r"ENERGIA GERAÇÃO - KWH FORA PONTA\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)", texto)
    m_geracao_linha_reservado = re.search(r"ENERGIA GERAÇÃO - KWH RESERVADO\s+\d+\s+\d+\s+[\d,]+\s+([\d,]+)", texto)
    m_geracao_linha = m_geracao_linha_foraponta+m_geracao_linha_ponta+m_geracao_linha_reservado

    if m_geracao_linha:
        dados["energia_gerada"] = normalizar_numero_br(m_geracao_linha.group(1))
    
    # Bloco SCEE
    idx_scee = texto.find("INFORMAÇÕES DO SCEE")
    if idx_scee != -1:
        bloco_busca = texto[idx_scee : idx_scee + 1000]
        
        # Fallback Geração
        if dados["energia_gerada"] == 0:
            m_ger_scee = re.search(r"GERAÇÃO CICLO.*?UC\s+\d+\s*:\s*([\d,]+)", bloco_busca)
            if m_ger_scee:
                dados["energia_gerada"] = normalizar_numero_br(m_ger_scee.group(1))
        
        # Crédito
        m_credito = re.search(r"CRÉDITO RECEBIDO.*?([\d\.]+,\d{2})", bloco_busca)
        if m_credito:
            dados["credito_recebido"] = normalizar_numero_br(m_credito.group(1))


        # Saldo (Captura todos os valores P, FP e HR e soma)
        m_saldo_bloco = re.search(r"SALDO KWH[\s\S]*?(?=[\n\r]|$)", bloco_busca)
        if m_saldo_bloco:
            trecho_saldos = m_saldo_bloco.group(0)
            # Encontra todos os padrões numéricos (ex: 0,00 ou 1.423,32) no trecho
            valores_encontrados = re.findall(r"[\d\.]*,\d{2}", trecho_saldos)
    
            # Soma todos os valores convertidos
            dados["saldo"] = sum(normalizar_numero_br(v) for v in valores_encontrados)

    # --- 7. VALOR ---
    m = re.search(r"TOTAL\s+([\d\.]+,\d{2})", texto)
    dados["valor_fatura"] = normalizar_numero_br(m.group(1)) if m else 0.0

    # --- 8. HISTÓRICO DE CONSUMO E DEMANDA (NOVO) ---
    
    

    return dados

