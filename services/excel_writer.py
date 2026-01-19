import openpyxl
from openpyxl.cell.cell import MergedCell

def preparar_planilha(caminho_entrada, qtd_geradoras, qtd_beneficiarias):
    """
    Abre a planilha base, cria as cópias das abas necessárias e retorna o workbook.
    """
    wb = openpyxl.load_workbook(caminho_entrada)
    
    # --- 1. PREPARAR ABAS GERADORAS ---
    if "UC GERADORA" in wb.sheetnames:
        ws_modelo_ger = wb["UC GERADORA"]
        
        for i in range(qtd_geradoras):
            nome_aba = f"UC GERADORA {i+1}" if i > 0 else "UC GERADORA"
            if i == 0:
                ws_modelo_ger.title = nome_aba
            else:
                nova_aba = wb.copy_worksheet(ws_modelo_ger)
                nova_aba.title = nome_aba

    # --- 2. PREPARAR ABAS BENEFICIÁRIAS ---
    # Procura modelo de beneficiária
    nome_modelo_benef = next((s for s in wb.sheetnames if "UC BENEF" in s.upper()), None)
    
    if nome_modelo_benef and qtd_beneficiarias > 0:
        ws_modelo_ben = wb[nome_modelo_benef]
        
        for i in range(qtd_beneficiarias):
            nome_aba = f"UC BENEF. {i+1}"
            if i == 0:
                ws_modelo_ben.title = nome_aba
            else:
                nova_aba = wb.copy_worksheet(ws_modelo_ben)
                nova_aba.title = nome_aba

    return wb

def safe_write(ws, col, row, value):
    """
    Escreve em uma célula com segurança. Se for uma célula mesclada (MergedCell),
    procura a célula superior esquerda (Top-Left) da mesclagem para escrever nela.
    """
    coord = f"{col}{row}"
    cell = ws[coord]
    
    if isinstance(cell, MergedCell):
        # Se for MergedCell, significa que NÃO é a célula principal.
        # Precisamos achar quem é a "dona" dessa área mesclada.
        for rng in ws.merged_cells.ranges:
            if coord in rng:
                # Encontrou o intervalo. Escreve na célula inicial (Top-Left)
                top_left = ws[rng.start_cell.coordinate]
                top_left.value = value
                return
    else:
        # Célula normal ou a Top-Left de uma mesclagem
        cell.value = value

def salvar_dados_multiplos(wb, dados_estruturados):
    
    mapa_meses = {
        "JAN": "Jan", "FEV": "Fev", "MAR": "Mar", "ABR": "Abr",
        "MAI": "Mai", "JUN": "Jun", "JUL": "Jul", "AGO": "Ago",
        "SET": "Set", "OUT": "Out", "NOV": "Nov", "DEZ": "Dez"
    }

    # Definição das Colunas para as Abas de Mês a Mês
    cols = {
        'leitura_ant': 'B',
        'leitura_atual': 'C',
        'geracao': 'I',        # Energia Geração
        'credito': 'J',        # Crédito Recebido
        'consumo': 'K',        # Energia Ativa
        'valor': 'N',          # Valor Fatura
        'saldo': 'P',          # Saldo Kwh (Coluna P conforme solicitado)
        'medidor': 'R',
        'leitura_med_ant': 'S',
        'leitura_med_atual': 'T'
    }

    # --- 1. PREENCHER AS ABAS DE CADA UC ---
    for item in dados_estruturados:
        tipo = item['tipo']
        indice = item['indice']
        faturas = item['dados']
        
        if tipo == 'geradora':
            nome_aba = "UC GERADORA" if indice == 1 and "UC GERADORA" in wb.sheetnames else f"UC GERADORA {indice}"
        else:
            nome_aba = f"UC BENEF. {indice}"
        
        if nome_aba in wb.sheetnames:
            ws = wb[nome_aba]
            
            for dados in faturas:
                mes_pdf = dados.get("mes", "")
                if not mes_pdf or mes_pdf not in mapa_meses:
                    continue
                
                mes_excel = mapa_meses[mes_pdf]
                
                linha_destino = None
                # Busca o mês na coluna A (Linhas 5 a 40)
                for row in range(5, 40):
                    celula = ws[f"A{row}"].value
                    if celula and str(celula).strip() == mes_excel:
                        linha_destino = row
                        break
                
                if linha_destino:
                    ws[f"{cols['leitura_ant']}{linha_destino}"] = dados["data_leitura_anterior"]
                    ws[f"{cols['leitura_atual']}{linha_destino}"] = dados["data_leitura_atual"]
                    ws[f"{cols['geracao']}{linha_destino}"] = dados["energia_gerada"]
                    ws[f"{cols['credito']}{linha_destino}"] = dados["credito_recebido"]
                    ws[f"{cols['consumo']}{linha_destino}"] = dados["energia_ativa"]
                    ws[f"{cols['valor']}{linha_destino}"] = dados["valor_fatura"]
                    ws[f"{cols['saldo']}{linha_destino}"] = dados["saldo"]
                    ws[f"{cols['medidor']}{linha_destino}"] = dados["medidor"]
                    ws[f"{cols['leitura_med_ant']}{linha_destino}"] = dados["leitura_anterior"]
                    ws[f"{cols['leitura_med_atual']}{linha_destino}"] = dados["leitura_atual"]

    # --- 2. PREENCHER O RESUMO (LISTA DE UCs) ---
    ws_resumo = None
    for sheet in wb.sheetnames:
        if "RESUMO" in sheet.upper():
            ws_resumo = wb[sheet]
            break
    
    if ws_resumo:
        linha_atual = 7  # Começa na linha 7
        
        # Ordem: Primeiro todas as Geradoras
        for item in dados_estruturados:
            if item['tipo'] == 'geradora' and item['dados']:
                dados_ref = item['dados'][0]
                
                # UC na coluna F
                safe_write(ws_resumo, "F", linha_atual, dados_ref.get("uc", ""))
                
                # Endereço na coluna G (que deve ser a principal da mesclagem GH)
                safe_write(ws_resumo, "G", linha_atual, dados_ref.get("endereco", ""))
                
                linha_atual += 1
        
        # Depois todas as Beneficiárias
        for item in dados_estruturados:
            if item['tipo'] == 'beneficiaria' and item['dados']:
                dados_ref = item['dados'][0]
                
                # UC na coluna F
                safe_write(ws_resumo, "F", linha_atual, dados_ref.get("uc", ""))
                
                # Endereço na coluna G
                safe_write(ws_resumo, "G", linha_atual, dados_ref.get("endereco", ""))
                
                linha_atual += 1

    return wb