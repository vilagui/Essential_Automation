import openpyxl
import datetime

def preparar_planilha(caminho_entrada, qtd_geradoras, qtd_beneficiarias):
    """Prepara o workbook criando as abas conforme a quantidade de UCs."""
    wb = openpyxl.load_workbook(caminho_entrada)
    
    if "UC GERADORA" in wb.sheetnames:
        ws_modelo_ger = wb["UC GERADORA"]
        for i in range(qtd_geradoras):
            nome_aba = f"UC GERADORA {i+1}" if i > 0 else "UC GERADORA"
            if i == 0: ws_modelo_ger.title = nome_aba
            else:
                nova = wb.copy_worksheet(ws_modelo_ger)
                nova.title = nome_aba

    nome_modelo_benef = next((s for s in wb.sheetnames if "UC BENEF" in s.upper()), None)
    if nome_modelo_benef and qtd_beneficiarias > 0:
        ws_modelo_ben = wb[nome_modelo_benef]
        for i in range(qtd_beneficiarias):
            nome_aba = f"UC BENEF. {i+1}"
            if i == 0: ws_modelo_ben.title = nome_aba
            else:
                nova = wb.copy_worksheet(ws_modelo_ben)
                nova.title = nome_aba
    return wb

def salvar_dados_A(wb, dados_estruturados):
    """Mapeia os dados detalhados para as colunas corretas de Geradoras e Beneficiárias."""
    
    meses_map = {"JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6, 
                 "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12}

    # Aba Consolidada do Grupo A
    nome_aba_geral = next((s for s in wb.sheetnames if "GRUPO A" in s.upper()), "GRUPO A")
    ws_geral = wb[nome_aba_geral] if nome_aba_geral in wb.sheetnames else None

    for item in dados_estruturados:
        tipo, indice, faturas = item['tipo'], item['indice'], item['dados']
        
        # Define a aba correta
        if tipo == 'geradora':
            nome_aba_uc = "UC GERADORA" if indice == 1 else f"UC GERADORA {indice}"
        else:
            nome_aba_uc = f"UC BENEF. {indice}"
            
        ws_uc = wb[nome_aba_uc] if nome_aba_uc in wb.sheetnames else None

        for dados in faturas:
            mes_num = meses_map.get(dados.get("mes"))
            if not mes_num: continue

            # --- 1. ABA DIMENSIONAMENTO GERAL (Colunas B, C, D e L, M, N) ---
            if ws_geral:
                for row in range(5, 25):
                    celula = ws_geral[f"A{row}"].value
                    if isinstance(celula, datetime.datetime) and celula.month == mes_num:
                        ws_geral[f"B{row}"] = dados.get("c_p", 0.0) # Consumo P
                        ws_geral[f"C{row}"] = dados.get("c_fp", 0.0) # Consumo FP
                        ws_geral[f"D{row}"] = dados.get("c_hr", 0.0) # Consumo HR
                        ws_geral[f"M{row}"] = dados.get("d_p", 0.0) # Demanda P
                        ws_geral[f"N{row}"] = dados.get("d_fp", 0.0) # Demanda FP
                        ws_geral[f"O{row}"] = dados.get("d_hr", 0.0) # Demanda HR
                        break

            # --- 2. ABA INDIVIDUAL DA UC (GERADORA VS BENEFICIÁRIA) ---
            if ws_uc:
                for row in range(5, 45):
                    celula = ws_uc[f"A{row}"].value
                    if isinstance(celula, datetime.datetime) and celula.month == mes_num:
                        # Consumo totalizado (P + FP + HR)
                        c_total = dados.get("c_p", 0) + dados.get("c_fp", 0) + dados.get("c_hr", 0)
                        
                        if tipo == 'geradora':
                            # Coluna I: Fornecida | G: Injetada | L: Valor
                            ws_uc[f"I{row}"] = dados.get("energia_gerada", 0.0)
                            ws_uc[f"J{row}"] = dados.get("credito_recebido", 0.0)
                            ws_uc[f"N{row}"] = dados.get("valor_fatura", 0.0)
                            ws_uc[f"P{row}"] = dados.get("saldo", 0.0)
                        else:
                            # Coluna F: Fornecida | H: Injetada | J: Valor | K: Saldo
                            ws_uc[f"F{row}"] = c_total
                            ws_uc[f"H{row}"] = dados.get("credito_recebido", 0.0)
                            ws_uc[f"J{row}"] = dados.get("valor_fatura", 0.0)
                            ws_uc[f"Q{row}"] = dados.get("saldo", 0.0)
                        break

    return wb