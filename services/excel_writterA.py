import openpyxl
from openpyxl.cell.cell import MergedCell
import datetime

def preparar_planilha(caminho_entrada, qtd_geradoras, qtd_beneficiarias):
    """Carrega o modelo e cria as abas necessárias para Geradoras e Beneficiárias."""
    wb = openpyxl.load_workbook(caminho_entrada)
    
    # Preparar Geradoras
    if "UC GERADORA" in wb.sheetnames:
        ws_modelo_ger = wb["UC GERADORA"]
        for i in range(qtd_geradoras):
            nome_aba = f"UC GERADORA {i+1}" if i > 0 else "UC GERADORA"
            if i == 0:
                ws_modelo_ger.title = nome_aba
            else:
                nova = wb.copy_worksheet(ws_modelo_ger)
                nova.title = nome_aba

    # Preparar Beneficiárias
    nome_modelo_benef = next((s for s in wb.sheetnames if "UC BENEF" in s.upper()), None)
    if nome_modelo_benef and qtd_beneficiarias > 0:
        ws_modelo_ben = wb[nome_modelo_benef]
        for i in range(qtd_beneficiarias):
            nome_aba = f"UC BENEF. {i+1}"
            if i == 0:
                ws_modelo_ben.title = nome_aba
            else:
                nova = wb.copy_worksheet(ws_modelo_ben)
                nova.title = nome_aba
    return wb

def salvar_dados_A(wb, dados_estruturados):
    """Mapeia os dados do Grupo A para a aba de Dimensionamento e abas de UC."""
    
    # Mapa para identificar o mês na coluna A (Excel costuma usar objetos datetime)
    meses_map = {
        "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
        "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12
    }

    # 1. Identificar a aba de Dimensionamento do Grupo A [cite: 6]
    nome_aba_geral = next((s for s in wb.sheetnames if "GRUPO A" in s.upper()), "GRUPO A")
    ws_geral = wb[nome_aba_geral] if nome_aba_geral in wb.sheetnames else None

    for item in dados_estruturados:
        tipo = item['tipo']
        indice = item['indice']
        faturas = item['dados']
        
        # 2. Identificar a aba individual da UC 
        if tipo == 'geradora':
            nome_aba_uc = "UC GERADORA" if indice == 1 else f"UC GERADORA {indice}"
        else:
            nome_aba_uc = f"UC BENEF. {indice}"
            
        ws_uc = wb[nome_aba_uc] if nome_aba_uc in wb.sheetnames else None

        for dados in faturas:
            mes_ref = dados.get("mes")
            mes_num = meses_map.get(mes_ref)
            if not mes_num: continue

            # --- PREENCHIMENTO NA ABA DIMENSIONAMENTO GERAL --- [cite: 6]
            if ws_geral:
                for row in range(5, 25): # Mapeia jan/25 a dez/25
                    celula = ws_geral[f"A{row}"].value
                    if isinstance(celula, datetime.datetime) and celula.month == mes_num:
                        # Consumo (Colunas B, C, D)
                        ws_geral[f"B{row}"] = dados.get("c_p", 0.0)
                        ws_geral[f"C{row}"] = dados.get("c_fp", 0.0)
                        ws_geral[f"D{row}"] = dados.get("c_hr", 0.0)
                        # Demanda (Colunas L, M, N)
                        ws_geral[f"L{row}"] = dados.get("d_p", 0.0)
                        ws_geral[f"M{row}"] = dados.get("d_fp", 0.0)
                        ws_geral[f"N{row}"] = dados.get("d_hr", 0.0)
                        break

            # --- PREENCHIMENTO NA ABA DA UC --- 
            if ws_uc:
                for row in range(5, 45):
                    celula = ws_uc[f"A{row}"].value
                    if isinstance(celula, datetime.datetime) and celula.month == mes_num:
                        # Soma totalizada para o Grupo A (P + FP + HR)
                        consumo_total = dados.get("c_p", 0) + dados.get("c_fp", 0) + dados.get("c_hr", 0)
                        
                        if tipo == 'geradora':
                            ws_uc[f"I{row}"] = consumo_total  # E. Fornecida [cite: 7]
                            ws_uc[f"G{row}"] = dados.get("energia_gerada", 0.0) # E. Injetada
                            ws_uc[f"L{row}"] = dados.get("valor_fatura", 0.0)
                        else:
                            ws_uc[f"F{row}"] = consumo_total  # E. Fornecida [cite: 8]
                            ws_uc[f"H{row}"] = dados.get("credito_recebido", 0.0) # E. Injetada
                            ws_uc[f"J{row}"] = dados.get("valor_fatura", 0.0)
                            ws_uc[f"K{row}"] = dados.get("saldo", 0.0) # Balanço
                        break

            # --- PREENCHIMENTO DE HISTÓRICO ---
            if "historico" in dados:
                for h in dados["historico"]:
                    h_mes_num = meses_map.get(h['mes'])
                    if not h_mes_num or h_mes_num == mes_num: continue

                    if ws_geral:
                        for row in range(5, 25):
                            c = ws_geral[f"A{row}"].value
                            if isinstance(c, datetime.datetime) and c.month == h_mes_num:
                                ws_geral[f"B{row}"], ws_geral[f"C{row}"], ws_geral[f"D{row}"] = h['c_p'], h['c_fp'], h['c_hr']
                                ws_geral[f"L{row}"], ws_geral[f"M{row}"], ws_geral[f"N{row}"] = h['d_p'], h['d_fp'], h['d_hr']
                                break
                                
    return wb