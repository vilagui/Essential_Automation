import openpyxl
from openpyxl.cell.cell import MergedCell

def preparar_planilha(caminho_entrada, qtd_geradoras, qtd_beneficiarias):
    """Carrega o modelo e cria as abas necessárias duplicando os modelos existentes."""
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

def safe_write(ws, col, row, value):
    """Escreve em células, tratando corretamente células mescladas."""
    coord = f"{col}{row}"
    cell = ws[coord]
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if coord in rng:
                ws[rng.start_cell.coordinate].value = value
                return
    else:
        cell.value = value

def salvar_dados_multiplos(wb, dados_estruturados, grupo):
    """Mapeia e escreve os dados nas abas corretas dependendo do grupo (A ou B)."""
    
    # Mapa para o Grupo A: jan/25, fev/25...
    mapa_meses_a = {
        "JAN": "jan/25", "FEV": "fev/25", "MAR": "mar/25", "ABR": "abr/25",
        "MAI": "mai/25", "JUN": "jun/25", "JUL": "jul/25", "AGO": "ago/25",
        "SET": "set/25", "OUT": "out/25", "NOV": "nov/25", "DEZ": "dez/25"
    }

    # Mapa para o Grupo B: Jan, Fev, Mar...
    mapa_meses_b = {
        "JAN": "Jan", "FEV": "Fev", "MAR": "Mar", "ABR": "Abr",
        "MAI": "Mai", "JUN": "Jun", "JUL": "Jul", "AGO": "Ago",
        "SET": "Set", "OUT": "Out", "NOV": "Nov", "DEZ": "Dez"
    }

    for item in dados_estruturados:
        tipo = item['tipo']
        indice = item['indice']
        faturas = item['dados']

        # Lógica de nomes de abas
        if grupo == "A":
            # Tenta encontrar a aba específica de dimensionamento ou a padrão
            nome_aba = "GRUPO A" if "GRUPO A" in wb.sheetnames else f"UC BENEF. {indice}"
            if tipo == 'geradora' and "UC GERADORA" in wb.sheetnames:
                nome_aba = "UC GERADORA" if indice == 1 else f"UC GERADORA {indice}"
        else:
            if tipo == 'geradora':
                nome_aba = "UC GERADORA" if indice == 1 and "UC GERADORA" in wb.sheetnames else f"UC GERADORA {indice}"
            else:
                nome_aba = f"UC BENEF. {indice}"

        if nome_aba not in wb.sheetnames:
            continue
        
        ws = wb[nome_aba]

        if grupo == "B":
            # ===============================
            # LÓGICA GRUPO B (BAIXA TENSÃO)
            # ===============================
            cols_base = {
                'leitura_ant': 'B', 'leitura_atual': 'C', 'geracao': 'I', 
                'credito': 'J', 'consumo': 'K', 'valor': 'N', 'saldo': 'P',
                'medidor': 'R', 'leitura_med_ant': 'S', 'leitura_med_atual': 'T'
            }
            
            if tipo == 'beneficiaria':
                cols_uso = {
                    **cols_base, 'consumo': 'F', 'credito': 'H', 'valor': 'J', 
                    'saldo': 'Q', 'leitura_med_ant': 'T', 'leitura_med_atual': 'U'
                }
            else:
                cols_uso = cols_base

            for dados in faturas:
                mes_pdf = dados.get("mes", "")
                mes_excel = mapa_meses_b.get(mes_pdf)
                
                if mes_excel:
                    for row in range(5, 40):
                        celula = ws[f"A{row}"].value
                        if celula and str(celula).strip().lower() == mes_excel.lower():
                            ws[f"{cols_uso['leitura_ant']}{row}"] = dados.get("data_leitura_anterior")
                            ws[f"{cols_uso['leitura_atual']}{row}"] = dados.get("data_leitura_atual")
                            ws[f"{cols_uso['geracao']}{row}"] = dados.get("energia_gerada")
                            ws[f"{cols_uso['credito']}{row}"] = dados.get("credito_recebido")
                            ws[f"{cols_uso['consumo']}{row}"] = dados.get("energia_ativa")
                            ws[f"{cols_uso['valor']}{row}"] = dados.get("valor_fatura")
                            ws[f"{cols_uso['saldo']}{row}"] = dados.get("saldo")
                            ws[f"{cols_uso['medidor']}{row}"] = dados.get("medidor")
                            ws[f"{cols_uso['leitura_med_ant']}{row}"] = dados.get("leitura_anterior")
                            ws[f"{cols_uso['leitura_med_atual']}{row}"] = dados.get("leitura_atual")
                            break
        
        else:
            # ===============================
            # LÓGICA GRUPO A (ALTA TENSÃO)
            # ===============================
            for dados in faturas:
                # 1. Preencher Mês Atual da Fatura
                mes_ref = mapa_meses_a.get(dados.get("mes"))
                if mes_ref:
                    for row in range(5, 20):
                        celula_a = str(ws[f"A{row}"].value).strip().lower()
                        if celula_a == mes_ref.lower():
                            # DADOS DE CONSUMO (Colunas B, C, D)
                            ws[f"B{row}"] = dados.get("c_p")   
                            ws[f"C{row}"] = dados.get("c_fp")  
                            ws[f"D{row}"] = dados.get("c_hr")  
                            # DADOS DE DEMANDA (Colunas L, M, N)
                            ws[f"L{row}"] = dados.get("d_p")   
                            ws[f"M{row}"] = dados.get("d_fp")  
                            ws[f"N{row}"] = dados.get("d_hr") # Incluída Demanda Reservada
                            break

                # 2. Preencher Histórico Retroativo (Preenchimento Automático)
                if "historico" in dados and dados["historico"]:
                    for h in dados["historico"]:
                        mes_h = mapa_meses_a.get(h['mes'])
                        # Pula o mês atual para não duplicar
                        if not mes_h or mes_h.lower() == str(mes_ref).lower():
                            continue
                        
                        for row in range(5, 20):
                            celula_a = str(ws[f"A{row}"].value).strip().lower()
                            if celula_a == mes_h.lower():
                                # Consumo
                                ws[f"B{row}"] = h.get('consumo_p')
                                ws[f"C{row}"] = h.get('consumo_fp')
                                ws[f"D{row}"] = h.get('consumo_hr')
                                # Demanda
                                ws[f"L{row}"] = h.get('demanda_p')
                                ws[f"M{row}"] = h.get('demanda_fp')
                                ws[f"N{row}"] = h.get('demanda_hr')
                                break
                                
    return wb