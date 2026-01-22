import openpyxl
from openpyxl.cell.cell import MergedCell
import datetime

def preparar_planilha(caminho_entrada, qtd_geradoras, qtd_beneficiarias):
    wb = openpyxl.load_workbook(caminho_entrada)
    
    # Preparar Geradoras
    if "UC GERADORA" in wb.sheetnames:
        ws_modelo_ger = wb["UC GERADORA"]
        for i in range(qtd_geradoras):
            nome_aba = f"UC GERADORA {i+1}" if i > 0 else "UC GERADORA"
            if i == 0: ws_modelo_ger.title = nome_aba
            else: 
                nova = wb.copy_worksheet(ws_modelo_ger)
                nova.title = nome_aba

    # Preparar Beneficiárias
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

def safe_write(ws, col, row, value):
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
    mapa_meses = {
        "JAN": "Jan", "FEV": "Fev", "MAR": "Mar", "ABR": "Abr",
        "MAI": "Mai", "JUN": "Jun", "JUL": "Jul", "AGO": "Ago",
        "SET": "Set", "OUT": "Out", "NOV": "Nov", "DEZ": "Dez"
    }

    # Definição das Colunas BASE (conforme image_c84419.png)
    cols = {
        'leitura_ant': 'B', 'leitura_atual': 'C',
        'geracao': 'I', 'credito': 'J',
        'consumo': 'K', 'valor': 'N',
        'saldo': 'P', 
        'medidor': 'R', 'leitura_med_ant': 'S', 'leitura_med_atual': 'T'
    }

    for item in dados_estruturados:
        tipo = item['tipo']
        indice = item['indice']
        faturas = item['dados']
        
        if tipo == 'geradora':
            nome_aba = "UC GERADORA" if indice == 1 and "UC GERADORA" in wb.sheetnames else f"UC GERADORA {indice}"
            col_saldo_atual = 'P'
            cols_uso = cols 
        else:
            nome_aba = f"UC BENEF. {indice}"
            col_saldo_atual = 'Q' # Coluna de Balanço/Saldo para beneficiária
            cols_uso = {
                **cols,
                'consumo': 'F', # Coluna E. Fornecida (Consumo)
                'credito': 'H', # Coluna E. Injetada (Compensada)
                'valor': 'J',   # Coluna Valor Fatura
                'medidor': 'S',
                'leitura_med_ant': 'T', 
                'leitura_med_atual': 'U'
            }

        if nome_aba in wb.sheetnames:
            ws = wb[nome_aba]

            for dados in faturas:
                # --- 1. DADOS DO MÊS ATUAL ---
                mes_pdf = dados.get("mes", "")
                if mes_pdf and mes_pdf in mapa_meses:
                    mes_excel_alvo = mapa_meses[mes_pdf].upper()

                    linha_destino = None
                    for row in range(5, 40):
                        celula = ws[f"A{row}"].value
                        if celula:
                            # Trata se a célula for uma Data do Excel ou Texto
                            texto_celula = celula.strftime("%b") if isinstance(celula, datetime.datetime) else str(celula)
                            if texto_celula.strip().upper()[:3] == mes_excel_alvo[:3]:
                                linha_destino = row
                                break
                        
                    if linha_destino:
                        ws[f"{cols_uso['leitura_ant']}{linha_destino}"] = dados.get("data_leitura_anterior")
                        ws[f"{cols_uso['leitura_atual']}{linha_destino}"] = dados.get("data_leitura_atual")
                        ws[f"{cols_uso['geracao']}{linha_destino}"] = dados.get("energia_gerada")
                        ws[f"{cols_uso['credito']}{linha_destino}"] = dados.get("credito_recebido")
                        ws[f"{cols_uso['consumo']}{linha_destino}"] = dados.get("energia_ativa")
                        ws[f"{cols_uso['valor']}{linha_destino}"] = dados.get("valor_fatura")
                        ws[f"{col_saldo_atual}{linha_destino}"] = dados.get("saldo")
                        ws[f"{cols_uso['medidor']}{linha_destino}"] = dados.get("medidor")

                # --- 2. PREENCHIMENTO RETROATIVO (HISTÓRICO) ---
                if "historico" in dados and dados["historico"]:
                    for hist in dados["historico"]:
                        mes_hist = hist['mes'].upper()
                        if mes_hist in mapa_meses:
                            linha_hist = None
                            for row in range(5, 40):
                                celula = ws[f"A{row}"].value
                                if celula:
                                    texto_celula = celula.strftime("%b") if isinstance(celula, datetime.datetime) else str(celula)
                                    if texto_celula.strip().upper()[:3] == mes_hist[:3]:
                                        linha_hist = row
                                        break
                            
                            if linha_hist and mes_hist != dados.get("mes", "").upper():
                                # Preenche apenas se a célula estiver vazia
                                if not ws[f"{cols_uso['consumo']}{linha_hist}"].value:
                                    ws[f"{cols_uso['consumo']}{linha_hist}"] = hist['consumo']

    # --- 3. RESUMO (UC e Endereço) ---
    ws_resumo = next((wb[s] for s in wb.sheetnames if "RESUMO" in s.upper()), None)
    if ws_resumo:
        linha_atual = 7
        for item in dados_estruturados:
            if item['dados']:
                dados_ref = item['dados'][0]
                safe_write(ws_resumo, "F", linha_atual, dados_ref.get("uc", ""))
                safe_write(ws_resumo, "G", linha_atual, dados_ref.get("endereco", ""))
                linha_atual += 1
                
    return wb