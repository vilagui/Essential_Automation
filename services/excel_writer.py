from openpyxl import load_workbook

MAPA_MESES = {
    "JAN": 5, "FEV": 6, "MAR": 7, "ABR": 8,
    "MAI": 9, "JUN": 10, "JUL": 11, "AGO": 12,
    "SET": 13, "OUT": 14, "NOV": 15, "DEZ": 16
}

def escrever_uc_geradora(caminho_excel, dados):
    wb = load_workbook(caminho_excel)
    ws = wb["UC GERADORA"]

    linha = MAPA_MESES[dados["mes"]]

    ws[f"B{linha}"] = dados["data_leitura_anterior"]
    ws[f"C{linha}"] = dados["data_leitura_atual"]
    ws[f"I{linha}"] = dados["energia_gerada"]
    ws[f"J{linha}"] = dados["credito_recebido"]
    ws[f"K{linha}"] = dados["energia_ativa"]
    ws[f"N{linha}"] = dados["valor_fatura"]
    ws[f"P{linha}"] = dados["saldo_kwh"]
    ws[f"R{linha}"] = dados["medidor"]
    ws[f"S{linha}"] = dados["leitura_anterior"]
    ws[f"T{linha}"] = dados["leitura_atual"]

    wb.save(caminho_excel)
