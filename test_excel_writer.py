from services.excel_writer import escrever_uc_geradora

dados = {
    "mes": "JAN",
    "data_leitura_anterior": "21/12/2024",
    "data_leitura_atual": "21/01/2025",
    "energia_gerada": 456.0,
    "credito_recebido": 436.0,
    "energia_ativa": 536,
    "valor_fatura": 154.04,
    "saldo_kwh": 0.0,
    "medidor": "13119425-9",
    "leitura_anterior": 15604,
    "leitura_atual": 16140,
}

escrever_uc_geradora(
    "BALANÇO E COMPENSAÇÃO.xlsx",
    dados
)

