from services.fatura_mapper import extrair_fatura

with open("texto.txt", "r", encoding="utf-8") as f:
    texto = f.read()

print(extrair_fatura(texto))
