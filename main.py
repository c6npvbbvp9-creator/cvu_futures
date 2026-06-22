"""
Ponto de entrada principal.
Limpa o Excel anterior, coleta dados do CME via TradingView e salva.
"""
import os

EXCEL_PATH = "data/CME_Futures_Database.xlsx"

# Sempre limpa o Excel antes de coletar para evitar duplicatas
if os.path.exists(EXCEL_PATH):
    os.remove(EXCEL_PATH)
    print(f"Excel anterior removido: {EXCEL_PATH}")

from scraper import collect_all
from update_excel import update_excel

if __name__ == "__main__":
    data = collect_all()
    update_excel(data)
