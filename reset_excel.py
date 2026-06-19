"""
Apaga completamente o Excel para forçar coleta limpa.
Roda antes do scraper a cada execução do workflow.
"""
import os

EXCEL_PATH = "data/CME_Futures_Database.xlsx"

if os.path.exists(EXCEL_PATH):
    os.remove(EXCEL_PATH)
    print(f"Excel removido: {EXCEL_PATH}")
else:
    print("Excel não existe — nada a remover.")
