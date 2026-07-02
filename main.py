"""
Ponto de entrada principal.
1. Coleta curva completa de todos os contratos (CME_Futures_Database.xlsx)
2. Atualiza aba "Cotações Diárias" com preço do front month de cada ticker
"""

from scraper import collect_all
from update_excel import update_excel
from daily_quotes import run_today

if __name__ == "__main__":
    # 1. Coleta curva completa
    data = collect_all()
    update_excel(data)

    # 2. Atualiza cotações diárias com preço de hoje
    run_today()
