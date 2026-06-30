"""
Ponto de entrada principal.
Coleta dados do CME via TradingView e ACUMULA no Excel histórico.
Cada execução adiciona uma nova linha por contrato/dia, sem apagar dias anteriores.
"""

from scraper import collect_all
from update_excel import update_excel

if __name__ == "__main__":
    data = collect_all()
    update_excel(data)
