"""
Ponto de entrada principal.
Executa: coleta CME → atualiza Excel → commita no repositório.
"""

from scraper import collect_all
from update_excel import update_excel

if __name__ == "__main__":
    data = collect_all()
    update_excel(data)
