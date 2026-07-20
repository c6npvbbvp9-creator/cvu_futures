"""
Setup único: valida as séries spot da EIA (HH e Brent) e reconstrói o
histórico dessas duas colunas na aba "Cotações Diárias", substituindo os
valores antigos que vinham dos futuros (TradingView).

Rode UMA vez, localmente, com sua EIA_API_KEY:

    export EIA_API_KEY=sua_chave
    python setup_spot_hh_brent.py                 # backfill 1 ano
    python setup_spot_hh_brent.py --days 60        # backfill últimos 60 dias
    python setup_spot_hh_brent.py --check          # só valida, não grava

Depois confira a planilha e, se estiver ok, faça o commit.
O uso DIÁRIO passa a ser automático via main.py (eia_spot.run()).
"""

import sys
from datetime import date, timedelta
import eia_spot


def check_only(date_from, date_to):
    """Só busca e imprime — não grava nada. Serve para validar as séries."""
    key = eia_spot._api_key()
    print(f"Validando séries da EIA de {date_from} a {date_to}...\n")
    ok = True
    for ticker, cfg in eia_spot.SPOT_CONFIG.items():
        prices = eia_spot.fetch_spot(ticker, date_from, date_to, key)
        if prices:
            lo, hi = min(prices), max(prices)
            print(f"  [OK]   {ticker:6s} {len(prices):3d} dias | "
                  f"{lo}={prices[lo]}  ...  {hi}={prices[hi]}")
        else:
            print(f"  [FALHA] {ticker:6s} nenhum dado retornado — "
                  f"verifique série '{cfg['series']}' / rota '{cfg['route']}'")
            ok = False
    print("\nResultado:", "todas as séries OK." if ok else "há séries com falha.")
    return ok


def main():
    days = 366
    if "--days" in sys.argv:
        i = sys.argv.index("--days")
        if i + 1 < len(sys.argv):
            days = int(sys.argv[i + 1])

    date_to = date.today().isoformat()
    date_from = (date.today() - timedelta(days=days)).isoformat()

    if "--check" in sys.argv:
        check_only(date_from, date_to)
        return

    # 1. valida antes de gravar
    if not check_only(date_from, date_to):
        print("\nAbortado: corrija as séries antes de gravar.")
        sys.exit(1)

    # 2. backfill com overwrite (substitui histórico de futuros pelo spot)
    print("\nGravando spot no histórico (overwrite=True)...")
    eia_spot.run(date_from=date_from, date_to=date_to, overwrite=True)
    print("\nConcluído. Confira a aba 'Cotações Diárias' e faça o commit.")


if __name__ == "__main__":
    main()
