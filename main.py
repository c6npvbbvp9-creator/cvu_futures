"""
Ponto de entrada principal.

1. Coleta a curva completa de todos os contratos futuros via TradingView
   (HH, Brent, NBP, JKM, TTF, Coal_API2) → CME_Futures_Database.xlsx
2. Atualiza a aba "Cotações Diárias" com o contrato de referência de cada
   ticker (metodologia m_plus_1, contratos vencidos descartados).
3. Adiciona o Diesel (POAEE00) via API da EIA — fonte externa, US¢/gal.

Tickers e fontes:
  HH                         → EIA spot Henry Hub (RNGWHHD, USD/MMBtu)
  Brent                      → EIA spot Europe Brent FOB (RBRTE, USD/bbl)
  NBP, JKM, TTF              → TradingView (curva de futuros, front-month)
  Coal_API2 (CSARM01)        → TradingView ICEEUR-ATW1! (curva de futuros)
  Diesel (POAEE00)           → EIA ULSD USGC (proxy, requer EIA_API_KEY)

  NOTA: HH e Brent migraram de futuros (TradingView) para SPOT (EIA) em
  jul/2026, por decisão da gestora — a aba deve refletir mercado spot.
  TTF/NBP/JKM não têm spot público gratuito e seguem no front-month.

Pendente (sem fonte pública automatizável — índices licenciados):
  Óleo Combustível (PUAAI00) → aguardando definição de fonte/proxy
"""

from scraper import collect_all
from update_excel import update_excel
from daily_quotes import run_today


def main():
    # 1. Curva completa (inclui o carvão via ATW)
    data = collect_all()
    update_excel(data)

    # 2. Cotações diárias a partir da curva (contrato de referência por ticker).
    #    Preenche apenas NBP, JKM, TTF e Coal_API2 — HH e Brent ficam vazios
    #    aqui de propósito e são preenchidos pelo spot da EIA no passo 3.
    run_today()

    # 3. HH e Brent SPOT via EIA (não vêm da curva; mercado spot, não futuros)
    try:
        import eia_spot
        eia_spot.run()
    except SystemExit as e:
        # Falta a EIA_API_KEY — não derruba o resto do pipeline
        print(f"\n[EIA spot] Pulado: {e}")
    except Exception as e:
        print(f"\n[EIA spot] ERRO (não fatal): {e}")

    # 4. Diesel via EIA (não vem da curva; fonte externa)
    try:
        import eia_diesel
        eia_diesel.run()
    except SystemExit as e:
        # Falta a EIA_API_KEY — não derruba o resto do pipeline
        print(f"\n[Diesel] Pulado: {e}")
    except Exception as e:
        print(f"\n[Diesel] ERRO (não fatal): {e}")

    # 5. Câmbio BCB (Euro e fechamento do dólar) via API PTAX
    try:
        import bcb_cambio
        bcb_cambio.run()
    except Exception as e:
        print(f"\n[Câmbio BCB] ERRO (não fatal): {e}")


if __name__ == "__main__":
    main()
