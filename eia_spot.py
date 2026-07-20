"""
HH (Henry Hub) e Brent — preços SPOT via API pública da EIA.

Decisão metodológica (jul/2026, aprovada pela gestora):
  A aba "Cotações Diárias" deve refletir o mercado SPOT, não contratos
  futuros. HH e Brent têm spot diário público e gratuito na EIA e passam
  a vir daqui, em vez das abas de curva do TradingView.

  Os demais tickers de gás sem spot público gratuito (TTF, NBP, JKM)
  permanecem no front-month da curva — ver daily_quotes.py.

Séries EIA (API v2):
  HH    → natural-gas/pri/fut, série RNGWHHD
          Henry Hub Natural Gas Spot Price, diária, USD/MMBtu
          https://www.eia.gov/dnav/ng/hist/rngwhhdD.htm
  Brent → petroleum/pri/spt, série RBRTE
          Europe Brent Spot Price FOB, diária, USD/barril
          https://www.eia.gov/dnav/pet/hist/rbrteD.htm

Ambas as séries são republicadas pela EIA a partir de fontes comerciais
(origem Refinitiv), distribuídas gratuitamente. Podem ter defasagem de
1-2 dias (feriados dos EUA, revisões). O gap é tratado por carry-forward
no daily_quotes (usa o último valor disponível).

Uso:
  python eia_spot.py --backfill        → 1 ano de histórico
  python eia_spot.py                   → só os dias recentes
  python eia_spot.py --api-key SUACHAVE

API key: gratuita em https://www.eia.gov/opendata/register.php
         Via --api-key ou variável de ambiente EIA_API_KEY.
"""

import os
import sys
import json
import urllib.parse
import urllib.request
from datetime import date, timedelta

# Configuração por ticker: rota da API v2, série e rótulo de origem.
# 'route' é o segmento entre /v2/ e /data/ na API da EIA.
SPOT_CONFIG = {
    "HH": {
        "route": "natural-gas/pri/fut",
        "series": "RNGWHHD",
        "label": "EIA spot Henry Hub (USD/MMBtu)",
    },
    "Brent": {
        "route": "petroleum/pri/spt",
        "series": "RBRTE",
        "label": "EIA spot Europe Brent FOB (USD/bbl)",
    },
}

EIA_BASE = "https://api.eia.gov/v2/{route}/data/"


def _api_key() -> str:
    for i, arg in enumerate(sys.argv):
        if arg == "--api-key" and i + 1 < len(sys.argv):
            return sys.argv[i + 1]
    key = os.environ.get("EIA_API_KEY", "").strip()
    if not key:
        raise SystemExit(
            "ERRO: falta a API key da EIA.\n"
            "  Registre-se (gratuito): https://www.eia.gov/opendata/register.php\n"
            "  Depois: export EIA_API_KEY=sua_chave\n"
            "  Ou use: python eia_spot.py --api-key SUA_CHAVE"
        )
    return key


def fetch_spot(ticker: str, date_from: str, date_to: str,
               api_key: str = None) -> dict:
    """
    Busca a série spot diária de um ticker (HH ou Brent) na API v2 da EIA.
    Retorna {date_str: price_str}. Preço na unidade nativa da série
    (USD/MMBtu para HH, USD/barril para Brent) — sem conversão.
    """
    cfg = SPOT_CONFIG[ticker]
    api_key = api_key or _api_key()

    params = [
        ("api_key", api_key),
        ("frequency", "daily"),
        ("data[0]", "value"),
        ("facets[series][]", cfg["series"]),
        ("start", date_from),
        ("end", date_to),
        ("sort[0][column]", "period"),
        ("sort[0][direction]", "asc"),
        ("length", "5000"),
    ]
    url = EIA_BASE.format(route=cfg["route"]) + "?" + urllib.parse.urlencode(params)

    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json",
    })

    try:
        raw = urllib.request.urlopen(req, timeout=30).read()
        payload = json.loads(raw)
    except Exception as e:
        print(f"  [ERRO] {ticker}: falha ao consultar a EIA: {e}")
        return {}

    rows = payload.get("response", {}).get("data", [])
    if not rows:
        print(f"  [AVISO] {ticker}: a EIA não retornou dados.")
        print(f"          Resposta: {json.dumps(payload)[:300]}")
        return {}

    out = {}
    for r in rows:
        period = r.get("period")          # 'YYYY-MM-DD'
        value  = r.get("value")           # unidade nativa da série
        if not period or value in (None, ""):
            continue
        try:
            price = float(value)
        except (TypeError, ValueError):
            continue
        out[str(period)[:10]] = f"{price:.4f}"

    return out


def run(date_from: str = None, date_to: str = None, overwrite: bool = False):
    """
    Coleta HH e Brent spot e grava na aba 'Cotações Diárias'.

    overwrite=False (padrão): uso diário — preenche sem apagar outras fontes.
    overwrite=True: usado no backfill/rebuild, para SUBSTITUIR o histórico
                    antigo de HH/Brent (que vinha dos futuros) pelo spot.
    """
    from daily_quotes import update_daily_quotes

    if date_to is None:
        date_to = date.today().isoformat()
    if date_from is None:
        date_from = (date.today() - timedelta(days=7)).isoformat()

    api_key = _api_key()
    by_date = {}

    for ticker, cfg in SPOT_CONFIG.items():
        print(f"\n[{ticker}] {cfg['label']} — {date_from} → {date_to}")
        prices = fetch_spot(ticker, date_from, date_to, api_key)
        print(f"  {len(prices)} dias coletados")
        if not prices:
            print("  Nada a gravar para este ticker.")
            continue

        lo, hi = min(prices), max(prices)
        print(f"  {lo}: {prices[lo]}  →  {hi}: {prices[hi]}")

        for d, p in prices.items():
            by_date.setdefault(d, {})[ticker] = (p, cfg["label"])

    if not by_date:
        print("\n[EIA spot] Nada coletado.")
        return

    update_daily_quotes(by_date, overwrite=overwrite)


if __name__ == "__main__":
    if "--backfill" in sys.argv:
        # Backfill reconstrói o histórico de HH/Brent com spot, SUBSTITUINDO
        # os valores antigos que vinham dos futuros (overwrite=True).
        run(date_from=(date.today() - timedelta(days=366)).isoformat(),
            overwrite=True)
    else:
        run()
