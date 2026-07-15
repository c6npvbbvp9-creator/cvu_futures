"""
Diesel (ODI) — coleta via API pública da EIA.

Indicador de referência da gestora:
  POAEE00 — Platts "No. 2", US Gulf Coast, US¢/gal

Proxy adotado:
  EIA — U.S. Gulf Coast Ultra-Low Sulfur No 2 Diesel Spot Price
  Série: EER_EPD2DXL0_PF4_RGC_DPG (diária, USD/gal desde 2006)
  https://www.eia.gov/dnav/pet/hist/eer_epd2dxl0_pf4_rgc_dpgD.htm

DESVIO CONHECIDO (documentar na metodologia):
  O EIA publica ULSD (ultra baixo enxofre, 15ppm). O Platts POAEE00 reflete
  "No. 2" genérico do USGC. São produtos vizinhos, não idênticos — há basis.
  Geografia e unidade batem; a especificação de enxofre não.
  O basis não foi medido (não temos a série Platts para comparar).

FONTE DA SÉRIE:
  A EIA republica preços spot wholesale a partir da Refinitiv. Ou seja, a
  série tem origem comercial mesmo sendo distribuída gratuitamente.

Uso:
  python eia_diesel.py --backfill        → 1 ano de histórico
  python eia_diesel.py                   → só os dias recentes
  python eia_diesel.py --api-key SUACHAVE

API key: gratuita em https://www.eia.gov/opendata/register.php
         Pode ser passada via --api-key ou variável de ambiente EIA_API_KEY.
"""

import os
import sys
import json
import time
import urllib.parse
import urllib.request
from datetime import date, timedelta

TICKER = "Diesel"

# Série EIA — ULSD spot, US Gulf Coast, diária
EIA_SERIES_ID = "EER_EPD2DXL0_PF4_RGC_DPG"
EIA_BASE = "https://api.eia.gov/v2/petroleum/pri/spt/data/"

# Platts publica em US¢/gal; a EIA publica em USD/gal.
CENTS_PER_DOLLAR = 100

SOURCE_LABEL = "EIA ULSD USGC (proxy POAEE00)"


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
            "  Ou use: python eia_diesel.py --api-key SUA_CHAVE"
        )
    return key


def fetch_diesel(date_from: str, date_to: str, api_key: str = None) -> dict:
    """
    Busca a série diária de ULSD USGC na API v2 da EIA.
    Retorna {date_str: price_str} já convertido para US¢/gal.
    """
    api_key = api_key or _api_key()

    params = [
        ("api_key", api_key),
        ("frequency", "daily"),
        ("data[0]", "value"),
        ("facets[series][]", EIA_SERIES_ID),
        ("start", date_from),
        ("end", date_to),
        ("sort[0][column]", "period"),
        ("sort[0][direction]", "asc"),
        ("length", "5000"),
    ]
    url = EIA_BASE + "?" + urllib.parse.urlencode(params)

    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json",
    })

    try:
        raw = urllib.request.urlopen(req, timeout=30).read()
        payload = json.loads(raw)
    except Exception as e:
        print(f"  [ERRO] Falha ao consultar a EIA: {e}")
        return {}

    rows = payload.get("response", {}).get("data", [])
    if not rows:
        print("  [AVISO] A EIA não retornou dados.")
        print(f"          Resposta: {json.dumps(payload)[:300]}")
        return {}

    out = {}
    for r in rows:
        period = r.get("period")          # 'YYYY-MM-DD'
        value  = r.get("value")           # USD/gal
        if not period or value in (None, ""):
            continue
        try:
            cents = float(value) * CENTS_PER_DOLLAR
        except (TypeError, ValueError):
            continue
        out[str(period)[:10]] = f"{cents:.4f}"

    return out


def run(date_from: str = None, date_to: str = None):
    """Coleta e grava na aba 'Cotações Diárias' via daily_quotes."""
    from daily_quotes import update_daily_quotes

    if date_to is None:
        date_to = date.today().isoformat()
    if date_from is None:
        date_from = (date.today() - timedelta(days=7)).isoformat()

    print(f"\n[Diesel] EIA ULSD USGC — {date_from} → {date_to}")
    prices = fetch_diesel(date_from, date_to)
    print(f"  {len(prices)} dias coletados (US¢/gal)")

    if not prices:
        print("  Nada a gravar.")
        return

    lo, hi = min(prices), max(prices)
    print(f"  {lo}: {prices[lo]}  →  {hi}: {prices[hi]}")

    by_date = {
        d: {TICKER: (p, SOURCE_LABEL)}
        for d, p in prices.items()
    }
    update_daily_quotes(by_date, overwrite=False)


if __name__ == "__main__":
    if "--backfill" in sys.argv:
        run(date_from=(date.today() - timedelta(days=366)).isoformat())
    else:
        run()
