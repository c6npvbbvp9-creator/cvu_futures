"""
Script de diagnóstico — rode localmente para inspecionar a resposta da API do CME.
Útil para confirmar que todos os contratos (até 2038) estão sendo retornados.

Uso:
    python debug_api.py
"""

import requests
import json

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.cmegroup.com/markets/energy/natural-gas/natural-gas.settlements.html",
})

TESTS = [
    {
        "label": "HH settlements (deveria ter ~150 contratos até 2038)",
        "url": "https://www.cmegroup.com/CmeWS/mvc/Settlements/futures/settlements/444/G",
    },
    {
        "label": "HH quotes (API antiga — retorna poucos contratos)",
        "url": "https://www.cmegroup.com/CmeWS/mvc/Quotes/Future/444/G",
    },
    {
        "label": "Brent settlements",
        "url": "https://www.cmegroup.com/CmeWS/mvc/Settlements/futures/settlements/425/G",
    },
    {
        "label": "TTF settlements",
        "url": "https://www.cmegroup.com/CmeWS/mvc/Settlements/futures/settlements/8309/G",
    },
    {
        "label": "JKM settlements",
        "url": "https://www.cmegroup.com/CmeWS/mvc/Settlements/futures/settlements/8391/G",
    },
    {
        "label": "NBP settlements",
        "url": "https://www.cmegroup.com/CmeWS/mvc/Settlements/futures/settlements/7872/G",
    },
    {
        "label": "Coal API2 settlements",
        "url": "https://www.cmegroup.com/CmeWS/mvc/Settlements/futures/settlements/993/G",
    },
]

# Aquece a sessão
SESSION.get("https://www.cmegroup.com", timeout=10)

for test in TESTS:
    print(f"\n{'='*60}")
    print(f"  {test['label']}")
    print(f"  URL: {test['url']}")
    try:
        r = SESSION.get(test["url"], timeout=20)
        print(f"  HTTP status: {r.status_code}")
        if r.status_code == 200:
            data = r.json()
            # Descobre onde estão os contratos
            if isinstance(data, list):
                items = data
            elif isinstance(data, dict):
                items = (
                    data.get("settlements")
                    or data.get("quotes")
                    or data.get("data")
                    or []
                )
                print(f"  Chaves do JSON: {list(data.keys())}")
            else:
                items = []

            print(f"  Total de contratos: {len(items)}")

            if items:
                # Mostra o primeiro e o último contrato
                first = items[0]
                last  = items[-1]
                print(f"  Campos disponíveis: {list(first.keys())}")
                print(f"  Primeiro contrato: {first}")
                print(f"  Último contrato:   {last}")
        else:
            print(f"  Erro: {r.text[:200]}")
    except Exception as e:
        print(f"  Falha: {e}")

