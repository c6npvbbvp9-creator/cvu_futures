"""
Debug: inspeciona a estrutura real da página do TradingView no GitHub Actions.
Salva screenshot + HTML + APIs capturadas para análise.
"""

import os, json, re
from playwright.sync_api import sync_playwright

os.makedirs("debug_output", exist_ok=True)

URL = "https://www.tradingview.com/symbols/NYMEX-NG1!/contracts/"

with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=True,
        args=["--no-sandbox", "--disable-dev-shm-usage",
              "--disable-blink-features=AutomationControlled"],
    )
    ctx = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36",
        viewport={"width": 1280, "height": 900},
        locale="en-US",
    )

    page = ctx.new_page()

    # Captura TODAS as chamadas de rede
    api_calls = []
    json_bodies = []

    def on_response(resp):
        api_calls.append(f"[{resp.status}] {resp.url[:150]}")
        try:
            ct = resp.headers.get("content-type","")
            if "json" in ct and resp.status == 200:
                body = resp.json()
                body_str = json.dumps(body)[:2000]
                json_bodies.append({"url": resp.url[:150], "body": body_str})
        except Exception:
            pass

    page.on("response", on_response)

    # Primeiro visita a home
    page.goto("https://www.tradingview.com/", wait_until="domcontentloaded", timeout=20000)
    page.wait_for_timeout(2000)

    # Agora vai para a página de contratos
    page.goto(URL, wait_until="networkidle", timeout=50000)
    page.wait_for_timeout(5000)

    # Screenshot
    page.screenshot(path="debug_output/tv_screenshot.png", full_page=True)
    print("Screenshot salvo")

    # HTML completo
    html = page.content()
    with open("debug_output/tv_page.html", "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTML salvo: {len(html)} chars")

    # APIs relevantes
    relevant = [x for x in api_calls if any(k in x for k in 
        ["scan","contract","quote","symbol","price","futures","data","api"])]
    print(f"\n=== {len(relevant)} APIs relevantes ===")
    for r in relevant[:30]:
        print(r)

    # JSONs capturados
    print(f"\n=== {len(json_bodies)} JSONs capturados ===")
    for jb in json_bodies[:10]:
        print(f"\nURL: {jb['url']}")
        print(f"Body: {jb['body'][:500]}")

    # Estrutura da página
    print("\n=== Estrutura DOM ===")
    print(f"<table> encontradas: {len(page.query_selector_all('table'))}")
    print(f"<tr> encontradas: {len(page.query_selector_all('tr'))}")
    print(f"<td> encontradas: {len(page.query_selector_all('td'))}")

    # Procura números que parecem preços
    prices_found = re.findall(r'\b\d+\.\d{3,4}\b', html)
    print(f"Números tipo preço no HTML: {prices_found[:20]}")

    # __NEXT_DATA__
    next_data = page.query_selector("#__NEXT_DATA__")
    if next_data:
        nd_text = next_data.inner_text()
        print(f"\n__NEXT_DATA__ encontrado: {len(nd_text)} chars")
        with open("debug_output/next_data.json", "w") as f:
            f.write(nd_text)
    else:
        print("\n__NEXT_DATA__: não encontrado")

    # Salva todas as APIs
    with open("debug_output/all_apis.txt", "w") as f:
        f.write("\n".join(api_calls))
    
    with open("debug_output/json_responses.json", "w") as f:
        json.dump(json_bodies, f, indent=2)

    browser.close()

print("\nDebug completo. Arquivos em debug_output/")
