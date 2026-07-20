"""
Diagnóstico do Coal API2 (NYMEX-MTF1!) no TradingView.
A URL é válida (a commodity existe), mas o scraper retornou 0 linhas.
Este script investiga POR QUÊ: timing, layout de tabela, ou JS tardio.
"""
import re
from playwright.sync_api import sync_playwright

URL = "https://www.tradingview.com/symbols/NYMEX-MTF1!/contracts/"

with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=True,
        args=["--no-sandbox", "--disable-dev-shm-usage",
              "--disable-blink-features=AutomationControlled"],
    )
    ctx = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36",
        viewport={"width": 1280, "height": 900}, locale="en-US",
    )
    home = ctx.new_page()
    try:
        home.goto("https://www.tradingview.com/", wait_until="domcontentloaded", timeout=20000)
        home.wait_for_timeout(2000)
    except Exception: pass
    home.close()

    page = ctx.new_page()
    resp = page.goto(URL, wait_until="networkidle", timeout=50000)
    print(f"HTTP: {resp.status if resp else '?'} | título: {page.title()[:70]}")

    # Espera progressiva — a tabela pode aparecer tarde numa página ilíquida
    for wait_s in [4, 8, 15]:
        page.wait_for_timeout((wait_s - (4 if wait_s==4 else 0)) * 1000)
        n_tbody = len(page.query_selector_all("table tbody tr"))
        n_any_tr = len(page.query_selector_all("tr"))
        n_tables = len(page.query_selector_all("table"))
        print(f"  após ~{wait_s}s: tabelas={n_tables}  tbody<tr>={n_tbody}  qualquer<tr>={n_any_tr}")
        if n_tbody > 0:
            print("  >>> TABELA ENCONTRADA")
            break

    # Se ainda vazio, dump de diagnóstico
    rows = page.query_selector_all("table tbody tr")
    if not rows:
        print("\n  Tabela vazia. Investigando estrutura alternativa...")
        # procura preços no HTML bruto
        html = page.content()
        prices = re.findall(r'\b\d{2,3}\.\d{2,4}\b', html)
        print(f"  Números tipo-preço no HTML: {prices[:15]}")
        # procura divs que possam ser linhas (TradingView às vezes usa div-grid)
        for sel in ["[class*='row']", "[class*='contract']", "[data-rowkey]", "[role='row']"]:
            n = len(page.query_selector_all(sel))
            print(f"  seletor {sel}: {n} elementos")
        page.screenshot(path="coal_debug.png", full_page=True)
        print("  screenshot: coal_debug.png")
    else:
        print(f"\n  {len(rows)} linhas. Amostra:")
        for row in rows[:5]:
            cells = row.query_selector_all("td")
            texts = [c.inner_text().strip().replace('\n',' ') for c in cells[:4]]
            print(f"    {texts}")

    browser.close()
print("\n=== fim ===")
