"""
CME Futures Scraper — TradingView via Playwright
"""

import os, re, time
from datetime import date
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

EXCEL_PATH = "data/CME_Futures_Database.xlsx"

# URLs corrigidas baseadas no log real do TradingView
TICKERS = {
    "HH": {
        "name": "Henry Hub Natural Gas",
        "tv_url": "https://www.tradingview.com/symbols/NYMEX-NG1!/contracts/",
        "root": "NG",
        "use_fallback": False,
    },
    "Brent": {
        "name": "Brent Crude Oil",
        "tv_url": "https://www.tradingview.com/symbols/NYMEX-BB1!/contracts/",
        "root": "BB",
        "use_fallback": False,
    },
    "NBP": {
        "name": "UK NBP Natural Gas Calendar Month",
        "tv_url": "https://www.tradingview.com/symbols/NYMEX-UKG1!/contracts/",
        "root": "UKG",
        "use_fallback": False,
    },
    "JKM": {
        "name": "LNG Japan Korea Marker (Platts)",
        "tv_url": "https://www.tradingview.com/symbols/NYMEX-JKM1!/contracts/",
        "root": "JKM",
        "use_fallback": False,
    },
    "TTF": {
        "name": "Dutch TTF Natural Gas Calendar Month",
        "tv_url": "https://www.tradingview.com/symbols/NYMEX-TTF1!/contracts/",
        "root": "TTF",
        "use_fallback": False,
    },
    "Coal_API2": {
        "name": "Coal API2 CIF ARA (Argus/McCloskey)",
        "tv_url": "https://www.tradingview.com/symbols/NYMEX-MTF1!/contracts/",
        "root": "MTF",
        "use_fallback": False,
    },
}

MONTH_CODES = {
    "F":"JAN","G":"FEB","H":"MAR","J":"APR","K":"MAY","M":"JUN",
    "N":"JUL","Q":"AUG","U":"SEP","V":"OCT","X":"NOV","Z":"DEC",
}

INVALID = {"-","","N/A","n/a","—","–","0","0.00","unch","UNCH","null","None"}


def _symbol_to_label(sym):
    s = sym.split(":")[-1]
    m = re.search(r"[A-Z]+([FGHJKMNQUVXZ])(\d{4})$", s)
    if not m:
        return ""
    month = MONTH_CODES.get(m.group(1), "")
    return f"{month} {m.group(2)[2:]}" if month else ""


def _scrape_ticker(ticker, config, ctx):
    today = date.today().isoformat()
    results = []
    page = ctx.new_page()

    try:
        page.goto(config["tv_url"], wait_until="networkidle", timeout=50000)
        page.wait_for_timeout(4000)

        rows = page.query_selector_all("table tbody tr")
        print(f"    <tr> encontradas: {len(rows)}")

        for row in rows:
            cells = row.query_selector_all("td")
            if len(cells) < 3:
                continue
            texts = [c.inner_text().strip() for c in cells]

            # Célula 0: símbolo+nome, Célula 1: data vencimento, Célula 2: preço
            sym_raw  = texts[0].split("\n")[0].strip()   # ex: "NGN2026"
            expiry   = texts[1].strip()                   # ex: "2026-06-26"
            price    = texts[2].strip()                   # ex: "3.198"

            label = _symbol_to_label(sym_raw)
            if not label:
                continue
            if not price or price in INVALID:
                continue
            # Valida que é número
            try:
                float(price.replace(",", "."))
            except ValueError:
                continue

            results.append({
                "collection_date": today,
                "ticker": ticker,
                "contract": label,
                "expiry_date": expiry,
                "settlement_price": price,
                "source": "CME via TradingView",
            })

    except Exception as e:
        print(f"    [ERRO] {e}")
    finally:
        page.close()

    print(f"    Contratos extraídos: {len(results)}")
    if results:
        print(f"    Exemplo: {results[0]}")
    return results


def collect_all():
    print(f"\n=== Coleta CME Futures — {date.today().isoformat()} ===\n")
    all_data = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage",
                  "--disable-blink-features=AutomationControlled"],
        )
        ctx = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 900},
            locale="en-US",
        )

        home = ctx.new_page()
        try:
            home.goto("https://www.tradingview.com/", wait_until="domcontentloaded", timeout=20000)
            home.wait_for_timeout(2000)
        except Exception:
            pass
        home.close()

        for ticker, config in TICKERS.items():
            print(f"Buscando {ticker} ({config['name']})...")
            try:
                rows = _scrape_ticker(ticker, config, ctx)
                all_data[ticker] = rows
            except Exception as e:
                print(f"  [ERRO] {ticker}: {e}")
                all_data[ticker] = []
            time.sleep(3)

        browser.close()

    total = sum(len(v) for v in all_data.values())
    print(f"\nTotal de contratos coletados: {total}")
    return all_data


if __name__ == "__main__":
    from update_excel import update_excel
    data = collect_all()
    update_excel(data)
