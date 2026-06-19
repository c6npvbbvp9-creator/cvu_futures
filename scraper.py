"""
CME Futures Scraper — TradingView via Playwright (browser headless)
Usa Chromium real para contornar o bloqueio de IP do CME.

Fonte: https://www.tradingview.com/symbols/NYMEX-NG1!/contracts/
O TradingView expõe a curva futura completa de cada ticker via sua API
interna (scanner endpoint) que é capturada durante o carregamento da página.

Fallback para TTF: contratos sem preço usam média dos últimos 12 meses.
"""

import json
import os
import time
from datetime import date
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

EXCEL_PATH = "data/CME_Futures_Database.xlsx"

# Mapeamento ticker → símbolo TradingView + URL de contratos
TICKERS = {
    "HH": {
        "name": "Henry Hub Natural Gas",
        "tv_symbol": "NYMEX:NG1!",
        "tv_url": "https://www.tradingview.com/symbols/NYMEX-NG1!/contracts/",
        "use_fallback": False,
    },
    "Brent": {
        "name": "Brent Crude Oil",
        "tv_symbol": "NYMEX:BB1!",
        "tv_url": "https://www.tradingview.com/symbols/NYMEX-BB1!/contracts/",
        "use_fallback": False,
    },
    "NBP": {
        "name": "UK NBP Natural Gas Calendar Month",
        "tv_symbol": "NYMEX:UKG1!",
        "tv_url": "https://www.tradingview.com/symbols/NYMEX-UKG1!/contracts/",
        "use_fallback": False,
    },
    "JKM": {
        "name": "LNG Japan Korea Marker (Platts)",
        "tv_symbol": "NYMEX:LNG1!",
        "tv_url": "https://www.tradingview.com/symbols/NYMEX-LNG1!/contracts/",
        "use_fallback": False,
    },
    "TTF": {
        "name": "Dutch TTF Natural Gas (Platts ENDEX)",
        "tv_symbol": "NYMEX:TTF1!",
        "tv_url": "https://www.tradingview.com/symbols/NYMEX-TTF1!/contracts/",
        "use_fallback": True,
    },
    "Coal_API2": {
        "name": "Coal API2 CIF ARA (Argus/McCloskey)",
        "tv_symbol": "NYMEX:MTF1!",
        "tv_url": "https://www.tradingview.com/symbols/NYMEX-MTF1!/contracts/",
        "use_fallback": False,
    },
}

INVALID_PRICES = {"-", "0", "0.00", "", "N/A", "n/a", "—", "–"}


# ---------------------------------------------------------------------------
# Fallback: média dos últimos 12 meses (GT CVU Estrutural, Equação 17)
# ---------------------------------------------------------------------------

def _load_historical(ticker: str) -> dict:
    if not os.path.exists(EXCEL_PATH):
        return {}
    try:
        wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    except Exception:
        return {}
    if ticker not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb[ticker]
    history = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0] or not row[1] or not row[3]:
            continue
        try:
            price = float(str(row[3]).replace(",", "."))
        except ValueError:
            continue
        history.setdefault(str(row[1]).strip(), []).append((str(row[0]), price))
    wb.close()
    return {c: [p for _, p in sorted(v)] for c, v in history.items()}


def _fallback_price(ticker: str, contract: str) -> tuple:
    prices = _load_historical(ticker).get(contract, [])
    if not prices:
        return "", ""
    last_12 = prices[-12:]
    avg = sum(last_12) / len(last_12)
    return f"{avg:.4f}", f"CONTINGÊNCIA (média {len(last_12)} meses)"


# ---------------------------------------------------------------------------
# Coleta via TradingView (Playwright)
# ---------------------------------------------------------------------------

def _fetch_with_playwright(ticker: str, config: dict, page) -> list:
    """
    Carrega a página de contratos do TradingView e intercepta a chamada
    da API interna que retorna todos os contratos futuros com preços.
    """
    today = date.today().isoformat()
    captured = []

    def handle_response(response):
        # TradingView usa endpoint /scan para popular a tabela de contratos
        if "scan" in response.url and response.status == 200:
            try:
                data = response.json()
                if "data" in data:
                    captured.extend(data["data"])
            except Exception:
                pass

    page.on("response", handle_response)

    try:
        page.goto(config["tv_url"], wait_until="networkidle", timeout=45000)
        page.wait_for_timeout(3000)  # aguarda carregamento completo da tabela
    except Exception as e:
        print(f"  [AVISO] {ticker}: timeout/erro ao carregar — {e}")

    page.remove_listener("response", handle_response)

    if not captured:
        # Fallback: tenta extrair diretamente do HTML da tabela
        return _parse_html_table(ticker, config, page, today)

    results = []
    ok = fallback = 0

    for item in captured:
        symbol = item.get("s", "")        # ex: "NYMEX:NGN2026"
        values = item.get("d", [])        # [close, expiration_date, description]

        if not symbol or not values:
            continue

        # Extrai mês/ano do símbolo (ex: NGN2026 → "JUL 26")
        contract = _symbol_to_contract(symbol)
        if not contract:
            continue

        price_raw = values[0] if values else None
        expiry    = values[1] if len(values) > 1 else ""

        price_str = str(price_raw).strip() if price_raw is not None else ""
        has_price = price_str and price_str not in INVALID_PRICES

        if has_price:
            results.append({
                "collection_date": today,
                "ticker": ticker,
                "contract": contract,
                "expiry_date": str(expiry),
                "settlement_price": price_str,
                "source": "CME via TradingView",
            })
            ok += 1
        elif config["use_fallback"]:
            fp, nota = _fallback_price(ticker, contract)
            if fp:
                results.append({
                    "collection_date": today,
                    "ticker": ticker,
                    "contract": contract,
                    "expiry_date": str(expiry),
                    "settlement_price": fp,
                    "source": nota,
                })
                fallback += 1

    if config["use_fallback"] and fallback:
        print(f"  [OK] {ticker}: {ok} contratos + {fallback} contingência")
    else:
        print(f"  [OK] {ticker}: {ok} contratos")

    return results


def _parse_html_table(ticker: str, config: dict, page, today: str) -> list:
    """
    Se a API não foi capturada, extrai os dados da tabela HTML diretamente.
    """
    results = []
    try:
        rows = page.query_selector_all("table tbody tr")
        for row in rows:
            cells = row.query_selector_all("td")
            if len(cells) < 2:
                continue
            symbol_text = cells[0].inner_text().strip()
            price_text  = cells[1].inner_text().strip()
            expiry_text = cells[2].inner_text().strip() if len(cells) > 2 else ""

            contract = _symbol_to_contract(symbol_text) or symbol_text
            if not price_text or price_text in INVALID_PRICES:
                continue

            results.append({
                "collection_date": today,
                "ticker": ticker,
                "contract": contract,
                "expiry_date": expiry_text,
                "settlement_price": price_text,
                "source": "CME via TradingView",
            })
    except Exception as e:
        print(f"  [AVISO] {ticker}: erro ao parsear HTML — {e}")
    return results


# Mês CME code → nome abreviado
_MONTH_CODES = {
    "F": "JAN", "G": "FEB", "H": "MAR", "J": "APR",
    "K": "MAY", "M": "JUN", "N": "JUL", "Q": "AUG",
    "U": "SEP", "V": "OCT", "X": "NOV", "Z": "DEC",
}

def _symbol_to_contract(symbol: str) -> str:
    """
    Converte símbolo TradingView/CME para label legível.
    Ex: "NYMEX:NGN2026" → "JUL 26"
        "NGN2026"        → "JUL 26"
    """
    # Remove prefixo de exchange
    s = symbol.split(":")[-1]  # "NGN2026"

    # Remove sufixo numérico do root (NG, BB, UKG, LNG, TTF, MTF)
    import re
    m = re.search(r"([A-Z]+)([FGHJKMNQUVXZ])(\d{4})$", s)
    if not m:
        return ""
    month_code = m.group(2)
    year       = m.group(3)  # "2026"
    month_name = _MONTH_CODES.get(month_code, "")
    if not month_name:
        return ""
    return f"{month_name} {year[2:]}"   # "JUL 26"


# ---------------------------------------------------------------------------
# Coleta principal
# ---------------------------------------------------------------------------

def collect_all() -> dict:
    print(f"\n=== Coleta CME Futures — {date.today().isoformat()} ===\n")

    all_data = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage"],
        )
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1280, "height": 800},
            locale="en-US",
        )

        for ticker, config in TICKERS.items():
            print(f"Buscando {ticker} ({config['name']})...")
            page = context.new_page()
            try:
                rows = _fetch_with_playwright(ticker, config, page)
                all_data[ticker] = rows
            except Exception as e:
                print(f"  [ERRO] {ticker}: {e}")
                all_data[ticker] = []
            finally:
                page.close()
            time.sleep(2)

        browser.close()

    total = sum(len(v) for v in all_data.values())
    print(f"\nTotal de contratos coletados: {total}")
    return all_data


if __name__ == "__main__":
    from update_excel import update_excel
    data = collect_all()
    update_excel(data)
