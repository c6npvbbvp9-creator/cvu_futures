"""
Cotações Diárias — aba resumo com 1 linha por dia e 1 coluna por ticker.
Datas mais recentes aparecem primeiro.

Estrutura da aba "Cotações Diárias":
  Data | HH | Brent | NBP | JKM | TTF

Uso:
  python daily_quotes.py --backfill   → popula histórico de junho
  python daily_quotes.py              → adiciona só o dia de hoje
"""

import sys, time, json, urllib.request
from datetime import date, datetime, timezone
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

EXCEL_PATH = "data/CME_Futures_Database.xlsx"
SHEET_NAME = "Cotações Diárias"

# Coal_API2 removido conforme solicitado
TICKERS = ["HH", "Brent", "NBP", "JKM", "TTF"]

YAHOO_SYMBOLS = {
    "HH":    "NG=F",
    "Brent": "BZ=F",
    "NBP":   "UKG=F",
    "JKM":   "JKM=F",
    "TTF":   "TTF=F",
}

HEADER_BG  = "2E4057"
HEADER_FG  = "FFFFFF"
ALT_ROW    = "EEF2F7"
BORDER_CLR = "B8C4CE"


# ---------------------------------------------------------------------------
# Formatação Excel
# ---------------------------------------------------------------------------

def _border():
    s = Side(style="thin", color=BORDER_CLR)
    return Border(left=s, right=s, top=s, bottom=s)


def _setup_sheet(ws):
    headers = ["Data"] + TICKERS
    widths  = [14] + [14] * len(TICKERS)
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font      = Font(name="Arial", bold=True, color=HEADER_FG, size=11)
        c.fill      = PatternFill("solid", start_color=HEADER_BG)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = _border()
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"
    ws.sheet_properties.tabColor = "2E86AB"


def _style_row(ws, row_idx):
    fill = ALT_ROW if row_idx % 2 == 0 else "FFFFFF"
    for col in range(1, len(TICKERS) + 2):
        c = ws.cell(row=row_idx, column=col)
        c.fill      = PatternFill("solid", start_color=fill)
        c.font      = Font(name="Arial", size=10)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = _border()


# ---------------------------------------------------------------------------
# Coleta de preços
# ---------------------------------------------------------------------------

def _fetch_yahoo(symbol: str, date_from: str, date_to: str) -> dict:
    """Busca histórico via API v8 do Yahoo Finance. Retorna {date_str: price_str}."""
    p1 = int(datetime.strptime(date_from, "%Y-%m-%d").replace(tzinfo=timezone.utc).timestamp())
    p2 = int(datetime.strptime(date_to, "%Y-%m-%d").replace(tzinfo=timezone.utc).timestamp()) + 86400

    for host in ["query1", "query2"]:
        url = (f"https://{host}.finance.yahoo.com/v8/finance/chart/{symbol}"
               f"?interval=1d&period1={p1}&period2={p2}&includePrePost=false")
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Accept": "application/json",
            })
            raw  = urllib.request.urlopen(req, timeout=20).read()
            data = json.loads(raw)
            r    = data["chart"]["result"][0]
            timestamps = r.get("timestamp", [])
            closes     = r["indicators"]["quote"][0].get("close", [])
            result = {}
            for t, c in zip(timestamps, closes):
                if c is None:
                    continue
                d = datetime.fromtimestamp(int(t), tz=timezone.utc).date().isoformat()
                result[d] = f"{float(c):.4f}"
            return result
        except Exception as e:
            print(f"    [{host}] ERRO: {e}")
            time.sleep(1)
    return {}


def _get_today_from_excel() -> dict:
    """Lê o primeiro contrato disponível de hoje no Excel (front month)."""
    prices = {}
    today  = date.today().isoformat()
    try:
        wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
        for ticker in TICKERS:
            prices[ticker] = ""
            if ticker not in wb.sheetnames:
                continue
            for row in wb[ticker].iter_rows(min_row=2, values_only=True):
                if row[0] and str(row[0])[:10] == today and row[3]:
                    prices[ticker] = str(row[3])
                    break
        wb.close()
    except Exception as e:
        print(f"  [AVISO] {e}")
    return prices


# ---------------------------------------------------------------------------
# Atualização do Excel — reconstrói a aba com datas mais recentes primeiro
# ---------------------------------------------------------------------------

def _rebuild_sheet(wb, prices_by_date: dict):
    """
    Reconstrói a aba "Cotações Diárias" do zero,
    ordenando as datas da mais recente para a mais antiga.
    """
    # Remove aba existente se houver
    if SHEET_NAME in wb.sheetnames:
        del wb[SHEET_NAME]

    # Cria nova aba na primeira posição
    ws = wb.create_sheet(title=SHEET_NAME, index=0)
    _setup_sheet(ws)

    # Ordena datas do mais recente ao mais antigo
    for row_idx, d in enumerate(sorted(prices_by_date.keys(), reverse=True), start=2):
        prices = prices_by_date[d]
        ws.cell(row_idx, 1, d)
        for i, ticker in enumerate(TICKERS, 2):
            ws.cell(row_idx, i, prices.get(ticker, ""))
        _style_row(ws, row_idx)

    return ws


def update_daily_quotes(new_prices: dict):
    """
    Atualiza a aba de cotações diárias.
    new_prices: {date_str: {ticker: price_str}}
    """
    wb = load_workbook(EXCEL_PATH)

    # Lê dados existentes da aba
    existing = {}
    if SHEET_NAME in wb.sheetnames:
        ws_old = wb[SHEET_NAME]
        for row in ws_old.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            d = str(row[0])[:10]
            existing[d] = {}
            for i, ticker in enumerate(TICKERS):
                val = row[i + 1] if len(row) > i + 1 else ""
                existing[d][ticker] = str(val) if val not in (None, "") else ""

    # Merge: novos dados sobrepõem existentes
    merged = {**existing}
    added = 0
    for d, prices in new_prices.items():
        if d not in merged:
            merged[d] = prices
            added += 1
        else:
            # Atualiza células vazias com novos dados
            for ticker, price in prices.items():
                if price and not merged[d].get(ticker):
                    merged[d][ticker] = price

    # Reconstrói aba ordenada (mais recente primeiro)
    _rebuild_sheet(wb, merged)

    # Garante que outras abas não mudaram de posição
    wb.save(EXCEL_PATH)
    print(f"  [Cotações Diárias] {added} dias novos adicionados ({len(merged)} total)")


# ---------------------------------------------------------------------------
# Modos de execução
# ---------------------------------------------------------------------------

def run_today():
    """Adiciona o dia de hoje usando dados já coletados no Excel."""
    print("\n[Cotações Diárias] Atualizando com dados de hoje...")
    today  = date.today().isoformat()
    prices = _get_today_from_excel()
    print(f"  Preços: {prices}")
    update_daily_quotes({today: prices})


def run_backfill(date_from="2026-06-01", date_to=None):
    """Popula histórico via Yahoo Finance."""
    if date_to is None:
        date_to = date.today().isoformat()
    print(f"\n[Cotações Diárias] Backfill {date_from} → {date_to} via Yahoo Finance...\n")

    all_history = {ticker: {} for ticker in TICKERS}

    for ticker, sym in YAHOO_SYMBOLS.items():
        print(f"  {ticker} ({sym})...")
        history = _fetch_yahoo(sym, date_from, date_to)
        filtered = {d: v for d, v in history.items() if date_from <= d <= date_to}
        all_history[ticker] = filtered
        print(f"    {len(filtered)} dias coletados", end="")
        if filtered:
            print(f" | {min(filtered)}: {filtered[min(filtered)]} → {max(filtered)}: {filtered[max(filtered)]}")
        else:
            print()
        time.sleep(1)

    # Reorganiza por data
    all_dates = set()
    for h in all_history.values():
        all_dates.update(h.keys())

    prices_by_date = {
        d: {ticker: all_history[ticker].get(d, "") for ticker in TICKERS}
        for d in sorted(all_dates)
    }

    print(f"\n  {len(prices_by_date)} dias coletados no total")
    if prices_by_date:
        update_daily_quotes(prices_by_date)
    else:
        print("  Nenhum dado — verifique a conexão com o Yahoo Finance")


if __name__ == "__main__":
    if "--backfill" in sys.argv:
        run_backfill()
    else:
        run_today()
