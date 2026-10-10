# ================================================================
# NSE STOCK MONTH / DAY ANALYSIS - STREAMLIT VERSION (PRO 3)
# Free hosting: Streamlit Community Cloud
# ================================================================
# Local run:
#   pip install -r requirements.txt
#   streamlit run app.py
#
# Features
#   - One-click NIFTY 50 / BANK NIFTY / NSE F&O stocks
#   - Search by code / company, sector filter, custom symbols
#   - Custom Month-YY range (Jan-26 -> Oct-26) + YTD / 6M / 12M / 24M
#   - Month wise table, Day wise table, Month HEATMAP
#   - COLUMN MANAGER (per user): show / hide, move to FRONT, move to BACK,
#     Compact / Full layouts, save & load layout file
#   - Excel / CSV export (exports exactly the layout you chose)
# ================================================================

import json
import re
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from io import BytesIO, StringIO
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components
import yfinance as yf

APP_NAME = "NSE Stock Month / Day Analysis"
DEVELOPER = "Venkatesan V"

st.set_page_config(page_title=APP_NAME, page_icon="📊", layout="wide",
                   initial_sidebar_state="expanded")

CACHE_DIR = Path(tempfile.gettempdir()) / "NSE_CACHE"
CACHE_DIR.mkdir(parents=True, exist_ok=True)

NSE_EQUITY_URL = "https://nsearchives.nseindia.com/content/equities/EQUITY_L.csv"
NSE_FNO_URL = "https://nsearchives.nseindia.com/content/fo/fo_mktlots.csv"
NIFTY50_URL = "https://nsearchives.nseindia.com/content/indices/ind_nifty50list.csv"
BANKNIFTY_URL = "https://nsearchives.nseindia.com/content/indices/ind_niftybanklist.csv"

NSE_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/",
}

MAX_WORKERS = 4          # free hosting: be gentle with Yahoo Finance
PRICE_TTL = 30 * 60      # seconds a downloaded price series stays cached

GREEN, RED = "#22c55e", "#ef4444"

# ================================================================
# FALLBACK LISTS (used when NSE blocks the hosting server)
# ================================================================

NIFTY_50_FALLBACK = {
    "ADANIENT", "ADANIPORTS", "APOLLOHOSP", "ASIANPAINT", "AXISBANK",
    "BAJAJ-AUTO", "BAJFINANCE", "BAJAJFINSV", "BEL", "BHARTIARTL",
    "CIPLA", "COALINDIA", "DRREDDY", "EICHERMOT", "ETERNAL", "GRASIM",
    "HCLTECH", "HDFCBANK", "HDFCLIFE", "HEROMOTOCO", "HINDALCO",
    "HINDUNILVR", "ICICIBANK", "INDUSINDBK", "INFY", "ITC", "JIOFIN",
    "JSWSTEEL", "KOTAKBANK", "LT", "M&M", "MARUTI", "MAXHEALTH",
    "NESTLEIND", "NTPC", "ONGC", "POWERGRID", "RELIANCE", "SBILIFE",
    "SBIN", "SHRIRAMFIN", "SUNPHARMA", "TATACONSUM", "TATASTEEL", "TCS",
    "TECHM", "TITAN", "TRENT", "ULTRACEMCO", "WIPRO",
}
BANK_NIFTY_FALLBACK = {
    "AUBANK", "AXISBANK", "BANKBARODA", "CANBK", "FEDERALBNK", "HDFCBANK",
    "ICICIBANK", "IDFCFIRSTB", "INDUSINDBK", "KOTAKBANK", "PNB", "SBIN",
}
FNO_FALLBACK = """
360ONE ABB ABCAPITAL ABFRL ACC ADANIENSOL ADANIENT ADANIGREEN ADANIPORTS
ADANIPOWER ALKEM AMBER AMBUJACEM ANGELONE APLAPOLLO APOLLOHOSP ASHOKLEY
ASIANPAINT ASTRAL AUBANK AUROPHARMA AXISBANK BAJAJ-AUTO BAJAJFINSV
BAJFINANCE BANDHANBNK BANKBARODA BANKINDIA BDL BEL BHARATFORG BHARTIARTL
BHEL BIOCON BLUESTARCO BOSCHLTD BPCL BRITANNIA BSE CAMS CANBK CDSL
CGPOWER CHOLAFIN CIPLA COALINDIA COFORGE COLPAL CONCOR CROMPTON
CUMMINSIND CYIENT DABUR DALBHARAT DELHIVERY DIVISLAB DIXON DLF DMART
DRREDDY EICHERMOT ETERNAL EXIDEIND FEDERALBNK FORTIS GAIL GLENMARK
GMRAIRPORT GODREJCP GODREJPROP GRANULES GRASIM HAL HAVELLS HCLTECH
HDFCAMC HDFCBANK HDFCLIFE HEROMOTOCO HFCL HINDALCO HINDPETRO HINDUNILVR
HINDZINC HUDCO ICICIBANK ICICIGI ICICIPRULI IDEA IDFCFIRSTB IEX IGL IIFL
INDHOTEL INDIANB INDIGO INDUSINDBK INDUSTOWER INFY INOXWIND IOC IRCTC
IREDA IRFC ITC JINDALSTEL JIOFIN JSWENERGY JSWSTEEL JUBLFOOD KALYANKJIL
KAYNES KEI KFINTECH KOTAKBANK KPITTECH LAURUSLABS LICHSGFIN LICI LODHA LT
LTF LTIM LUPIN M&M M&MFIN MANAPPURAM MANKIND MARICO MARUTI MAXHEALTH
MAZDOCK MCX MFSL MOTHERSON MPHASIS MUTHOOTFIN NATIONALUM NAUKRI NBCC NCC
NESTLEIND NHPC NMDC NTPC NUVAMA NYKAA OBEROIRLTY OFSS OIL ONGC PAGEIND
PATANJALI PAYTM PERSISTENT PETRONET PFC PGEL PHOENIXLTD PIDILITIND PIIND
PNB PNBHOUSING POLICYBZR POLYCAB POWERGRID POWERINDIA PPLPHARMA PRESTIGE
RBLBANK RECLTD RELIANCE RVNL SAIL SAMMAANCAP SBICARD SBILIFE SBIN
SHREECEM SHRIRAMFIN SIEMENS SOLARINDS SONACOMS SRF SUNPHARMA SUPREMEIND
SUZLON SYNGENE TATACONSUM TATAELXSI TATAPOWER TATASTEEL TATATECH TCS
TECHM TIINDIA TITAGARH TITAN TORNTPHARM TORNTPOWER TRENT TVSMOTOR
ULTRACEMCO UNIONBANK UNITDSPR UNOMINDA UPL VBL VEDL VOLTAS WIPRO YESBANK
ZYDUSLIFE
"""
FNO_INDEX_NAMES = {"NIFTY", "BANKNIFTY", "FINNIFTY", "MIDCPNIFTY", "NIFTYNXT50"}

SECTOR_GROUPS = {
    "IT": "TCS INFY HCLTECH WIPRO TECHM LTIM MPHASIS PERSISTENT COFORGE OFSS "
          "KPITTECH TATAELXSI CYIENT LTTS TATATECH",
    "Banking": "HDFCBANK ICICIBANK SBIN AXISBANK KOTAKBANK INDUSINDBK "
               "BANKBARODA PNB CANBK UNIONBANK FEDERALBNK IDFCFIRSTB AUBANK "
               "BANDHANBNK RBLBANK YESBANK INDIANB BANKINDIA IDBI",
    "Financial Services": "BAJFINANCE BAJAJFINSV CHOLAFIN SHRIRAMFIN MUTHOOTFIN "
                          "MANAPPURAM LICHSGFIN PFC RECLTD M&MFIN SBICARD "
                          "HDFCAMC JIOFIN IRFC POONAWALLA ANGELONE CDSL BSE MCX "
                          "360ONE NUVAMA PNBHOUSING IIFL KFINTECH CAMS PAYTM "
                          "POLICYBZR LTF IREDA HUDCO SAMMAANCAP ABCAPITAL MFSL",
    "Insurance": "HDFCLIFE SBILIFE ICICIPRULI ICICIGI LICI",
    "Pharma & Healthcare": "SUNPHARMA CIPLA DRREDDY DIVISLAB LUPIN AUROPHARMA "
                           "TORNTPHARM ZYDUSLIFE ALKEM BIOCON GLENMARK MANKIND "
                           "LAURUSLABS GRANULES APOLLOHOSP MAXHEALTH FORTIS "
                           "SYNGENE PPLPHARMA ABBOTINDIA IPCALAB",
    "Automobile": "MARUTI M&M TATAMOTORS TMPV TMCV EICHERMOT HEROMOTOCO "
                  "BAJAJ-AUTO TVSMOTOR ASHOKLEY BOSCHLTD MOTHERSON BHARATFORG "
                  "EXIDEIND ARE&M BALKRISIND MRF APOLLOTYRE SONACOMS UNOMINDA "
                  "TIINDIA HYUNDAI",
    "Metals & Mining": "TATASTEEL JSWSTEEL HINDALCO VEDL JINDALSTEL SAIL NMDC "
                       "NATIONALUM HINDZINC APLAPOLLO JSL COALINDIA LLOYDSME",
    "Oil & Gas": "RELIANCE ONGC BPCL IOC GAIL HINDPETRO PETRONET OIL IGL MGL "
                 "GUJGASLTD ATGL",
    "Power": "NTPC POWERGRID TATAPOWER ADANIGREEN ADANIPOWER ADANIENSOL NHPC "
             "JSWENERGY TORNTPOWER SJVN NTPCGREEN INOXWIND SUZLON PREMIERENE "
             "CESC WAAREEENER IEX",
    "Capital Goods": "LT BHEL SIEMENS ABB CGPOWER HAVELLS POLYCAB KEI "
                     "CUMMINSIND THERMAX DIXON KAYNES BLUESTARCO VOLTAS "
                     "CROMPTON SUPREMEIND ASTRAL RVNL TITAGARH IRCON NBCC "
                     "AMBER POWERINDIA PGEL",
    "Defence": "BEL HAL BDL MAZDOCK SOLARINDS COCHINSHIP",
    "FMCG": "ITC HINDUNILVR NESTLEIND BRITANNIA DABUR MARICO GODREJCP COLPAL "
            "TATACONSUM UNITDSPR UBL VBL PATANJALI RADICO",
    "Consumer": "ASIANPAINT TITAN TRENT PAGEIND KALYANKJIL DMART JUBLFOOD NYKAA "
                "ETERNAL SWIGGY PIDILITIND BERGEPAINT INDHOTEL NAUKRI ABFRL",
    "Cement": "ULTRACEMCO GRASIM AMBUJACEM SHREECEM DALBHARAT ACC",
    "Telecom": "BHARTIARTL IDEA INDUSTOWER TATACOMM HFCL",
    "Infrastructure": "ADANIPORTS ADANIENT INDIGO CONCOR IRCTC GMRAIRPORT "
                      "DELHIVERY NCC",
    "Realty": "DLF LODHA GODREJPROP OBEROIRLTY PRESTIGE PHOENIXLTD",
    "Chemicals": "SRF UPL PIIND DEEPAKNTR NAVINFLUOR TATACHEM COROMANDEL "
                 "CHAMBLFERT",
}
SECTOR_MAP = {sym: sec for sec, syms in SECTOR_GROUPS.items()
              for sym in syms.split()}

UNIVERSE_LABELS = ["ALL NSE STOCKS", "NIFTY 50", "BANK NIFTY",
                   "NSE F&O (Options / Futures)"]

# ================================================================
# COLUMN DEFINITIONS
# ================================================================

META_COLS = ["Stock", "Company", "Sector", "Index"]
MONTH_COLS = META_COLS + [
    "Month", "Open", "High", "Low", "Close", "Change", "Change %",
    "Low → High", "Low → High %", "Low Date", "High Date",
    "Direction", "Candle", "Open vs Close", "Volume"]
DAY_COLS = META_COLS + [
    "Date", "Open", "High", "Low", "Close", "Change", "Change %",
    "Day Range", "Range %", "Direction", "Candle", "Open vs Close", "Volume"]

COMPACT = {
    "MONTH": ["Stock", "Sector", "Month", "Open", "High", "Low", "Close",
              "Change %", "Direction"],
    "DAY": ["Stock", "Sector", "Date", "Open", "High", "Low", "Close",
            "Change %", "Direction"],
}
TREND_COLS = ["Close", "Change", "Change %", "Candle", "Direction"]
DATE_COLS = ["Month", "Date", "Low Date", "High Date"]


# ================================================================
# HELPERS
# ================================================================

def safe_float(v):
    try:
        return float(v)
    except Exception:
        return 0.0


def ym_label(ym):
    return date(ym // 100, ym % 100, 1).strftime("%b-%y")


def ym_add(ym, delta):
    n = (ym // 100) * 12 + (ym % 100 - 1) + delta
    return (n // 12) * 100 + (n % 12) + 1


def ym_to_dates(a, b):
    y1, m1 = divmod(a, 100)
    y2, m2 = divmod(b, 100)
    nxt = date(y2 + (1 if m2 == 12 else 0), 1 if m2 == 12 else m2 + 1, 1)
    return date(y1, m1, 1), nxt - timedelta(days=1)


def sector_for(symbol, company):
    if symbol in SECTOR_MAP:
        return SECTOR_MAP[symbol]
    name = str(company).lower()
    rules = [
        ("IT", ["software", "technologies", "informatics", "computer"]),
        ("Banking", ["bank"]),
        ("Pharma & Healthcare", ["pharma", "drug", "hospital", "healthcare",
                                 "laborator", "life science"]),
        ("Automobile", ["motor", "automobile", "tyres", "tire", "auto "]),
        ("Cement", ["cement"]),
        ("Telecom", ["telecom"]),
        ("Power", ["power", "electricity", "energy"]),
        ("Oil & Gas", ["petroleum", "refiner", "gas "]),
        ("Metals & Mining", ["steel", "metal", "aluminium", "aluminum",
                             "mining"]),
        ("FMCG", ["foods", "beverages", "consumer"]),
        ("Chemicals", ["chemical", "fertili", "pesticid"]),
        ("Realty", ["realty", "properties", "developers"]),
        ("Financial Services", ["finance", "financial", "capital market",
                                "securities", "housing fin"]),
        ("Textiles", ["textile", "garment", "fabrics", "spinning"]),
    ]
    for sector, words in rules:
        if any(w in name for w in words):
            return sector
    return "Other"


# ================================================================
# NSE UNIVERSE  (cached 6 hours)
# ================================================================

def fetch_csv(url, cache_name):
    cache_file = CACHE_DIR / cache_name
    try:
        s = requests.Session()
        s.headers.update(NSE_HEADERS)
        try:
            s.get("https://www.nseindia.com/", timeout=10)
        except Exception:
            pass
        r = s.get(url, timeout=20)
        r.raise_for_status()
        text = r.content.decode("utf-8-sig", errors="replace")
        if len(text) < 50:
            raise ValueError("empty")
        cache_file.write_text(text, encoding="utf-8")
        return pd.read_csv(StringIO(text), skipinitialspace=True,
                           on_bad_lines="skip"), "live"
    except Exception:
        if cache_file.exists():
            try:
                return pd.read_csv(cache_file, skipinitialspace=True,
                                   on_bad_lines="skip"), "cache"
            except Exception:
                pass
    return None, "none"


def parse_symbols(df):
    if df is None or df.empty:
        return set()
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    if "SYMBOL" not in df.columns:
        return set()
    return {str(s).strip().upper() for s in df["SYMBOL"].dropna()
            if str(s).strip()}


def parse_fno(df):
    if df is None or df.empty:
        return {}
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    if "SYMBOL" not in df.columns:
        return {}
    out = {}
    for _, row in df.iterrows():
        sym = str(row["SYMBOL"]).strip().upper()
        if (not sym or sym in ("SYMBOL", "NAN") or sym in FNO_INDEX_NAMES
                or not re.fullmatch(r"[A-Z0-9&\-]+", sym)):
            continue
        try:
            v = pd.to_numeric(row.iloc[2], errors="coerce")
            out[sym] = int(v) if pd.notna(v) else 0
        except Exception:
            out[sym] = 0
    return out


@st.cache_data(ttl=6 * 3600, show_spinner="Downloading NSE lists ...")
def get_universe():
    info = {}

    n50 = parse_symbols(fetch_csv(NIFTY50_URL, "n50.csv")[0])
    if len(n50) >= 40:
        info["n50"] = "live/cache"
    else:
        n50, info["n50"] = set(NIFTY_50_FALLBACK), "built-in"

    bank = parse_symbols(fetch_csv(BANKNIFTY_URL, "bank.csv")[0])
    if len(bank) >= 8:
        info["bank"] = "live/cache"
    else:
        bank, info["bank"] = set(BANK_NIFTY_FALLBACK), "built-in"

    fno_df, fno_src = fetch_csv(NSE_FNO_URL, "fno.csv")
    fno = parse_fno(fno_df)
    if len(fno) >= 100:
        info["fno"] = fno_src
    else:
        fno, info["fno"] = {s: 0 for s in FNO_FALLBACK.split()}, "built-in"

    def make(symbol, name):
        tags = []
        if symbol in n50:
            tags.append("NIFTY 50")
        if symbol in bank:
            tags.append("BANK NIFTY")
        if symbol in fno:
            tags.append("F&O")
        return {"symbol": symbol, "name": name,
                "sector": sector_for(symbol, name),
                "index": ", ".join(tags) if tags else "NSE Equity",
                "n50": symbol in n50, "bank": symbol in bank,
                "fno": symbol in fno, "lot": fno.get(symbol, 0)}

    eq, src = fetch_csv(NSE_EQUITY_URL, "equity.csv")
    info["equity"] = src
    stocks = []
    if eq is not None:
        eq.columns = [str(c).strip().upper() for c in eq.columns]
        if "SERIES" in eq.columns:
            eq["SERIES"] = eq["SERIES"].astype(str).str.strip().str.upper()
            eq = eq[eq["SERIES"].isin(["EQ", "BE", "BZ"])]
        if "SYMBOL" in eq.columns:
            eq["SYMBOL"] = eq["SYMBOL"].astype(str).str.strip().str.upper()
            eq = eq.drop_duplicates("SYMBOL")
            has_name = "NAME OF COMPANY" in eq.columns
            for _, r in eq.iterrows():
                sym = str(r["SYMBOL"]).strip()
                nm = str(r["NAME OF COMPANY"]).strip() if has_name else sym
                stocks.append(make(sym, nm))
    if not stocks:                        # NSE blocked the server
        info["equity"] = "fallback"
        syms = sorted(set(FNO_FALLBACK.split()) | NIFTY_50_FALLBACK
                      | BANK_NIFTY_FALLBACK)
        stocks = [make(s, s) for s in syms]

    stocks.sort(key=lambda s: s["symbol"])
    return {"stocks": stocks, "info": info}


# ================================================================
# PRICE DATA
# ================================================================

@st.cache_resource
def price_store():
    return {"lock": threading.Lock(), "data": {}}


def download_daily(store, symbol, start, end, retries=3):
    key = (symbol, start, end)
    with store["lock"]:
        hit = store["data"].get(key)
    if hit and time.time() - hit[0] < PRICE_TTL:
        return hit[1].copy()

    ticker = symbol if symbol.endswith(".NS") else symbol + ".NS"
    last = None
    for attempt in range(retries):
        try:
            df = yf.Ticker(ticker).history(
                start=start.strftime("%Y-%m-%d"),
                end=(end + timedelta(days=1)).strftime("%Y-%m-%d"),
                interval="1d", auto_adjust=False, actions=False)
            if df is None or df.empty:
                raise ValueError(f"No Yahoo Finance data for {symbol}")
            df.columns = [str(c).strip().title() for c in df.columns]
            need = ["Open", "High", "Low", "Close", "Volume"]
            miss = [c for c in need if c not in df.columns]
            if miss:
                raise ValueError(f"{symbol}: missing {miss}")
            df = df[need].copy()
            df.index = pd.to_datetime(df.index)
            if df.index.tz is not None:
                df.index = df.index.tz_localize(None)
            df.index = df.index.normalize()
            df = df.dropna(subset=["Open", "High", "Low", "Close"])
            if df.empty:
                raise ValueError(f"No valid OHLC data for {symbol}")
            with store["lock"]:
                if len(store["data"]) > 800:
                    store["data"].clear()
                store["data"][key] = (time.time(), df)
            return df.copy()
        except Exception as exc:
            last = exc
            if attempt < retries - 1:
                time.sleep(2 + 2 * attempt)
    raise last or RuntimeError(f"No data for {symbol}")


def _candle(op, cl):
    if cl > op:
        return "UP", "Bullish", "Open < Close"
    if cl < op:
        return "DOWN", "Bearish", "Open > Close"
    return "FLAT", "Doji", "Open = Close"


def monthly_rows(daily):
    df = daily.sort_index().copy()
    df["_M"] = df.index.to_period("M")
    rows = []
    for period, g in df.groupby("_M", sort=True):
        op = safe_float(g.iloc[0]["Open"])
        hi, lo = safe_float(g["High"].max()), safe_float(g["Low"].min())
        cl = safe_float(g.iloc[-1]["Close"])
        d, c, rel = _candle(op, cl)
        rows.append({
            "Month": period.to_timestamp(), "Open": op, "High": hi,
            "Low": lo, "Close": cl, "Change": cl - op,
            "Change %": (cl - op) / op * 100 if op else 0.0,
            "Low → High": hi - lo,
            "Low → High %": (hi - lo) / lo * 100 if lo else 0.0,
            "Low Date": g["Low"].idxmin(), "High Date": g["High"].idxmax(),
            "Direction": d, "Candle": c, "Open vs Close": rel,
            "Volume": safe_float(g["Volume"].sum())})
    return rows


def daily_rows(daily):
    df = daily.sort_index()
    rows, prev = [], None
    for idx, r in df.iterrows():
        op, hi = safe_float(r["Open"]), safe_float(r["High"])
        lo, cl = safe_float(r["Low"]), safe_float(r["Close"])
        base = prev if prev is not None else op
        d, c, rel = _candle(op, cl)
        rows.append({
            "Date": idx, "Open": op, "High": hi, "Low": lo, "Close": cl,
            "Change": cl - base,
            "Change %": (cl - base) / base * 100 if base else 0.0,
            "Day Range": hi - lo,
            "Range %": (hi - lo) / lo * 100 if lo else 0.0,
            "Direction": d, "Candle": c, "Open vs Close": rel,
            "Volume": safe_float(r["Volume"])})
        prev = cl
    return rows


def fetch_rows(store, symbol, start, end, mode):
    daily = download_daily(store, symbol, start, end)
    return monthly_rows(daily) if mode == "MONTH" else daily_rows(daily)


def run_load(stocks, mode, start, end):
    store = price_store()
    rows, errors = [], []
    bar = st.progress(0.0, text="Starting ...")
    total = len(stocks)
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futs = {pool.submit(fetch_rows, store, s["symbol"], start, end, mode): s
                for s in stocks}
        for i, fut in enumerate(as_completed(futs), 1):
            s = futs[fut]
            try:
                for r in fut.result():
                    rows.append({"Stock": s["symbol"], "Company": s["name"],
                                 "Sector": s["sector"], "Index": s["index"],
                                 **r})
            except Exception as exc:
                errors.append({"symbol": s["symbol"], "error": str(exc)})
            bar.progress(i / total, text=f"{i}/{total}  {s['symbol']}")
    bar.empty()

    cols = MONTH_COLS if mode == "MONTH" else DAY_COLS
    df = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
    if not df.empty:
        key = "Month" if mode == "MONTH" else "Date"
        for c in DATE_COLS:
            if c in df.columns:
                df[c] = pd.to_datetime(df[c])
        df = df.sort_values(["Stock", key]).reset_index(drop=True)
    return {"df": df, "errors": errors, "mode": mode,
            "start": start, "end": end}


# ================================================================
# TRADINGVIEW
# ================================================================

def tv_symbol(sym):
    """NSE symbol -> TradingView symbol (M&M -> M_M, BAJAJ-AUTO -> BAJAJ_AUTO)."""
    return re.sub(r"[&\-]", "_", sym)


def tv_link(sym):
    """Daily-timeframe full chart page for the NSE stock."""
    return ("https://www.tradingview.com/chart/?symbol=NSE%3A"
            f"{quote(tv_symbol(sym))}&interval=D")


def tv_url(sym, label):
    """Link whose visible text (regex on the #fragment) is `label`."""
    return tv_link(sym) + "#" + str(label).replace("#", "")


def tv_embed(sym, height=620):
    html = """
    <div id="tv_chart" style="height:%dpx;width:100%%"></div>
    <script src="https://s3.tradingview.com/tv.js"></script>
    <script>
      new TradingView.widget({
        "autosize": true, "symbol": "NSE:%s", "interval": "D",
        "timezone": "Asia/Kolkata", "theme": "dark", "style": "1",
        "locale": "en", "enable_publishing": false,
        "allow_symbol_change": true, "hide_side_toolbar": false,
        "container_id": "tv_chart"
      });
    </script>""" % (height - 10, tv_symbol(sym))
    components.html(html, height=height)


def chart_panel(symbols, picked, key):
    """Chart section: row-click selection, or choose a stock from the list."""
    st.markdown("#### 📈 TradingView chart · Daily timeframe")
    if not symbols:
        return
    if picked in symbols:
        sym = picked
        st.caption(f"Chart for the row you clicked: **{sym}**  "
                   "(untick the row to choose another stock)")
    else:
        sym = st.selectbox("Chart stock (or click a row in the table above)",
                           symbols, key=f"chart_sel_{key}")
    c1, c2 = st.columns([1, 3])
    c1.link_button("↗ Open full chart in TradingView", tv_link(sym))
    show = c2.toggle("Show chart on this page", value=True,
                     key=f"chart_tog_{key}")
    if show:
        tv_embed(sym)
        st.caption("If the embedded chart says the symbol is unavailable, use "
                   "the 'Open full chart' button.")


def picked_symbol(event, symbols):
    try:
        rows = event.selection.rows
        if rows:
            return symbols[rows[0]]
    except Exception:
        pass
    return None


# ================================================================
# DISPLAY HELPERS
# ================================================================

def build_column_config(cols):
    C = st.column_config
    cfg = {}
    for c in ["Open", "High", "Low", "Close", "Change", "Low → High",
              "Day Range"]:
        cfg[c] = C.NumberColumn(c, format="%.2f")
    for c in ["Change %", "Low → High %", "Range %"]:
        cfg[c] = C.NumberColumn(c, format="%.2f")
    cfg["Volume"] = C.NumberColumn("Volume", format="%d")
    cfg["Month"] = C.DateColumn("Month", format="MMM-YY")
    cfg["Date"] = C.DateColumn("Date", format="DD/MM/YYYY")
    cfg["Low Date"] = C.DateColumn("Low Date", format="DD/MM/YYYY")
    cfg["High Date"] = C.DateColumn("High Date", format="DD/MM/YYYY")
    link_text = r"#(.*)$"
    try:
        cfg["Stock"] = C.LinkColumn("Stock ↗", display_text=link_text,
                                    pinned=True)
    except TypeError:
        cfg["Stock"] = C.LinkColumn("Stock ↗", display_text=link_text)
    cfg["Company"] = C.LinkColumn("Company ↗ (click = TradingView Daily)",
                                  display_text=link_text)
    return {k: v for k, v in cfg.items() if k in cols}


def styled(view, dirs):
    if view.size > 120_000:
        return view

    def trend(col):
        out = []
        for d in dirs.loc[col.index]:
            out.append(f"color:{GREEN};font-weight:600" if d == "UP"
                       else f"color:{RED};font-weight:600" if d == "DOWN"
                       else "")
        return out

    subset = [c for c in TREND_COLS if c in view.columns]
    sty = view.style
    if subset:
        sty = sty.apply(trend, subset=subset)
    return sty


def heat_css(v):
    a = min(abs(v) / 12.0, 1.0)
    t = (34, 197, 94) if v >= 0 else (239, 68, 68)
    b = (17, 27, 46)
    w = 0.18 + 0.62 * a
    r, g, bl = (int(x + (y - x) * w) for x, y in zip(b, t))
    return f"background-color:rgb({r},{g},{bl});color:#fff;font-weight:600"


def heat_table(df):
    sec = df.drop_duplicates("Stock").set_index("Stock")["Sector"]
    piv = (df.pivot_table(index="Stock", columns="Month", values="Change %",
                          aggfunc="first").sort_index())
    months = [c.strftime("%b-%y") for c in piv.columns]
    piv.columns = months
    out = piv.copy()
    out.insert(0, "Sector", sec.reindex(out.index))
    out["UP"] = (piv > 0).sum(axis=1).astype(float)
    out["DOWN"] = (piv < 0).sum(axis=1).astype(float)
    out["Avg %"] = piv.mean(axis=1)

    foot = pd.DataFrame(index=["AVERAGE %", "UP stocks", "DOWN stocks"],
                        columns=out.columns, dtype=float)
    foot["Sector"] = ""
    foot.loc["AVERAGE %", months] = piv.mean()
    foot.loc["UP stocks", months] = (piv > 0).sum()
    foot.loc["DOWN stocks", months] = (piv < 0).sum()
    full = pd.concat([out, foot])
    for c in months + ["UP", "DOWN", "Avg %"]:
        full[c] = full[c].astype(float)

    stk = list(out.index)
    ix = pd.IndexSlice
    sty = full.style.format(precision=2, na_rep="-")
    mapper = getattr(sty, "map", None) or sty.applymap
    sty = mapper(lambda v: heat_css(v) if pd.notna(v) else "",
                 subset=ix[stk, months])
    sty = (sty
           .format("{:+.1f}%", subset=ix[stk + ["AVERAGE %"], months], na_rep="-")
           .format("{:.0f}", subset=ix[["UP stocks", "DOWN stocks"], months],
                   na_rep="-")
           .format("{:.0f}", subset=ix[stk, ["UP", "DOWN"]])
           .format("{:+.2f}%", subset=ix[stk, ["Avg %"]]))
    return sty, len(full)


def make_excel(view, full_df, errors, month_mode):
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    num_cols = {"Open", "High", "Low", "Close", "Change", "Change %",
                "Low → High", "Low → High %", "Day Range", "Range %"}

    def style(ws):
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        fill = PatternFill("solid", fgColor="1F2A44")
        for c in ws[1]:
            c.font = Font(bold=True, color="FBBF24")
            c.fill = fill
            c.alignment = Alignment(horizontal="center")
        for col in ws.columns:
            h = str(col[0].value or "")
            letter = get_column_letter(col[0].column)
            if h in DATE_COLS:
                f = "mmm-yy" if h == "Month" else "dd/mm/yyyy"
                for c in col[1:]:
                    c.number_format = f
                ws.column_dimensions[letter].width = 12
                continue
            if h in num_cols:
                for c in col[1:]:
                    c.number_format = "#,##0.00"
            elif h == "Volume":
                for c in col[1:]:
                    c.number_format = "#,##0"
            ws.column_dimensions[letter].width = min(
                max(len(str(c.value or "")) for c in col) + 3, 40)

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        view.to_excel(w, sheet_name="NSE_Data", index=False)
        style(w.book["NSE_Data"])
        if month_mode and not full_df.empty:
            mat = (full_df.assign(MonthKey=full_df["Month"].dt.strftime("%Y-%m"))
                   .pivot_table(index=["Stock", "Sector"], columns="MonthKey",
                                values="Change %", aggfunc="first")
                   .reset_index())
            mat.columns.name = None
            mat.to_excel(w, sheet_name="Month_Matrix", index=False)
            ws = w.book["Month_Matrix"]
            style(ws)
            if ws.max_column >= 3 and ws.max_row >= 2:
                ws.conditional_formatting.add(
                    f"C2:{get_column_letter(ws.max_column)}{ws.max_row}",
                    ColorScaleRule(start_type="num", start_value=-10,
                                   start_color="F87171", mid_type="num",
                                   mid_value=0, mid_color="FFFFFF",
                                   end_type="num", end_value=10,
                                   end_color="4ADE80"))
        if errors:
            pd.DataFrame(errors).to_excel(w, sheet_name="Failed", index=False)
    return buf.getvalue()


# ================================================================
# SESSION STATE / CALLBACKS
# ================================================================

def set_picked(symbols):
    st.session_state["picked"] = sorted(symbols)


def add_picked(symbols):
    st.session_state["picked"] = sorted(set(st.session_state["picked"])
                                        | set(symbols))


def remove_picked(symbols):
    st.session_state["picked"] = sorted(set(st.session_state["picked"])
                                        - set(symbols))


def add_custom():
    raw = st.session_state.get("custom_text", "")
    syms = [s.strip().upper().replace(".NS", "")
            for s in re.split(r"[,\s]+", raw) if s.strip()]
    extra = st.session_state.setdefault("extra", [])
    for s in syms:
        if re.fullmatch(r"[A-Z0-9&\-]+", s) and s not in extra:
            extra.append(s)
    st.session_state["picked"] = sorted(set(st.session_state["picked"]) | set(syms))
    st.session_state["custom_text"] = ""


def set_preset(months, cur):
    st.session_state["from_ym"] = (cur // 100) * 100 + 1 if months == 0 \
        else ym_add(cur, -(months - 1))
    st.session_state["to_ym"] = cur


def layout_defaults(mode_key, all_cols):
    st.session_state.setdefault(f"show_{mode_key}", list(all_cols))
    st.session_state.setdefault(f"front_{mode_key}", [])
    st.session_state.setdefault(f"back_{mode_key}", [])


def layout_full(mode_key, all_cols):
    st.session_state[f"show_{mode_key}"] = list(all_cols)
    st.session_state[f"front_{mode_key}"] = []
    st.session_state[f"back_{mode_key}"] = []


def layout_compact(mode_key):
    st.session_state[f"show_{mode_key}"] = list(COMPACT[mode_key])
    st.session_state[f"front_{mode_key}"] = []
    st.session_state[f"back_{mode_key}"] = []


# ================================================================
# PAGE
# ================================================================

st.markdown("""
<style>
.block-container{padding-top:1.4rem;padding-bottom:1rem;max-width:100%}
[data-testid="stMetricValue"]{font-size:1.35rem}
div[data-testid="stSidebar"] .stButton>button{width:100%}
</style>
""", unsafe_allow_html=True)

uni = get_universe()
stocks_all = list(uni["stocks"])
info = uni["info"]

# user-added custom symbols
for sym in st.session_state.get("extra", []):
    if not any(s["symbol"] == sym for s in stocks_all):
        stocks_all.append({"symbol": sym, "name": f"{sym} (custom)",
                           "sector": "Other", "index": "Custom", "n50": False,
                           "bank": False, "fno": False, "lot": 0})
stocks_all.sort(key=lambda s: s["symbol"])
by_sym = {s["symbol"]: s for s in stocks_all}
all_syms = list(by_sym)

groups = {k: [s["symbol"] for s in stocks_all if s[k]]
          for k in ("n50", "bank", "fno")}
sectors = sorted({s["sector"] for s in stocks_all})

if "picked" not in st.session_state:
    st.session_state["picked"] = ["BHEL"] if "BHEL" in by_sym else []
st.session_state["picked"] = [s for s in st.session_state["picked"]
                              if s in by_sym]

# month list / defaults
today = date.today()
cur_ym = today.year * 100 + today.month
months = []
ym = 201001
while ym <= cur_ym:
    months.append(ym)
    ym = ym_add(ym, 1)
st.session_state.setdefault("from_ym", today.year * 100 + 1)
st.session_state.setdefault("to_ym", cur_ym)
st.session_state.setdefault("mode", "Month wise")

# ---------------------------------------------------------------- sidebar
with st.sidebar:
    st.markdown(f"### 📊 {APP_NAME}")
    st.caption(f"Developed By {DEVELOPER}")

    st.markdown("##### ⚡ Quick select (one click)")
    q1, q2 = st.columns(2)
    q1.button("⭐ NIFTY 50", on_click=set_picked, args=(groups["n50"],),
              type="primary" if set(st.session_state["picked"]) == set(groups["n50"]) else "secondary")
    q2.button("🏦 BANK NIFTY", on_click=set_picked, args=(groups["bank"],),
              type="primary" if set(st.session_state["picked"]) == set(groups["bank"]) else "secondary")
    q3, q4 = st.columns(2)
    q3.button("📑 NSE F&O", on_click=set_picked, args=(groups["fno"],),
              type="primary" if set(st.session_state["picked"]) == set(groups["fno"]) else "secondary")
    q4.button("✖ CLEAR ALL", on_click=set_picked, args=([],))

    st.markdown("##### 🔎 Browse / filter list")
    universe = st.selectbox("List", UNIVERSE_LABELS, key="universe")
    sector = st.selectbox("Sector", ["ALL SECTORS"] + sectors, key="sector")
    query = st.text_input("Search code / company", key="query",
                          placeholder="e.g. BHEL, TCS, bank").strip().lower()

    shown = []
    for s in stocks_all:
        if universe == "NIFTY 50" and not s["n50"]:
            continue
        if universe == "BANK NIFTY" and not s["bank"]:
            continue
        if universe.startswith("NSE F&O") and not s["fno"]:
            continue
        if sector != "ALL SECTORS" and s["sector"] != sector:
            continue
        if query and query not in f"{s['symbol']} {s['name']}".lower():
            continue
        shown.append(s["symbol"])
    st.caption(f"{len(shown):,} stocks match")
    a1, a2 = st.columns(2)
    a1.button("➕ Add shown", on_click=add_picked, args=(shown,))
    a2.button("➖ Remove shown", on_click=remove_picked, args=(shown,))

    st.multiselect(
        f"Selected stocks ({len(st.session_state['picked'])})",
        options=all_syms, key="picked",
        format_func=lambda s: f"{s} — {by_sym[s]['name'][:34]}")

    with st.expander("➕ Add any NSE symbol"):
        st.text_input("Symbols (comma separated)", key="custom_text",
                      placeholder="e.g. IRFC, SUZLON")
        st.button("Add symbols", on_click=add_custom)

    if st.button("⟳ Refresh NSE lists"):
        get_universe.clear()
        st.rerun()

    src = {"live": "live from NSE", "cache": "local cache",
           "built-in": "built-in fallback", "none": "unavailable"}
    st.caption(f"Stock list: {src.get(info['equity'], info['equity'])} · "
               f"F&O list: {src.get(info['fno'], info['fno'])}")
    if info["equity"] == "fallback":
        st.warning("NSE blocked the server, so only NIFTY 50 / BANK NIFTY / "
                   "F&O stocks are listed. Use 'Add any NSE symbol' for "
                   "others.")

# ------------------------------------------------------------------- main
st.title("📊 NSE Stock Month / Day Analysis")

with st.container(border=True):
    c1, c2, c3 = st.columns([1.2, 1, 1])
    mode_label = c1.radio("Analysis mode", ["Month wise", "Day wise"],
                          horizontal=True, key="mode")
    c2.selectbox("From (Month-YY)", months, format_func=ym_label,
                 key="from_ym")
    c3.selectbox("To (Month-YY)", months, format_func=ym_label, key="to_ym")

    p1, p2, p3, p4, p5, p6 = st.columns([1, 1, 1, 1, 0.3, 2])
    for col, (label, m) in zip((p1, p2, p3, p4),
                               (("YTD", 0), ("6M", 6), ("12M", 12), ("24M", 24))):
        col.button(label, on_click=set_preset, args=(m, cur_ym),
                   key=f"pre_{label}")
    load = p6.button("▶  LOAD DATA", type="primary", width="stretch")

from_ym, to_ym = st.session_state["from_ym"], st.session_state["to_ym"]
mode = "MONTH" if mode_label == "Month wise" else "DAY"
if from_ym <= to_ym:
    a, b = ym_to_dates(from_ym, to_ym)
    n = (to_ym // 100 - from_ym // 100) * 12 + (to_ym % 100 - from_ym % 100) + 1
    st.caption(f"Range: {a:%d/%m/%Y} → {b:%d/%m/%Y}  ({n} month"
               f"{'s' if n != 1 else ''})  ·  {len(st.session_state['picked'])} "
               f"stock(s) selected")
else:
    st.error("From month is after To month.")

if load:
    picked = st.session_state["picked"]
    if not picked:
        st.error("Please select at least one stock "
                 "(use the NIFTY 50 / F&O quick buttons).")
    elif from_ym > to_ym:
        st.error("From month cannot be after To month.")
    elif mode == "DAY" and len(picked) > 60 and \
            not st.session_state.get("big_ok"):
        st.warning(f"Day wise data for {len(picked)} stocks is very large. "
                   "Tick the confirmation box in the sidebar section below "
                   "or use Month wise.")
        st.session_state["need_big"] = True
    else:
        st.session_state.pop("need_big", None)
        st.session_state["res"] = run_load(
            [by_sym[s] for s in picked], mode, a, b)
        st.session_state.pop("excel_bytes", None)

if st.session_state.get("need_big"):
    st.checkbox("I understand - load Day wise data for many stocks",
                key="big_ok")

res = st.session_state.get("res")

tab_data, tab_heat, tab_fno, tab_help = st.tabs(
    ["📋 Data table", "🔥 Month heatmap", "📑 F&O / Options stocks", "❓ Help"])

# ---------------------------------------------------------------- data tab
with tab_data:
    if not res:
        st.info("Select stocks, choose a Month-YY range and press "
                "**LOAD DATA**.")
    elif res["df"].empty:
        st.error("No data returned. Failed stocks: " +
                 ", ".join(e["symbol"] for e in res["errors"][:20]))
    else:
        df = res["df"]
        rmode = res["mode"]
        month = rmode == "MONTH"
        all_cols = MONTH_COLS if month else DAY_COLS
        if rmode != mode:
            st.caption(f"Showing {'Month' if month else 'Day'} wise data "
                       "from the last load. Press LOAD DATA to refresh.")

        up = int((df["Direction"] == "UP").sum())
        dn = int((df["Direction"] == "DOWN").sum())
        fl = int((df["Direction"] == "FLAT").sum())
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Stocks", f"{df['Stock'].nunique():,}")
        m2.metric("Rows", f"{len(df):,}")
        m3.metric("▲ UP", f"{up:,}")
        m4.metric("▼ DOWN", f"{dn:,}", delta=None)
        m5.metric("Market signal",
                  "▲ BULLISH" if up > dn else "▼ BEARISH" if dn > up
                  else "● NEUTRAL", f"{up:,} vs {dn:,}")
        if res["errors"]:
            with st.expander(f"⚠ Failed stocks: {len(res['errors'])}"):
                st.dataframe(pd.DataFrame(res["errors"]), hide_index=True,
                             width="stretch")

        # ---- filters
        f1, f2, f3 = st.columns([2, 2, 2])
        txt = f1.text_input("Filter rows", placeholder="stock / company / sector",
                            key="row_filter").strip().lower()
        dirs_sel = f2.multiselect("Direction", ["UP", "DOWN", "FLAT"],
                                  default=["UP", "DOWN", "FLAT"],
                                  key="dir_filter")
        sec_sel = f3.multiselect("Sector", sorted(df["Sector"].unique()),
                                 key="sec_filter", placeholder="All sectors")

        view_df = df[df["Direction"].isin(dirs_sel)]
        if sec_sel:
            view_df = view_df[view_df["Sector"].isin(sec_sel)]
        if txt:
            mask = (view_df["Stock"] + " " + view_df["Company"] + " " +
                    view_df["Sector"]).str.lower().str.contains(txt, regex=False)
            view_df = view_df[mask]

        # ---- column manager (per user session)
        mk = rmode
        layout_defaults(mk, all_cols)
        with st.expander("⚙️ Customize columns  —  show / hide · move to FRONT · move to BACK"):
            lb1, lb2, lb3 = st.columns(3)
            lb1.button("Full layout", on_click=layout_full, args=(mk, all_cols),
                       key=f"bfull_{mk}")
            lb2.button("Compact layout", on_click=layout_compact, args=(mk,),
                       key=f"bcomp_{mk}")
            lb3.button("Reset", on_click=layout_full, args=(mk, all_cols),
                       key=f"breset_{mk}")

            up_file = st.file_uploader("Load saved layout (.json)",
                                       type=["json"], key="layout_upload")
            if up_file is not None:
                sig = f"{up_file.name}:{up_file.size}"
                if st.session_state.get("layout_sig") != sig:
                    try:
                        data = json.loads(up_file.getvalue().decode("utf-8"))
                        for k in ("MONTH", "DAY"):
                            cols_k = MONTH_COLS if k == "MONTH" else DAY_COLS
                            if k in data:
                                for part in ("show", "front", "back"):
                                    st.session_state[f"{part}_{k}"] = [
                                        c for c in data[k].get(part, [])
                                        if c in cols_k]
                        st.session_state["layout_sig"] = sig
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Invalid layout file: {exc}")

            st.multiselect("👁 Visible columns  (remove a column to HIDE it)",
                           all_cols, key=f"show_{mk}")
            st.multiselect("⏮ Move to FRONT  (in this order)", all_cols,
                           key=f"front_{mk}")
            st.multiselect("⏭ Move to BACK  (in this order)", all_cols,
                           key=f"back_{mk}")

            payload = {k: {part: st.session_state.get(f"{part}_{k}", [])
                           for part in ("show", "front", "back")}
                       for k in ("MONTH", "DAY")}
            st.download_button("💾 Save my layout", json.dumps(payload, indent=2),
                               file_name="my_column_layout.json",
                               mime="application/json")

        show = st.session_state[f"show_{mk}"]
        visible = [c for c in all_cols if c in show]
        front = [c for c in st.session_state[f"front_{mk}"] if c in visible]
        back = [c for c in st.session_state[f"back_{mk}"]
                if c in visible and c not in front]
        middle = [c for c in visible if c not in front and c not in back]
        order = front + middle + back
        if not order:
            st.warning("All columns are hidden - choose at least one visible "
                       "column.")
        else:
            view = view_df[order]
            view_disp = view.copy()
            row_syms = view_df["Stock"].tolist()
            if "Stock" in view_disp.columns:
                view_disp["Stock"] = [tv_url(x, x) for x in row_syms]
            if "Company" in view_disp.columns:
                view_disp["Company"] = [tv_url(x, n) for x, n in
                                        zip(row_syms, view_df["Company"])]
            st.caption(f"{len(view):,} rows · {len(order)} columns · "
                       "**click a Stock / Company name → TradingView Daily chart "
                       "opens** · tick a row to show the chart below · click a "
                       "header to sort")
            event = st.dataframe(
                styled(view_disp, view_df["Direction"]), hide_index=True,
                width="stretch",
                height=min(720, 40 + 35 * (len(view) + 1)),
                column_config=build_column_config(order),
                on_select="rerun", selection_mode="single-row",
                key=f"main_table_{rmode}")

            d1, d2, d3 = st.columns([1, 1, 3])
            d1.download_button("⇩ CSV (this layout)",
                               view.to_csv(index=False).encode("utf-8"),
                               file_name="nse_data.csv", mime="text/csv")
            if d2.button("Prepare Excel"):
                st.session_state["excel_bytes"] = make_excel(
                    view, df, res["errors"], month)
            if st.session_state.get("excel_bytes"):
                d3.download_button(
                    "⇩ Download Excel (this layout + month matrix)",
                    st.session_state["excel_bytes"],
                    file_name=f"NSE_{'Month' if month else 'Day'}_Wise_"
                              f"{datetime.now():%Y%m%d_%H%M}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument."
                         "spreadsheetml.sheet")

            chart_panel(sorted(set(row_syms)), picked_symbol(event, row_syms),
                        "data")

# ------------------------------------------------------------- heatmap tab
with tab_heat:
    if not res or res["df"].empty or res["mode"] != "MONTH":
        st.info("Load **Month wise** data to see the heatmap "
                "(Change % = Close vs Open of each month).")
    else:
        sty, nrows = heat_table(res["df"])
        st.caption("Green = up month · Red = down month · bottom rows show "
                   "the average and number of stocks up / down per month.")
        heat_event = st.dataframe(
            sty, width="stretch", height=min(760, 40 + 35 * (nrows + 1)),
            on_select="rerun", selection_mode="single-row", key="heat_table")
        heat_syms = sorted(res["df"]["Stock"].unique())
        pk = None
        try:
            r0 = heat_event.selection.rows
            if r0 and r0[0] < len(heat_syms):
                pk = heat_syms[r0[0]]
        except Exception:
            pass
        chart_panel(heat_syms, pk, "heat")

# ------------------------------------------------------------------ F&O tab
with tab_fno:
    fno_rows = [{"Stock": s["symbol"], "Company": s["name"],
                 "Sector": s["sector"], "Lot size": s["lot"] or None,
                 "NIFTY 50": s["n50"], "BANK NIFTY": s["bank"]}
                for s in stocks_all if s["fno"]]
    st.markdown(f"**NSE Futures & Options stocks: {len(fno_rows)}**  ·  source: "
                f"{src.get(info['fno'], info['fno'])}")
    fs = st.selectbox("Sector filter", ["ALL SECTORS"] + sectors,
                      key="fno_sector")
    fdf = pd.DataFrame(fno_rows)
    if fs != "ALL SECTORS" and not fdf.empty:
        fdf = fdf[fdf["Sector"] == fs]
    st.dataframe(fdf, hide_index=True, width="stretch", height=480)
    if not fdf.empty:
        st.button(f"➕ Add these {len(fdf)} stocks to my selection",
                  on_click=add_picked, args=(list(fdf["Stock"]),),
                  key="fno_add")

# ----------------------------------------------------------------- help tab
with tab_help:
    st.markdown(f"""
**Quick select** — NIFTY 50 / BANK NIFTY / NSE F&O replace your selection in one click
(the active button turns blue). Use **Add shown** to build a custom basket from filters.

**Month-YY range** — pick *From* and *To* (e.g. `Jan-26` → `Oct-26`). Data runs from the 1st day of
*From* to the last day of *To*. Presets: YTD, 6M, 12M, 24M.

**Month wise** — Open = first trading day's open · High / Low = month extremes ·
Close = last trading day's close · Change % = (Close − Open) / Open.

**Customize columns (per user)** — hide any column, push columns to the **front** or **back**,
use *Compact* / *Full* layouts, and **save your layout** as a file to re-load next visit.
The Excel / CSV download follows the same layout.

**TradingView** — click any *Stock* or *Company* name in the table to open that stock's
Daily chart on TradingView (new tab). Tick a row (or pick a stock) to show the chart on this page.

**Data sources** — NSE archives (stock lists) and Yahoo Finance via `yfinance` (prices).
Prices are cached for 30 minutes to reduce requests.

Developed By {DEVELOPER}
""")
