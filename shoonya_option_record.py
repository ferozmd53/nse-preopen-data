# CANDLE COLUMN ORDER: Call Change OI is beside Call OI; Put Change OI is beside Put OI.
import os
import sys
import time
import math
from datetime import datetime as dt
import warnings
from datetime import timezone, timedelta

warnings.filterwarnings("ignore")
try:
    import requests
    response = requests.get('https://api.ipify.org', timeout=5)
    print(f"\nYour current IP address is: {response.text}")
except:
    pass
# ---- dependency bootstrap ----
for pkg in ["pandas", "pyotp", "xlwings", "requests", "numpy",
            "selenium", "webdriver_manager", "openpyxl", "scipy"]:
    try:
        __import__(pkg)
    except ImportError:
        os.system(f"{sys.executable} -m pip install -U {pkg}")

import pandas as pd
import numpy as np
import pyotp
import xlwings as xw
import requests
from scipy.optimize import brentq

# Fix NumPy 2.0 compatibility
if not hasattr(np, 'PINF'):
    np.PINF = np.inf
if not hasattr(np, 'NINF'):
    np.NINF = -np.inf

from NorenRestApiPy.NorenApi import NorenApi

# ============================================================
# NSE IV ENGINE
# ============================================================
# NSE publicly states that its option-chain IV uses a 10% interest rate and
# is dynamic. NSE does not publish the complete production formula/time
# convention on the option-chain page. The engine below reproduces the
# observed NSE convention from the supplied 11-Aug-2026 snapshot:
#   * spot underlying (not the futures price)
#   * 10% continuously compounded risk-free rate
#   * Black-Scholes European pricing
#   * calendar-day DTE, using a minimum of 1 day on expiry day
#   * independent CE and PE IV solves
# This is intentionally independent of the third-party GetIVGreeks package.
# ============================================================

# ============================================================
# CONFIGURATION - Hardcoded symbol
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WORKBOOK_NAME = os.path.join(BASE_DIR, "shoonya_OptionChain.xlsx")
SYMBOL = "NIFTY"  # Hardcoded symbol
EXCHANGE = "NFO"
SPOT_EXCHANGE = "NSE"
NUMBER_OF_STRIKES = 2
STRIKE_STEP = 50
NIFTY_SPOT_TOKEN = "26000"  # Hardcoded NIFTY spot token
AGGREGATION_INTERVAL = 1  # Default candle interval in minutes
# ============================================================

# Timezone setup
IST = timezone(timedelta(hours=5, minutes=30))
MARKET_START = 9 * 60 + 15  # 9:15 AM in minutes
MARKET_END = 23 * 60 + 59   # 3:30 PM in minutes

live_data = {}
subs_lst = []
feed_opened = False
OptionChain_template = []
api = None
tick_counter = 0
# Tick history is no longer capped at 1000 rows. Every received tick is kept until Excel's row limit.
max_ticks_to_store = None

# Rebuild candles at most once per refresh cycle when new ticks are added.
last_candle_build_tick = 0

# Timestamp variables
feed_time = ""
request_time = ""
last_traded_time = ""

# Last market-data signature written to Tick_History.
# Time columns are intentionally excluded so a repeated snapshot is not written again.
last_tick_signature = None

# Store raw tick data for aggregation
raw_ticks = []
last_aggregation_time = None
last_candle_tick_counter = 0

# Global variables for NSE IV calculation
current_expiry_date = None
current_spot = 0.0
current_future = 0.0
current_iv_underlying = 0.0     # the price actually fed into Black-Scholes
current_iv_underlying_mode = "SPOT"  # "SPOT" or "FUTURE", read from OptionChain!B7
current_atm_strike = 0.0
current_atm_call_price = 0.0
current_atm_put_price = 0.0
current_iv_dte_days = 0
current_iv_T = 0.0
current_iv_timestamp = None
NSE_IV_RATE = 0.10
NSE_IV_DAYS_PER_YEAR = 365.0


def convert_to_float(item):
    try:
        if item is None:
            return 0.0
        if isinstance(item, (int, float)):
            return float(item)
        if isinstance(item, str):
            item = item.replace(',', '').strip()
            if item == '' or item == '-':
                return 0.0
            return float(item)
        return 0.0
    except (ValueError, TypeError):
        return 0.0


def nearest_strike(spot, step=50.0):
    """Calculate the nearest strike price using the efficient method"""
    a = (spot // step) * step
    b = a + step
    return int(b if spot - a > b - spot else a)


def epoch_to_ist(epoch_time):
    try:
        utc_dt = dt.fromtimestamp(int(epoch_time), timezone.utc)
        ist_dt = utc_dt.astimezone(IST)
        return ist_dt
    except Exception as e:
        print(f"⚠️ Error converting epoch {epoch_time}: {e}")
        return dt.now()


def format_timestamp(ts_value):
    """Return HH:MM:SS safely. Never return an Excel serial/date."""
    if ts_value is None or ts_value == "":
        return ""
    try:
        if isinstance(ts_value, (int, float, np.integer, np.floating)):
            x = float(ts_value)
            # Unix epoch seconds only
            if x > 100000000:
                return epoch_to_ist(x).strftime("%H:%M:%S")
            # Excel serial time/day: convert only if clearly a fraction of a day
            if 0 <= x < 1:
                total = int(round(x * 86400))
                h = (total // 3600) % 24
                m = (total % 3600) // 60
                s = total % 60
                return f"{h:02d}:{m:02d}:{s:02d}"
            return ""
        s = str(ts_value).strip()
        if not s or s.upper() in ("NIFTY 50", "NIFTY50"):
            return ""
        try:
            x = float(s)
            if x > 100000000:
                return epoch_to_ist(x).strftime("%H:%M:%S")
            if 0 <= x < 1:
                total = int(round(x * 86400))
                return f"{(total//3600)%24:02d}:{(total%3600)//60:02d}:{total%60:02d}"
        except Exception:
            pass
        # Handle HH:MM, HH:MM:SS, and minute:second formats.
        parts = s.split(":")
        if len(parts) == 3:
            h = int(float(parts[0])); m = int(float(parts[1])); sec = int(float(parts[2]))
            return f"{h%24:02d}:{m%60:02d}:{sec%60:02d}"
        if len(parts) == 2:
            a = int(float(parts[0])); b = int(float(parts[1]))
            if a < 24:
                return f"{a:02d}:{b%60:02d}:00"
            return f"{a//60:02d}:{a%60:02d}:{b%60:02d}"
    except Exception:
        pass
    return ""

def format_feed_datetime(ts_value):
    """Return a real text datetime, never an Excel 1900 date."""
    if ts_value is None or ts_value == "":
        return ""
    try:
        if isinstance(ts_value, (int, float, np.integer, np.floating)):
            x = float(ts_value)
            if x > 100000000:
                return epoch_to_ist(x).strftime("%Y-%m-%d %H:%M:%S")
            t = format_timestamp(x)
            return f"{dt.now(IST).strftime('%Y-%m-%d')} {t}" if t else ""
        s = str(ts_value).strip()
        if not s or s.upper() in ("NIFTY 50", "NIFTY50"):
            return ""
        try:
            x = float(s)
            if x > 100000000:
                return epoch_to_ist(x).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            pass
        t = format_timestamp(s)
        if t:
            return f"{dt.now(IST).strftime('%Y-%m-%d')} {t}"
        return s
    except Exception:
        return ""

def parse_date(date_input):
    if date_input is None:
        return None
    
    if hasattr(date_input, 'date'):
        return date_input.date() if hasattr(date_input, 'date') else date_input
    
    date_str = str(date_input).strip()
    
    formats = [
        '%d-%m-%Y',
        '%Y-%m-%d',
        '%d/%m/%Y',
        '%Y/%m/%d',
        '%d-%b-%Y',
        '%d %b %Y',
        '%d-%m-%y',
        '%Y-%m-%d %H:%M:%S',
    ]
    
    for fmt in formats:
        try:
            return dt.strptime(date_str, fmt).date()
        except (ValueError, TypeError):
            continue
    
    try:
        return pd.to_datetime(date_str).date()
    except:
        raise ValueError(f"Unable to parse date: {date_str}")


def get_field(token_key, field, default=0):
    """Get a field from live_data with proper handling"""
    try:
        data = live_data.get(token_key, {})
        value = data.get(field, default)
        if value is None:
            return default
        return value
    except Exception:
        return default


def _market_session_date(now_ist):
    """Return the trading-session date the CURRENT market data actually
    belongs to.

    Shoonya's LTP/quote feed freezes at the last traded price whenever the
    market is shut (weekends, before 9:15 AM), so a live option's price on a
    Sunday is really Friday's close - it still has Friday's time value, not
    Sunday's. Feeding wall-clock 'now' into the DTE calc on a non-trading
    day/hour silently shortens T and inflates every IV. This walks the
    valuation date back to the session the live prices actually reflect.
    """
    d = now_ist.date()
    minutes_now = now_ist.hour * 60 + now_ist.minute

    # Weekday before the market opens: still quoting yesterday's close.
    if now_ist.weekday() < 5 and minutes_now < MARKET_START:
        d -= timedelta(days=1)

    # Roll back over weekends (Sat=5, Sun=6) to the prior trading day.
    while d.weekday() >= 5:
        d -= timedelta(days=1)

    return d


def _nse_calendar_dte(expiry_date, valuation_dt=None):
    """Return NSE-style observed calendar DTE.

    NSE's public page documents the 10% rate and dynamic IV, but does not
    publish the complete production time convention. For the supplied
    expiry-day snapshot, using one calendar day on expiry day reproduces the
    NSE IV scale (~6.8% for the 24,450 CE) instead of the ~39% produced by
    using only the intraday hours remaining.
    """
    if expiry_date is None:
        return 0

    if valuation_dt is None:
        valuation_dt = dt.now(IST).replace(tzinfo=None)

    today = valuation_dt.date() if hasattr(valuation_dt, "date") else valuation_dt
    try:
        dte = (expiry_date - today).days
    except Exception:
        return 0

    # Expiry-day convention: keep a minimum of one calendar day.
    return max(1, int(dte))


def _normal_cdf(x):
    return 0.5 * (1.0 + math.erf(x / math.sqrt(2.0)))


def _bs_call_price(spot, strike, rate, T, sigma):
    if T <= 0 or sigma <= 0:
        return max(spot - strike, 0.0)
    sqrt_t = math.sqrt(T)
    d1 = (math.log(spot / strike) + (rate + 0.5 * sigma * sigma) * T) / (sigma * sqrt_t)
    d2 = d1 - sigma * sqrt_t
    return spot * _normal_cdf(d1) - strike * math.exp(-rate * T) * _normal_cdf(d2)


def _bs_put_price(spot, strike, rate, T, sigma):
    if T <= 0 or sigma <= 0:
        return max(strike - spot, 0.0)
    sqrt_t = math.sqrt(T)
    d1 = (math.log(spot / strike) + (rate + 0.5 * sigma * sigma) * T) / (sigma * sqrt_t)
    d2 = d1 - sigma * sqrt_t
    return strike * math.exp(-rate * T) * _normal_cdf(-d2) - spot * _normal_cdf(-d1)


def _solve_iv(option_price, pricing_function, lower_bound, upper_bound, intrinsic=0.0):
    """Solve IV via scipy's brentq (validated against NSE's published IV in
    IV_EXACT_MATCH.py - Brent's method root-find, same convention as NSE's
    option-chain page: 10% rate, calendar-day T, spot/future underlying).
    """
    if option_price <= 0:
        return 0.0

    # An option trading at/below intrinsic value has no time value to solve
    # a volatility from - mirrors IV_EXACT_MATCH.py's intrinsic check.
    if option_price <= intrinsic:
        return 0.0

    # Market price must be inside the model's no-arbitrage range.
    if option_price < lower_bound - 1e-7 or option_price > upper_bound + 1e-7:
        return 0.0

    objective = lambda sigma: pricing_function(sigma) - option_price
    try:
        iv = brentq(objective, 1e-4, 5.0)
    except ValueError:
        return 0.0

    return iv * 100.0


def init_iv_calculator(spot_ltp, future_ltp, atm_strike, atm_call_price, atm_put_price,
                        expiry_date, underlying_mode="SPOT"):
    """Refresh the NSE IV engine with the current market snapshot.

    underlying_mode selects which price is fed into Black-Scholes as S:
      "SPOT"   -> raw index LTP (the old, only, behaviour)
      "FUTURE" -> the NIFTY future LTP
    Controlled live from OptionChain!B7 so you can A/B test against the
    NSE page without editing code - see run_option_chain().
    """
    global current_expiry_date, current_spot, current_future
    global current_atm_strike, current_atm_call_price, current_atm_put_price
    global current_iv_dte_days, current_iv_T, current_iv_timestamp
    global current_iv_underlying, current_iv_underlying_mode

    current_spot = convert_to_float(spot_ltp)
    current_future = convert_to_float(future_ltp)
    current_atm_strike = convert_to_float(atm_strike)
    current_atm_call_price = convert_to_float(atm_call_price)
    current_atm_put_price = convert_to_float(atm_put_price)
    current_expiry_date = expiry_date
    current_iv_timestamp = dt.now(IST).replace(tzinfo=None)

    current_iv_underlying_mode = "FUTURE" if str(underlying_mode).strip().upper().startswith("F") else "SPOT"
    current_iv_underlying = current_future if current_iv_underlying_mode == "FUTURE" and current_future > 0 else current_spot

    # DTE must be dated to the trading session the live prices belong to,
    # not to wall-clock 'now' - see _market_session_date().
    session_date = _market_session_date(current_iv_timestamp)
    dte_valuation_dt = dt.combine(session_date, current_iv_timestamp.time())
    current_iv_dte_days = _nse_calendar_dte(expiry_date, dte_valuation_dt)
    current_iv_T = current_iv_dte_days / NSE_IV_DAYS_PER_YEAR

    if current_iv_underlying <= 0 or current_iv_dte_days <= 0:
        return False

    print(
        f"🔄 NSE IV engine: Underlying({current_iv_underlying_mode})={current_iv_underlying:.2f} "
        f"[Spot={current_spot:.2f} Fut={current_future:.2f}] "
        f"Expiry={expiry_date} DTE={current_iv_dte_days} "
        f"T={current_iv_T:.8f} r={NSE_IV_RATE:.2%}"
    )
    return True


def calculate_iv_for_strike(strike_price, call_price, put_price):
    """Calculate independent NSE-style CE IV and PE IV for one strike.

    Prices off current_iv_underlying (set by init_iv_calculator from the
    OptionChain!B7 SPOT/FUTURE toggle), not always raw spot.
    """
    strike = convert_to_float(strike_price)
    call = convert_to_float(call_price)
    put = convert_to_float(put_price)
    underlying = current_iv_underlying

    if (
        underlying <= 0 or strike <= 0 or current_iv_T <= 0
        or (call <= 0 and put <= 0)
    ):
        return 0.0, 0.0

    discount_factor = math.exp(-NSE_IV_RATE * current_iv_T)

    # European Black-Scholes no-arbitrage bounds using NSE's 10% rate.
    call_lower = max(underlying - strike * discount_factor, 0.0)
    call_upper = underlying
    put_lower = max(strike * discount_factor - underlying, 0.0)
    put_upper = strike * discount_factor

    # Simple (undiscounted) intrinsic value - same check IV_EXACT_MATCH.py
    # uses to reject option prices with no time value before solving.
    call_intrinsic = max(underlying - strike, 0.0)
    put_intrinsic = max(strike - underlying, 0.0)

    call_iv = 0.0
    put_iv = 0.0

    if call > 0:
        call_iv = _solve_iv(
            call,
            lambda sigma: _bs_call_price(
                underlying, strike, NSE_IV_RATE, current_iv_T, sigma
            ),
            call_lower,
            call_upper,
            call_intrinsic,
        )

    if put > 0:
        put_iv = _solve_iv(
            put,
            lambda sigma: _bs_put_price(
                underlying, strike, NSE_IV_RATE, current_iv_T, sigma
            ),
            put_lower,
            put_upper,
            put_intrinsic,
        )

    return round(call_iv, 2), round(put_iv, 2)


# ----------------------------------------------------------------------
# Workbook setup

# ----------------------------------------------------------------------
def get_or_create_workbook():
    global AGGREGATION_INTERVAL, tick_counter
    
    if os.path.exists(WORKBOOK_NAME):
        wb = xw.Book(WORKBOOK_NAME)
    else:
        wb = xw.Book()
        wb.save(WORKBOOK_NAME)

    sheet_names = [s.name.lower() for s in wb.sheets]
    
    if "login" not in sheet_names:
        login_sheet = wb.sheets.add("Login")
        login_sheet.range("A1").value = "Field"
        login_sheet.range("B1").value = "Value"
        labels = ["", "User ID (no _U)", "Password", "TOTP Secret", "Secret Code (OAuth)"]
        for i, label in enumerate(labels, start=1):
            login_sheet.range(f"A{i}").value = label
        login_sheet.range("C2").value = "Fill B2:B5 then run the script."
    else:
        login_sheet = wb.sheets["Login"]

    if "optionchain" not in sheet_names:
        oc_sheet = wb.sheets.add("OptionChain")
        oc_sheet.range("A1").value = "Expiry (dd-mm-yyyy) =>"
        oc_sheet.range("A2").value = "NoOfStrikes each side =>"
        oc_sheet.range("A3").value = "RefreshRate(sec) =>"
        oc_sheet.range("A4").value = "Aggregation Interval (min) =>"
        oc_sheet.range("B2").value = NUMBER_OF_STRIKES
        oc_sheet.range("B3").value = 1
        oc_sheet.range("B4").value = AGGREGATION_INTERVAL
        oc_sheet.range("A5").value = "Post-Market ON/OFF =>"
        oc_sheet.range("B5").value = "ON"
        oc_sheet.range("A6").value = "Run Until (HH:MM) =>"
        oc_sheet.range("B6").value = "18:00"
        oc_sheet.range("A7").value = "IV Underlying (SPOT/FUTURE) =>"
        oc_sheet.range("B7").value = "SPOT"
        oc_sheet.range("C1").value = "Available expiries -->"
        oc_sheet.range("G1").value = f"Symbol: {SYMBOL}"
        oc_sheet.range("H1").value = f"Token: {NIFTY_SPOT_TOKEN}"
        oc_sheet.range("I4").value = "Enter minutes (1, 2, 3, 5, 10, 15, 30, 60)"
    else:
        oc_sheet = wb.sheets["OptionChain"]
        oc_sheet.range("G1").value = f"Symbol: {SYMBOL}"
        oc_sheet.range("H1").value = f"Token: {NIFTY_SPOT_TOKEN}"
        oc_sheet.range("B2").value = NUMBER_OF_STRIKES
        if not oc_sheet.range("A5").value:
            oc_sheet.range("A5").value = "Post-Market ON/OFF =>"
        if not oc_sheet.range("B5").value:
            oc_sheet.range("B5").value = "ON"
        if not oc_sheet.range("A6").value:
            oc_sheet.range("A6").value = "Run Until (HH:MM) =>"
        if not oc_sheet.range("B6").value:
            oc_sheet.range("B6").value = "18:00"
        if not oc_sheet.range("A7").value:
            oc_sheet.range("A7").value = "IV Underlying (SPOT/FUTURE) =>"
        if not oc_sheet.range("B7").value:
            oc_sheet.range("B7").value = "SPOT"
        # Read aggregation interval from Excel
        agg_interval = oc_sheet.range("B4").value
        if agg_interval and isinstance(agg_interval, (int, float)) and agg_interval > 0:
            AGGREGATION_INTERVAL = int(agg_interval)
            print(f"✅ Aggregation interval set to {AGGREGATION_INTERVAL} minutes")

    # CANONICAL Tick_History layout: exactly 31 columns A:AE.
    # Keep this list defined even when the sheet already exists.
    headers = [
        "Feed Time", "Request Time", "Last Traded Time",
        "Spot", "ATM Strike", "OTM Call Strike",
        "Call LTP", "Call Volume", "Call IV", "Call OI", "Call Change OI",
        "Call Total Buy", "Call Total Sell", "Call Buy-Sell Diff",
        "Call Bid", "Call Ask", "Call Bid-Ask Diff",
        "OTM Put Strike", "Put LTP", "Put Volume", "Put IV", "Put OI", "Put Change OI",
        "Put Total Buy", "Put Total Sell", "Put Buy-Sell Diff",
        "Put Bid", "Put Ask", "Put Bid-Ask Diff",
        "PCR OI", "PCR Change OI"
    ]

    # Canonical Tick_History headers MUST always be defined before any refresh.
    # This fixes: "cannot access local variable headers" on existing workbooks.
    headers = [
        "Feed Time", "Request Time", "Last Traded Time",
        "Spot", "ATM Strike", "OTM Call Strike",
        "Call LTP", "Call Volume", "Call IV", "Call OI", "Call Change OI",
        "Call Total Buy", "Call Total Sell", "Call Buy-Sell Diff",
        "Call Bid", "Call Ask", "Call Bid-Ask Diff",
        "OTM Put Strike", "Put LTP", "Put Volume", "Put IV", "Put OI", "Put Change OI",
        "Put Total Buy", "Put Total Sell", "Put Buy-Sell Diff",
        "Put Bid", "Put Ask", "Put Bid-Ask Diff", "PCR OI", "PCR Change OI"
    ]

    history_sheet_name = "Tick_History"
    if history_sheet_name.lower() not in sheet_names:
        history_sheet = wb.sheets.add(history_sheet_name)
        history_sheet.range("A1:AE1").value = headers
        history_sheet.range("A:C").number_format = "@"
        tick_counter = 0
    else:
        history_sheet = wb.sheets[history_sheet_name]
        # Check if headers exist, if not create them
        try:
            header_check = history_sheet.range("A1").value
            if header_check is None or header_check == "" or header_check != "Feed Time":
                print("📝 Creating missing headers in Tick_History...")
                history_sheet.range("A1:AE1").value = headers
                history_sheet.range("A:C").number_format = "@"
        except Exception as e:
            print(f"⚠️ Error checking headers: {e}")

    # Always keep the Tick_History header structure current.
    # This does not rewrite or delete existing tick data.
    try:
        history_sheet.range("A1:AE1").value = headers
        history_sheet.range("A:C").number_format = "@"
    except Exception as e:
        print(f"⚠️ Could not refresh Tick_History headers: {e}")

    # Count existing rows in Tick_History
    try:
        history_last_row = get_last_actual_row(history_sheet, "A", 1)
        tick_counter = max(0, history_last_row - 1) if history_last_row >= 2 else 0
        print(f"📚 Existing Tick_History rows: {tick_counter}")
    except Exception:
        tick_counter = 0

    # New sheet for aggregated candles - KEEP EXISTING DATA
    # Reuse the existing candle sheet. Prefer the user's existing "Candel" sheet;
    # otherwise reuse/create "Candles". Never create a second candle sheet on restart.
    if "candel" in sheet_names:
        candle_sheet_name = next(s.name for s in wb.sheets if s.name.lower() == "candel")
    elif "candles" in sheet_names:
        candle_sheet_name = next(s.name for s in wb.sheets if s.name.lower() == "candles")
    else:
        candle_sheet_name = "Candel"
    if candle_sheet_name.lower() not in sheet_names:
        candle_sheet = wb.sheets.add(candle_sheet_name)
        # Create headers for new candle sheet with IV
        candle_headers = [
            "Time", "Close",
            "Call Strike", "Call LTP", "Call Volume", "Call IV", "Call OI", "Call Change OI", "Call Bid-Ask Avg",
            "Call Total Buy", "Call Total Sell", "Call Buy-Sell Diff",
            "Put Strike", "Put LTP", "Put Volume", "Put IV", "Put OI", "Put Change OI", "Put Bid-Ask Avg",
            "Put Total Buy", "Put Total Sell", "Put Buy-Sell Diff",
            "PCR OI", "PCR Change OI",
            "Ticks"
        ]
        candle_sheet.range("A1").value = candle_headers
    else:
        candle_sheet = wb.sheets[candle_sheet_name]
        # Check if headers exist
        try:
            header_check = candle_sheet.range("A1").value
            if header_check is None or header_check == "":
                candle_headers = [
                    "Time", "Close",
                    "Call Strike", "Call IV", "Call OI", "Call Bid-Ask Avg", 
                    "Call Total Buy", "Call Total Sell", "Call Buy-Sell Diff",
                    "Put Strike", "Put IV", "Put OI", "Put Bid-Ask Avg",
                    "Put Total Buy", "Put Total Sell", "Put Buy-Sell Diff",
                    "Call Change OI", "Put Change OI",
                    "PCR OI", "PCR Change OI",
                    "Ticks"
                ]
                candle_sheet.range("A1:Y1").value = [candle_headers]
        except Exception:
            pass

    try:
        if "sheet1" in sheet_names and len(wb.sheets) > 2:
            wb.sheets["Sheet1"].delete()
    except Exception:
        pass

    # IMPORTANT: physically migrate old Candles data once.
    # This fixes existing rows where Call Change OI was stored in Q.
    fix_candle_column_order(candle_sheet)

    print(
        f"📍 Real Tick_History last row: {get_last_actual_row(history_sheet, 'A', 1)} | "
        f"Real Candles last row: {get_last_actual_row(candle_sheet, 'A', 1)}"
    )

    wb.save()
    return wb, login_sheet, oc_sheet, history_sheet, candle_sheet


# ----------------------------------------------------------------------
# Function to check Tick_History status
# ----------------------------------------------------------------------
def check_tick_history_status(history_sheet):
    """Debug function to check Tick_History sheet status"""
    try:
        print("🔍 Checking Tick_History status...")
        
        # Check if sheet exists
        last_row = get_last_actual_row(history_sheet, "A", 1)
        print(f"   Used range rows: {last_row}")
        
        if last_row >= 1:
            # Check headers
            headers = history_sheet.range("A1:Z1").value
            if headers and any(headers):
                print(f"   Headers found: {headers[0] if headers else 'None'}")
            else:
                print("   No headers found!")
        
        # Check if there's data in row 2
        if last_row >= 2:
            row2_data = history_sheet.range("A2:Z2").value
            if row2_data and any(row2_data):
                print(f"   Data found in row 2: {row2_data[0] if row2_data else 'None'}")
            else:
                print("   No data found in row 2!")
        
        print("✅ Tick_History status check complete")
        return True
    except Exception as e:
        print(f"⚠️ Error checking Tick_History: {e}")
        return False


# ----------------------------------------------------------------------
# Function to aggregate tick data into candles - UPDATED WITH IV
# ----------------------------------------------------------------------
def _parse_tick_datetime(value):
    """Convert Tick_History Request Time into a real pandas datetime."""
    if value is None or value == "":
        return pd.NaT

    # xlwings may return a real Python datetime.
    if isinstance(value, (dt, pd.Timestamp)):
        try:
            return pd.Timestamp(value)
        except Exception:
            return pd.NaT

    # Excel may return a numeric Excel serial.
    if isinstance(value, (int, float, np.integer, np.floating)):
        try:
            # Values around 1e9 are Unix timestamps; smaller values are Excel serials.
            if float(value) > 100000000:
                return pd.to_datetime(float(value), unit="s", errors="coerce")
            return pd.to_datetime(float(value), unit="D", origin="1899-12-30", errors="coerce")
        except Exception:
            return pd.NaT

    s = str(value).strip()
    if not s or s.lower() == "none":
        return pd.NaT

    # First try normal date/time parsing.
    ts = pd.to_datetime(s, errors="coerce", dayfirst=False)
    if not pd.isna(ts):
        return ts

    # Old Tick_History versions may contain only HH:MM:SS.
    try:
        parsed_time = dt.strptime(s.split(".")[0], "%H:%M:%S").time()
        today = dt.now().date()
        return pd.Timestamp(dt.combine(today, parsed_time))
    except Exception:
        return pd.NaT


def _market_anchored_bin(ts, interval_minutes):
    """
    Return the candle start time anchored at 09:15.
    This is important for 60-minute candles: they become 09:15, 10:15,
    11:15... instead of pandas' normal 09:00, 10:00... bins.
    """
    market_start = ts.normalize() + pd.Timedelta(hours=9, minutes=15)
    elapsed_seconds = (ts - market_start).total_seconds()
    bucket = int(elapsed_seconds // (interval_minutes * 60))
    return market_start + pd.Timedelta(minutes=bucket * interval_minutes)


def aggregate_candles(history_sheet, candle_sheet, interval_minutes):
    """
    REAL-TIME candle updater.

    IMPORTANT:
    - Does NOT recalculate/write old candle rows.
    - Only calculates the CURRENT candle from Tick_History.
    - If the current minute is already the last candle row:
        -> compare all 17 candle values
        -> write that ONE row only if something changed.
    - If a new minute starts:
        -> append exactly ONE new row.
    - Therefore rows 2..N are NEVER rewritten every refresh.
    - No whole-sheet clear, no whole-history rewrite, no repeated appends.
    """
    try:
        interval_minutes = max(1, int(interval_minutes))

        # ------------------------------------------------------------
        # READ TICK HISTORY
        # We still read history to calculate the current candle,
        # but we NEVER write historical candle rows again.
        # ------------------------------------------------------------
        last_row = get_last_actual_row(history_sheet, "A", 1)

        if last_row < 2:
            return 0

        # Only the ticks belonging to the CURRENT candle are needed here.
        # Reading the entire Tick_History sheet on every refresh (the old
        # behaviour) means this COM read + DataFrame build gets slower all
        # day as history grows - it's the main reason things crawl by the
        # afternoon. Bound the read to a recent window instead: generously
        # more rows than any single interval could accumulate even at the
        # fastest allowed refresh rate (0.05s -> ~1200 ticks/min).
        LOOKBACK_ROWS = max(500, interval_minutes * 1500)
        start_row = max(2, last_row - LOOKBACK_ROWS + 1)

        headers = history_sheet.range("A1:AE1").value
        data_rows = history_sheet.range(f"A{start_row}:AE{last_row}").value
        if not data_rows:
            return 0
        if data_rows and not isinstance(data_rows[0], list):
            data_rows = [data_rows]  # xlwings flattens a single-row read
        df = pd.DataFrame(data_rows, columns=headers)

        required = [
            "Request Time", "Spot", "OTM Call Strike", "Call LTP", "Call Volume", "Call IV", "Call OI", "Call Change OI",
            "Call Total Buy", "Call Total Sell", "Call Bid", "Call Ask",
            "OTM Put Strike", "Put LTP", "Put Volume", "Put IV", "Put OI", "Put Change OI", "Put Total Buy",
            "Put Total Sell", "Put Bid", "Put Ask", "PCR OI", "PCR Change OI"
        ]

        missing = [c for c in required if c not in df.columns]
        if missing:
            print(f"⚠️ Tick_History missing columns: {missing}")
            return 0

        # ------------------------------------------------------------
        # PARSE / CLEAN
        # ------------------------------------------------------------
        df["datetime"] = df["Request Time"].apply(_parse_tick_datetime)

        numeric_cols = [
            "Spot", "OTM Call Strike", "Call LTP", "Call Volume", "Call IV", "Call OI", "Call Change OI",
            "Call Total Buy", "Call Total Sell", "Call Bid", "Call Ask",
            "OTM Put Strike", "Put LTP", "Put Volume", "Put IV", "Put OI", "Put Change OI", "Put Total Buy",
            "Put Total Sell", "Put Bid", "Put Ask", "PCR OI", "PCR Change OI"
        ]

        for col in numeric_cols:
            df[col] = pd.to_numeric(
                df[col],
                errors="coerce"
            ).fillna(0.0)

        df = df.dropna(subset=["datetime"])
        df = df[df["Spot"] > 0].copy()

        if df.empty:
            return 0

        df = df.sort_values("datetime").reset_index(drop=True)

        # ------------------------------------------------------------
        # RUN WINDOW
        # ------------------------------------------------------------
        # Tick_History has already been filtered by store_tick_data.
        # Do not apply a second 15:30 cutoff here; otherwise the optional
        # post-market candle window would always remain blank.

        # ------------------------------------------------------------
        # IMPORTANT:
        # Find ONLY the latest/current candle.
        # Do not build every historical candle.
        # ------------------------------------------------------------
        latest_tick_time = df["datetime"].iloc[-1]

        current_bin = _market_anchored_bin(
            latest_tick_time,
            interval_minutes
        )

        current_group = df[
            df["datetime"].apply(
                lambda x: _market_anchored_bin(
                    x,
                    interval_minutes
                ) == current_bin
            )
        ].copy()

        if current_group.empty:
            return 0

        current_group = current_group.sort_values("datetime")
        last = current_group.iloc[-1]

        # ------------------------------------------------------------
        # BUILD ONLY CURRENT CANDLE
        # ------------------------------------------------------------
        call_bid_ask = (
            current_group["Call Ask"]
            - current_group["Call Bid"]
        ).mean()

        put_bid_ask = (
            current_group["Put Ask"]
            - current_group["Put Bid"]
        ).mean()

        call_buy = current_group["Call Total Buy"].sum()
        call_sell = current_group["Call Total Sell"].sum()

        put_buy = current_group["Put Total Buy"].sum()
        put_sell = current_group["Put Total Sell"].sum()

        # Use 24-hour format. This avoids 02:xx / 14:xx confusion.
        candle_time = current_bin.strftime("%Y-%m-%d %H:%M")

        new_values = [
            candle_time,
            float(last["Spot"]),
            float(last["OTM Call Strike"]),
            float(last["Call LTP"]),
            float(last["Call Volume"]),
            float(last["Call IV"]),
            float(current_group["Call OI"].mean()),
            float(last["Call Change OI"]),
            float(call_bid_ask),
            float(call_buy),
            float(call_sell),
            float(call_buy - call_sell),
            float(last["OTM Put Strike"]),
            float(last["Put LTP"]),
            float(last["Put Volume"]),
            float(last["Put IV"]),
            float(current_group["Put OI"].mean()),
            float(last["Put Change OI"]),
            float(put_bid_ask),
            float(put_buy),
            float(put_sell),
            float(put_buy - put_sell),
            float(last["PCR OI"]),
            float(last["PCR Change OI"]),
            int(len(current_group)),
        ]

        # ------------------------------------------------------------
        # HEADERS ONLY IF MISSING
        # NEVER REWRITE THEM EVERY REFRESH.
        # Header row moved to row 1 - the "N-Minute Candle Data" title is
        # dropped since that interval is already shown on OptionChain!B4.
        # ------------------------------------------------------------
        candle_headers = [
"Time", "Close",
            "Call Strike", "Call LTP", "Call Volume", "Call IV", "Call OI", "Call Change OI", "Call Bid-Ask Avg",
            "Call Total Buy", "Call Total Sell", "Call Buy-Sell Diff",
            "Put Strike", "Put LTP", "Put Volume", "Put IV", "Put OI", "Put Change OI", "Put Bid-Ask Avg",
            "Put Total Buy", "Put Total Sell", "Put Buy-Sell Diff",
            "PCR OI", "PCR Change OI",
            "Ticks"
        ]

        try:
            current_headers = candle_sheet.range("A1:Y1").value
            if current_headers and isinstance(current_headers, list):
                if len(current_headers) == 1 and isinstance(
                    current_headers[0], list
                ):
                    current_headers = current_headers[0]
        except Exception:
            current_headers = None

        if current_headers != candle_headers:
            candle_sheet.range("A1:Y1").value = [candle_headers]

        # ------------------------------------------------------------
        # FIND ONLY THE LAST CANDLE ROW
        #
        # We intentionally DO NOT read rows 2:N.
        # Only the latest row matters.
        # ------------------------------------------------------------
        try:
            candle_last_row = get_last_actual_row(candle_sheet, "A", 1)
        except Exception:
            candle_last_row = 1

        if candle_last_row < 2:
            # No candle exists yet -> add the first one.
            target_row = 2

            candle_sheet.range(
                f"A{target_row}:Y{target_row}"
            ).value = [new_values]

            candle_sheet.range(f"A{target_row}").number_format = "@"
            candle_sheet.range(
                f"B{target_row}:Y{target_row}"
            ).number_format = "#,##0.00"

            print(
                f"🆕 Candle created: {candle_time}"
            )

            return 1

        # ------------------------------------------------------------
        # READ ONLY THE LAST ROW
        # ------------------------------------------------------------
        last_row_values = candle_sheet.range(
            f"A{candle_last_row}:Y{candle_last_row}"
        ).value

        if last_row_values and isinstance(last_row_values, list):
            if (
                len(last_row_values) == 1
                and isinstance(last_row_values[0], list)
            ):
                last_row_values = last_row_values[0]
        else:
            last_row_values = []

        # ------------------------------------------------------------
        # NORMALIZE TIME
        # Excel can return datetime, string, or serial.
        # Always convert it to YYYY-MM-DD HH:MM.
        # ------------------------------------------------------------
        def normalize_candle_time(value):
            if value is None or value == "":
                return ""

            if isinstance(value, (dt, pd.Timestamp)):
                return pd.Timestamp(value).strftime(
                    "%Y-%m-%d %H:%M"
                )

            if isinstance(value, (int, float, np.integer, np.floating)):
                try:
                    # Excel datetime serial
                    parsed = pd.to_datetime(
                        float(value),
                        unit="D",
                        origin="1899-12-30",
                        errors="coerce"
                    )

                    if not pd.isna(parsed):
                        return parsed.strftime(
                            "%Y-%m-%d %H:%M"
                        )
                except Exception:
                    pass

            s = str(value).strip()

            # Try common datetime formats.
            for fmt in (
                "%Y-%m-%d %H:%M",
                "%Y-%m-%d %H:%M:%S",
                "%d/%m/%Y %H:%M",
                "%d/%m/%Y %H:%M:%S",
                "%m/%d/%Y %H:%M",
                "%m/%d/%Y %H:%M:%S",
                "%Y-%m-%d %I:%M",
                "%d/%m/%Y %I:%M",
                "%m/%d/%Y %I:%M",
            ):
                try:
                    return dt.strptime(
                        s,
                        fmt
                    ).strftime(
                        "%Y-%m-%d %H:%M"
                    )
                except Exception:
                    pass

            # Final pandas parser.
            try:
                parsed = pd.to_datetime(
                    s,
                    errors="coerce",
                    dayfirst=False
                )

                if not pd.isna(parsed):
                    return pd.Timestamp(parsed).strftime(
                        "%Y-%m-%d %H:%M"
                    )
            except Exception:
                pass

            return s

        existing_time = normalize_candle_time(
            last_row_values[0] if last_row_values else None
        )

        # ------------------------------------------------------------
        # CASE 1:
        # SAME MINUTE -> compare ONLY this row.
        # ------------------------------------------------------------
        if existing_time == candle_time:

            def values_equal(old, new):
                if old is None and new is None:
                    return True

                if old is None or new is None:
                    return False

                try:
                    old_f = float(old)
                    new_f = float(new)

                    if (
                        math.isfinite(old_f)
                        and math.isfinite(new_f)
                    ):
                        return abs(old_f - new_f) <= 1e-9

                except Exception:
                    pass

                return str(old).strip() == str(new).strip()

            changed = False

            for old, new in zip(
                last_row_values,
                new_values
            ):
                if not values_equal(old, new):
                    changed = True
                    break

            if not changed:
                # NOTHING changed.
                # ZERO Excel writes.
                return 0

            # One row ONLY.
            candle_sheet.range(
                f"A{candle_last_row}:Y{candle_last_row}"
            ).value = [new_values]

            candle_sheet.range(
                f"B{candle_last_row}:S{candle_last_row}"
            ).number_format = "#,##0.00"

            print(
                f"🔄 Updated ONLY current candle: "
                f"{candle_time}"
            )

            return 1

        # ------------------------------------------------------------
        # CASE 2:
        # NEW MINUTE -> append exactly ONE row.
        # ------------------------------------------------------------
        target_row = candle_last_row + 1

        candle_sheet.range(
            f"A{target_row}:Y{target_row}"
        ).value = [new_values]

        candle_sheet.range(
            f"B{target_row}:S{target_row}"
        ).number_format = "#,##0.00"

        print(
            f"🆕 New candle added ONLY: "
            f"{candle_time}"
        )

        return 1

    except Exception as e:
        print(
            f"⚠️ Error updating current candle: {e}"
        )
        import traceback
        traceback.print_exc()
        return 0




# ----------------------------------------------------------------------
# FIND LAST REAL DATA ROW
# ----------------------------------------------------------------------
def get_last_actual_row(sheet, column="A", header_row=1):
    """
    Returns the last row containing REAL content in the specified column.

    IMPORTANT:
    Do NOT use sheet.get_last_actual_row(...) for data positioning.
    Excel's used range can include old formatting, colors, borders, etc.
    That can create huge blank gaps such as writing Candles at row 274
    or Tick_History at row 1344.

    Column A is used because:
      Tick_History -> Feed Time is column A
      Candles      -> Time is column A
    """
    try:
        max_row = sheet.cells.last_cell.row
        last = sheet.range(f"{column}{max_row}").end("up").row

        # Never return a row above the header.
        return max(int(last), int(header_row))
    except Exception:
        return int(header_row)


# ----------------------------------------------------------------------
# FIX / MIGRATE EXISTING CANDLES SHEET
# ----------------------------------------------------------------------
def fix_candle_column_order(candle_sheet):
    """Ensure one canonical 25-column Candel/Candles layout.

    Old candle files had 21 columns. New files include Call/Put LTP and
    Volume, so we migrate old rows without creating another sheet.
    """
    headers = [
        "Time", "Close",
        "Call Strike", "Call LTP", "Call Volume", "Call IV", "Call OI", "Call Change OI",
        "Call Bid-Ask Avg", "Call Total Buy", "Call Total Sell", "Call Buy-Sell Diff",
        "Put Strike", "Put LTP", "Put Volume", "Put IV", "Put OI", "Put Change OI",
        "Put Bid-Ask Avg", "Put Total Buy", "Put Total Sell", "Put Buy-Sell Diff",
        "PCR OI", "PCR Change OI", "Ticks"
    ]
    try:
        last_row = get_last_actual_row(candle_sheet, "A", 1)
        if last_row < 1:
            candle_sheet.range("A1:Y1").value = [headers]
            return

        raw = candle_sheet.range(f"A1:Y{last_row}").value
        if raw and raw and not isinstance(raw[0], list):
            raw = [raw]
        first = list(raw[0]) if raw else []
        first_norm = [str(x).strip() if x is not None else "" for x in first]

        if first_norm[:25] == headers:
            print("✅ Candles column order already correct.")
            return

        # Old 21-column layout from the previous working versions.
        old_headers = [
            "Time", "Close", "Call Strike", "Call IV", "Call OI", "Call Bid-Ask Avg",
            "Call Total Buy", "Call Total Sell", "Call Buy-Sell Diff", "Put Strike", "Put IV",
            "Put OI", "Put Bid-Ask Avg", "Put Total Buy", "Put Total Sell", "Put Buy-Sell Diff",
            "Call Change OI", "Put Change OI", "PCR OI", "PCR Change OI", "Ticks"
        ]

        data = raw[1:] if first_norm[:21] == old_headers else []
        if data:
            migrated = []
            for row in data:
                r = list(row) + [None] * max(0, 21-len(row))
                # We cannot recover historical LTP/Volume if they were never stored.
                migrated.append([
                    r[0], r[1], r[2], 0.0, 0.0, r[3], r[4], r[16], r[5],
                    r[6], r[7], r[8], r[9], 0.0, 0.0, r[10], r[11], r[17],
                    r[12], r[13], r[14], r[15], r[18], r[19], r[20]
                ])
            candle_sheet.range(f"A1:Y{last_row}").clear_contents()
            candle_sheet.range("A1:Y1").value = [headers]
            if migrated:
                candle_sheet.range(f"A2:Y{1+len(migrated)}").value = migrated
            print(f"🔧 Migrated {len(migrated)} old candle rows to canonical LTP/Volume layout.")
        else:
            # Unknown/malformed header: fix only the header; do not create a sheet.
            candle_sheet.range("A1:Y1").value = [headers]
            print("🔧 Repaired Candel headers to canonical 25-column layout.")

        candle_sheet.range("A:A").number_format = "@"
        candle_sheet.range("B:Y").number_format = "#,##0.00"
    except Exception as e:
        print(f"⚠️ Candles column migration error: {e}")


# ----------------------------------------------------------------------
# Function to get all login data from Excel
# ----------------------------------------------------------------------
def get_all_login_data(login_sheet):
    data = login_sheet.range("A1:B30").value
    
    login_data = {
        'client_id': '',
        'user_id': '',
        'password': '',
        'totp_secret': '',
        'secret_code': '',
        'auth_code': '',
        'token': '',
        'usertoken': '',
        'ip_address': '',
        'token_timestamp': None
    }
    
    for row in data:
        if row and len(row) >= 2:
            field = str(row[0]).strip().upper() if row[0] else ""
            value = str(row[1]).strip() if row[1] else ""
            
            if "CLIENT_ID" in field:
                login_data['client_id'] = value
            elif "USER_ID" in field:
                login_data['user_id'] = value
            elif "PASSWORD" in field:
                login_data['password'] = value
            elif "TOTP_SECRET" in field:
                login_data['totp_secret'] = value
            elif "SECRET_CODE" in field:
                login_data['secret_code'] = value
            elif "AUTH CODE" in field:
                login_data['auth_code'] = value
            elif "TOKEN" in field and "USERTOKEN" not in field and "TIMESTAMP" not in field:
                login_data['token'] = value
            elif "USERTOKEN" in field:
                login_data['usertoken'] = value
            elif "IP_ADRESS" in field or "IP_ADDRESS" in field:
                login_data['ip_address'] = value
            elif "TOKEN_TIMESTAMP" in field or "TOKEN TIME" in field:
                try:
                    login_data['token_timestamp'] = dt.strptime(value, '%Y-%m-%d %H:%M:%S')
                except:
                    pass
    
    if not login_data['client_id'] and login_data['user_id']:
        login_data['client_id'] = f"{login_data['user_id']}_U"
    
    return login_data


def update_login_data(login_sheet, auth_code, token, usertoken):
    data = login_sheet.range("A1:B30").value
    
    for i, row in enumerate(data):
        if row and len(row) >= 2:
            field = str(row[0]).strip().upper() if row[0] else ""
            
            if "AUTH CODE" in field:
                login_sheet.range(f"B{i+1}").value = auth_code
            elif "TOKEN" in field and "USERTOKEN" not in field and "TIMESTAMP" not in field:
                login_sheet.range(f"B{i+1}").value = token
            elif "USERTOKEN" in field:
                login_sheet.range(f"B{i+1}").value = usertoken
    
    current_time = dt.now().strftime('%Y-%m-%d %H:%M:%S')
    timestamp_exists = False
    for i, row in enumerate(data):
        if row and len(row) >= 2:
            field = str(row[0]).strip().upper() if row[0] else ""
            if "TOKEN_TIMESTAMP" in field or "TOKEN TIME" in field:
                login_sheet.range(f"B{i+1}").value = current_time
                timestamp_exists = True
                break
    
    if not timestamp_exists:
        last_row = len(data) + 1
        login_sheet.range(f"A{last_row}").value = "TOKEN_TIMESTAMP"
        login_sheet.range(f"B{last_row}").value = current_time
    
    login_sheet.book.save()


def is_token_valid(login_data):
    if not login_data['auth_code'] or not login_data['token'] or not login_data['usertoken']:
        return False, "Missing tokens"
    
    if login_data['token_timestamp']:
        token_age = (dt.now() - login_data['token_timestamp']).total_seconds() / 3600
        if token_age > 24:
            return False, f"Tokens are {token_age:.1f} hours old (expired)"
        print(f"   Tokens are {token_age:.1f} hours old")
    
    try:
        class TestApi(NorenApi):
            def __init__(self):
                super().__init__(
                    host="https://api.shoonya.com/NorenWClientAPI/",
                    websocket="wss://api.shoonya.com/NorenWSAPI/",
                )
        
        test_api = TestApi()
        test_api.uid = login_data['usertoken']
        test_api.token = login_data['token']
        test_api.actid = login_data['user_id']
        
        test_response = test_api.get_quotes(SPOT_EXCHANGE, str(NIFTY_SPOT_TOKEN))
        
        if test_response and 'lp' in test_response:
            print(f"   ✅ Token validation successful! Spot: {test_response.get('lp')}")
            return True, "Valid token"
        else:
            return False, "Token validation failed - no data received"
            
    except Exception as e:
        error_msg = str(e)
        if "invalid token" in error_msg.lower() or "expired" in error_msg.lower():
            return False, f"Token expired: {error_msg[:50]}"
        return False, f"Token validation error: {error_msg[:50]}"


# ----------------------------------------------------------------------
# Selenium: fetch OAuth auth code
# ----------------------------------------------------------------------
def get_auth_code_via_selenium(client_id, user_id, password, totp_secret):
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager

    login_url = (
        f"https://trade.shoonya.com/OAuthlogin/"
        f"investor-entry-level/login?api_key={client_id}&route_to={user_id}"
    )

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.maximize_window()
    driver.get(login_url)
    time.sleep(3)

    try:
        inputs = [x for x in driver.find_elements(By.TAG_NAME, "input") if x.is_displayed()]
        inputs[0].send_keys(user_id)
        inputs[1].send_keys(password)
        otp = pyotp.TOTP(totp_secret).now()
        inputs[2].send_keys(otp)
        time.sleep(1)

        for b in driver.find_elements(By.TAG_NAME, "button"):
            if "LOGIN" in b.text.upper():
                driver.execute_script("arguments[0].click();", b)
                break

        code = None
        for _ in range(150):
            url = driver.current_url
            if "#/?code=" in url or "?code=" in url:
                code = url.split("code=")[1].split("&")[0]
                break
            time.sleep(0.2)
        return code
    finally:
        driver.quit()


# ----------------------------------------------------------------------
# OAuth login with existing token check
# ----------------------------------------------------------------------
def _make_api():
    class ShoonyaApiPy(NorenApi):
        def __init__(self):
            super().__init__(
                host="https://api.shoonya.com/NorenWClientAPI/",
                websocket="wss://api.shoonya.com/NorenWSAPI/",
            )
    return ShoonyaApiPy()


def _load_saved_session(login_data):
    """Load the saved token without calling any validation REST API."""
    global api
    if not login_data.get("token") or not login_data.get("usertoken"):
        return False
    api = _make_api()
    api.uid = login_data["usertoken"]
    api.token = login_data["token"]
    api.actid = login_data.get("user_id") or login_data["usertoken"]
    print("✅ REUSING SAVED TOKEN")
    print("   🔐 No token validation API call")
    print("   🔐 No browser login")
    return True


def _login_from_auth_code(login_sheet, login_data, auth_code):
    global api
    if not auth_code:
        return False
    api = _make_api()
    print("🔄 Calling getAccessToken() using saved AUTH CODE...")
    result = api.getAccessToken(
        auth_code,
        login_data["secret_code"],
        login_data["client_id"],
        login_data["user_id"],
    )
    print(f"getAccessToken() raw result: {result}")
    if result is None:
        print("❌ Saved AUTH CODE returned None.")
        return False
    try:
        acc_tok, usrid, ref_tok, actid = result
    except Exception:
        print("❌ Unexpected getAccessToken() response.")
        return False
    update_login_data(login_sheet, auth_code, acc_tok, usrid)
    api.uid = usrid
    api.token = acc_tok
    api.actid = actid or usrid
    print(f"✅ AUTH CODE produced fresh access token for {usrid}")
    return True


def shoonya_login(login_sheet):
    """Login order: saved TOKEN -> saved AUTH CODE -> browser only if AUTH CODE returns None."""
    global api
    login_data = get_all_login_data(login_sheet)

    print("🔐 LOGIN POLICY:")
    print("   1. Saved TOKEN (no validation API call)")
    print("   2. Saved AUTH CODE")
    print("   3. Fresh browser OAuth ONLY when saved AUTH CODE returns None")

    if _load_saved_session(login_data):
        return True

    if not login_data.get("user_id") or not login_data.get("secret_code"):
        print("❌ Missing User ID or OAuth Secret Code in Login sheet.")
        return False

    if login_data.get("auth_code"):
        if _login_from_auth_code(login_sheet, login_data, login_data["auth_code"]):
            return True

    # Only this path is allowed to generate a new auth code.
    print("⚠️ Saved AUTH CODE returned None/failed.")
    print("🔄 NOW and ONLY NOW opening browser for a fresh OAuth code...")
    if not login_data.get("password") or not login_data.get("totp_secret"):
        print("❌ Password/TOTP missing; cannot perform fresh browser login.")
        return False

    api = _make_api()
    auth_code = get_auth_code_via_selenium(
        login_data["client_id"], login_data["user_id"],
        login_data["password"], login_data["totp_secret"]
    )
    if not auth_code:
        print("❌ Browser did not return an OAuth code.")
        return False
    return _login_from_auth_code(login_sheet, login_data, auth_code)


# ----------------------------------------------------------------------
# WebSocket callbacks
# ----------------------------------------------------------------------
def event_handler_quote_update(msg):
    global live_data, feed_time, last_traded_time
    
    fields = ["ts", "lp", "pc", "c", "o", "h", "l", "v", "ltq", "ltp",
              "bp1", "sp1", "bq1", "sq1", "ap", "oi", "poi", "toi",
              "tbq", "tsq", "ft", "exch_tm", "ltt"]
    message = {f: msg[f] for f in set(fields) & set(msg.keys())}
    key = msg["e"] + "|" + msg["tk"]
    live_data[key] = {**live_data.get(key, {}), **message}
    
    # IMPORTANT: Shoonya field `ts` is the trading symbol (for example Nifty 50),
    # NOT a timestamp. Never put `ts` into Feed Time or Last Traded Time.
    if 'ft' in msg:
        _ft = format_feed_datetime(msg['ft'])
        if _ft:
            feed_time = _ft
    elif 'exch_tm' in msg:
        _ft = format_feed_datetime(msg['exch_tm'])
        if _ft:
            feed_time = _ft

    # Last Traded Time comes only from ltt.
    if 'ltt' in msg:
        _ltt = format_timestamp(msg['ltt'])
        if _ltt:
            last_traded_time = _ltt

    if not last_traded_time and feed_time:
        last_traded_time = feed_time


def event_handler_order_update(msg):
    pass


def open_callback():
    global feed_opened
    feed_opened = True


def event_handler_socket_closed():
    print("Socket closed, reconnecting...")
    time.sleep(2)


def subscribe_token(exchange, token):
    api.subscribe([f"{exchange}|{token}"])


# ----------------------------------------------------------------------
# Instrument loading
# ----------------------------------------------------------------------
def load_instruments():
    global df_ins_NSE, df_ins_NFO

    print("Downloading NSE instrument list...")
    zip_file = "NSE_symbols.txt.zip"
    r = requests.get(f"https://api.shoonya.com/{zip_file}", allow_redirects=True)
    open(zip_file, "wb").write(r.content)
    df_ins_NSE = pd.read_csv(zip_file)
    os.remove(zip_file)

    print("Downloading NFO instrument list...")
    zip_file = "NFO_symbols.txt.zip"
    r = requests.get(f"https://api.shoonya.com/{zip_file}", allow_redirects=True)
    open(zip_file, "wb").write(r.content)
    df_ins_NFO = pd.read_csv(zip_file)
    df_ins_NFO["Expiry"] = pd.to_datetime(df_ins_NFO["Expiry"]).apply(lambda x: x.date())
    df_ins_NFO = df_ins_NFO.sort_values(by=["Expiry", "Symbol", "StrikePrice"])
    df_ins_NFO = df_ins_NFO.astype({"StrikePrice": str})
    os.remove(zip_file)
    print("Instrument load complete.")


def get_index_future_token():
    df_temp = df_ins_NFO[
        (df_ins_NFO.Symbol == SYMBOL) & (df_ins_NFO["Instrument"].isin(["FUTIDX"]))
    ].sort_values(by="Expiry")
    if len(df_temp) == 0:
        return None
    return df_temp.iloc[0]["Token"]


def build_option_chain_template():
    global OptionChain_template
    if any(t["symbol"] == SYMBOL for t in OptionChain_template):
        return

    df_sym = df_ins_NFO[(df_ins_NFO.Symbol == SYMBOL) & (df_ins_NFO["OptionType"].isin(["CE", "PE"]))]
    expiry_strike_list = []
    for expiry in df_sym["Expiry"].unique():
        if str(expiry) == "NaT":
            continue
        df_exp = df_sym[df_sym.Expiry == expiry]
        lot_size = df_exp.iloc[0]["LotSize"]
        strikes = []
        for strike in df_exp["StrikePrice"].unique():
            pe_row = df_exp[(df_exp.StrikePrice == strike) & (df_exp.OptionType == "PE")]
            ce_row = df_exp[(df_exp.StrikePrice == strike) & (df_exp.OptionType == "CE")]
            strikes.append({
                "strike": strike,
                "PE_Token": pe_row.iloc[0]["Token"] if len(pe_row) else "NA",
                "CE_Token": ce_row.iloc[0]["Token"] if len(ce_row) else "NA",
            })
        expiry_strike_list.append({"Expiry": expiry, "LotSize": lot_size, "Strike_list": strikes})

    OptionChain_template.append({"symbol": SYMBOL, "Expiry_Strike_token": expiry_strike_list})


def dump_available_expiries(oc_sheet):
    template = [t for t in OptionChain_template if t["symbol"] == SYMBOL][0]
    expiries = sorted(e["Expiry"] for e in template["Expiry_Strike_token"])
    oc_sheet.range("D1:D50").value = None
    oc_sheet.range("D1").options(transpose=True).value = [str(e.strftime("%d-%m-%Y")) for e in expiries]


def _nearest_upcoming_expiry(today=None):
    """Return the nearest expiry date >= today from the loaded NIFTY chain
    (falls back to the nearest overall expiry if every listed one has
    already passed)."""
    if today is None:
        today = dt.now(IST).date()
    matches = [t for t in OptionChain_template if t["symbol"] == SYMBOL]
    if not matches:
        return None
    expiries = sorted(e["Expiry"] for e in matches[0]["Expiry_Strike_token"])
    if not expiries:
        return None
    upcoming = [e for e in expiries if e >= today]
    return upcoming[0] if upcoming else expiries[-1]


def _auto_select_expiry_if_blank(oc_sheet, announce=True):
    """If B1 (Expiry) is empty, fill it with the nearest upcoming expiry so
    the tool can run without requiring a manual date first. Returns the
    (possibly auto-filled) expiry string currently in B1."""
    expiry_str = oc_sheet.range("B1").value
    if expiry_str:
        return expiry_str

    nearest = _nearest_upcoming_expiry()
    if nearest is None:
        return expiry_str

    expiry_str = nearest.strftime("%d-%m-%Y")
    oc_sheet.range("B1").value = expiry_str
    if announce:
        print(f"🗓️ B1 (Expiry) was blank - auto-selected nearest expiry {expiry_str}")
    return expiry_str


def _report_blank_config_cells(oc_sheet):
    """Print which of the core config cells (B1-B4) are currently blank, so
    it's obvious from the console why defaults are being used."""
    cell_labels = {
        "B1": "Expiry (dd-mm-yyyy)",
        "B2": "NoOfStrikes each side",
        "B3": "RefreshRate(sec)",
        "B4": "Aggregation Interval (min)",
    }
    blanks = [f"{cell} [{label}]" for cell, label in cell_labels.items()
              if not oc_sheet.range(cell).value]
    if blanks:
        print(f"⚠️ Blank config cells (defaults will be used): {', '.join(blanks)}")
    else:
        print("✅ B1-B4 config cells are all filled in.")


# ============================================================
# ATM HIGHLIGHT - GREEN BACKGROUND
# ============================================================
_last_atm_highlight_row = None


def apply_atm_highlight(oc_sheet, row_num):
    """Highlight the ATM row in the OptionChain data area.

    OptionChain headers are on row 10 and data starts on row 11.
    row_num is computed by the caller directly from atm_idx/lo (it already
    knows exactly which row ATM lands on) instead of clearing and reading
    a 990x27 block back from Excel every refresh just to find it again.
    Only the previous highlighted row (if it moved) and the new one are
    touched - two small COM calls instead of two full-block ones.
    """
    global _last_atm_highlight_row
    try:
        if row_num == _last_atm_highlight_row:
            return  # ATM strike hasn't moved rows - nothing to do

        if _last_atm_highlight_row is not None:
            oc_sheet.range(f"A{_last_atm_highlight_row}:AA{_last_atm_highlight_row}").color = None

        oc_sheet.range(f"A{row_num}:AA{row_num}").color = (0, 255, 0)
        _last_atm_highlight_row = row_num

    except Exception as e:
        print(f"⚠️ ATM highlight error: {e}")

# ----------------------------------------------------------------------
# Function to get first OTM Call and Put strikes
# ----------------------------------------------------------------------
def get_first_otm_strikes(df_full, spot_ltp, atm_strike, strike_step=50):
    """Get the first OTM Call (strike ABOVE spot) and first OTM Put (strike BELOW spot)
    from the FULL dataframe before trimming"""
    
    # Get all unique strikes from the FULL dataframe
    strikes = sorted(df_full["strike"].unique())
    
    # Find OTM Call (strike ABOVE spot, closest to spot) - LEFT side of Option Chain
    otm_call_strike = None
    
    for strike in strikes:
        if strike > spot_ltp:
            otm_call_strike = strike
            break
    
    # If no OTM Call found above spot, use ATM + step
    if otm_call_strike is None:
        otm_call_strike = atm_strike + strike_step
        # Find the closest strike above ATM
        for strike in strikes:
            if strike > atm_strike:
                otm_call_strike = strike
                break
    
    # Find OTM Put (strike BELOW spot, closest to spot) - RIGHT side of Option Chain
    otm_put_strike = None
    
    for strike in reversed(strikes):
        if strike < spot_ltp:
            otm_put_strike = strike
            break
    
    # If no OTM Put found below spot, use ATM - step
    if otm_put_strike is None:
        otm_put_strike = atm_strike - strike_step
        # Find the closest strike below ATM
        for strike in reversed(strikes):
            if strike < atm_strike:
                otm_put_strike = strike
                break
    
    return otm_call_strike, otm_put_strike


# ----------------------------------------------------------------------
# Function to store data in history sheet - APPEND new data WITH IV
# ----------------------------------------------------------------------
def store_tick_data(history_sheet, df_full, spot_ltp, atm_strike, otm_call_strike, otm_put_strike, expiry_input, future_ltp, atm_call_price, atm_put_price):
    """
    Store a market snapshot in Tick_History.

    IMPORTANT:
    1. Never write after 23:59 IST.
    2. Never write the same market snapshot twice.
       Request Time / Feed Time are NOT part of duplicate detection.
    3. If ANY real market value changes (Spot, OI, IV, Bid, Ask,
       Buy/Sell, strike, etc.), a new row is written.
    """
    global tick_counter, feed_time, request_time, last_traded_time
    global last_tick_signature

    try:
        # ============================================================
        # MARKET-CLOSE PROTECTION
        # ============================================================
        now_ist = dt.now(IST)
        market_minutes = now_ist.hour * 60 + now_ist.minute

        # Normal market is always allowed. After-market can be enabled
        # from OptionChain!B5 and stopped at OptionChain!B6.
        if market_minutes < MARKET_START:
            return

        if market_minutes > MARKET_END:
            if not POST_MARKET_ENABLED:
                return
            if market_minutes > POST_MARKET_END:
                return
            if market_minutes == POST_MARKET_END and now_ist.second > 0:
                return

        now_dt = now_ist.replace(tzinfo=None)
        request_time = now_dt.strftime("%Y-%m-%d %H:%M:%S")

        # ------------------------------------------------------------
        # Find OTM Call (above/equal ATM)
        # ------------------------------------------------------------
        otm_call_row = df_full[df_full["strike"] == otm_call_strike]

        if otm_call_row.empty:
            df_above_spot = df_full[df_full["strike"] > spot_ltp]
            if not df_above_spot.empty:
                otm_call_strike = df_above_spot.iloc[0]["strike"]
                otm_call_row = df_full[df_full["strike"] == otm_call_strike]
            else:
                return

        # ------------------------------------------------------------
        # Find OTM Put (below/equal ATM)
        # ------------------------------------------------------------
        otm_put_row = df_full[df_full["strike"] == otm_put_strike]

        if otm_put_row.empty:
            df_below_spot = df_full[df_full["strike"] < spot_ltp]
            if not df_below_spot.empty:
                otm_put_strike = df_below_spot.iloc[-1]["strike"]
                otm_put_row = df_full[df_full["strike"] == otm_put_strike]
            else:
                return

        if otm_call_row.empty or otm_put_row.empty:
            return

        call_data = otm_call_row.iloc[0]
        put_data = otm_put_row.iloc[0]

        # ------------------------------------------------------------
        # Calculate IV independently at each option's own strike
        # ------------------------------------------------------------
        call_iv = 0.0
        put_iv = 0.0

        if expiry_input and current_iv_T > 0:
            try:
                call_same = df_full[df_full["strike"] == otm_call_strike]
                if not call_same.empty:
                    r = call_same.iloc[0]
                    call_iv, _ = calculate_iv_for_strike(
                        otm_call_strike,
                        r.get("CE_lp", 0),
                        r.get("PE_lp", 0)
                    )

                put_same = df_full[df_full["strike"] == otm_put_strike]
                if not put_same.empty:
                    r = put_same.iloc[0]
                    _, put_iv = calculate_iv_for_strike(
                        otm_put_strike,
                        r.get("CE_lp", 0),
                        r.get("PE_lp", 0)
                    )
            except Exception as e:
                print(f"⚠️ OTM IV calculation error: {e}")

        call_iv = round(call_iv, 2)
        put_iv = round(put_iv, 2)

        # ------------------------------------------------------------
        # Build market-data portion of the row.
        # IMPORTANT: columns A-C (time fields) are excluded from the
        # duplicate signature because they naturally change each refresh.
        # Everything from Spot through Put Bid-Ask Diff is compared.
        # ------------------------------------------------------------
        # PCR is calculated ONLY for the selected OTM strikes in this row.
        # PCR OI        = OTM Put OI / OTM Call OI
        # PCR Change OI = OTM Put Change OI / OTM Call Change OI
        call_oi = convert_to_float(call_data.get("CE_oi", 0))
        put_oi = convert_to_float(put_data.get("PE_oi", 0))
        call_chg_oi = convert_to_float(call_data.get("CE_coi", 0))
        put_chg_oi = convert_to_float(put_data.get("PE_coi", 0))

        pcr_oi = (put_oi / call_oi) if call_oi != 0 else 0.0
        pcr_change_oi = (put_chg_oi / call_chg_oi) if call_chg_oi != 0 else 0.0

        market_values = [
            convert_to_float(spot_ltp),
            convert_to_float(atm_strike),
            convert_to_float(otm_call_strike),
            convert_to_float(call_data.get("CE_lp", 0)),
            convert_to_float(call_data.get("CE_v", 0)),
            call_iv,
            convert_to_float(call_data.get("CE_oi", 0)),
            convert_to_float(call_data.get("CE_coi", 0)),
            convert_to_float(call_data.get("CE_total_buy", 0)),
            convert_to_float(call_data.get("CE_total_sell", 0)),
            convert_to_float(call_data.get("CE_total_buy", 0)) -
            convert_to_float(call_data.get("CE_total_sell", 0)),
            convert_to_float(call_data.get("CE_bp1", 0)),
            convert_to_float(call_data.get("CE_sp1", 0)),
            convert_to_float(call_data.get("CE_sp1", 0)) -
            convert_to_float(call_data.get("CE_bp1", 0)),
            convert_to_float(otm_put_strike),
            convert_to_float(put_data.get("PE_lp", 0)),
            convert_to_float(put_data.get("PE_v", 0)),
            put_iv,
            convert_to_float(put_data.get("PE_oi", 0)),
            convert_to_float(put_data.get("PE_coi", 0)),
            convert_to_float(put_data.get("PE_total_buy", 0)),
            convert_to_float(put_data.get("PE_total_sell", 0)),
            convert_to_float(put_data.get("PE_total_buy", 0)) -
            convert_to_float(put_data.get("PE_total_sell", 0)),
            convert_to_float(put_data.get("PE_bp1", 0)),
            convert_to_float(put_data.get("PE_sp1", 0)),
            convert_to_float(put_data.get("PE_sp1", 0)) -
            convert_to_float(put_data.get("PE_bp1", 0)),
            pcr_oi,
            pcr_change_oi,
        ]

        # ------------------------------------------------------------
        # DUPLICATE SNAPSHOT PROTECTION
        # ------------------------------------------------------------
        # Convert to a stable tuple. Rounded numeric values prevent tiny
        # floating-point representation differences from creating fake ticks.
        signature = tuple(
            round(float(v), 8) if isinstance(v, (int, float, np.number)) else str(v)
            for v in market_values
        )

        if last_tick_signature == signature:
            # Same market data as the last row. Do NOT write another row.
            return

        # ------------------------------------------------------------
        # Timestamp fields are only created AFTER duplicate check.
        # ------------------------------------------------------------
        if not feed_time:
            feed_time = now_dt.strftime("%H:%M:%S")
        if not last_traded_time or last_traded_time in ("Nifty 50", "NIFTY 50"):
            last_traded_time = feed_time if feed_time else now_dt.strftime("%H:%M:%S")

        # Use current exchange/feed timestamp when available; otherwise
        # keep the existing value maintained by the main loop.
        headers = [
            "Feed Time", "Request Time", "Last Traded Time",
            "Spot", "ATM Strike",
            "OTM Call Strike", "Call LTP", "Call Volume", "Call IV", "Call OI", "Call Change OI",
            "Call Total Buy", "Call Total Sell", "Call Buy-Sell Diff",
            "Call Bid", "Call Ask", "Call Bid-Ask Diff",
            "OTM Put Strike", "Put LTP", "Put Volume", "Put IV", "Put OI", "Put Change OI",
            "Put Total Buy", "Put Total Sell", "Put Buy-Sell Diff",
            "Put Bid", "Put Ask", "Put Bid-Ask Diff",
            "PCR OI", "PCR Change OI"
        ]

        # Ensure headers exist.
        try:
            if history_sheet.range("A1").value != "Feed Time":
                history_sheet.range("A1:AE1").value = headers
                history_sheet.range("A:C").number_format = "@"
        except Exception:
            history_sheet.range("A1:AE1").value = headers

        row_data = [
            feed_time,
            now_dt.strftime("%Y-%m-%d %H:%M:%S"),
            last_traded_time,
            *market_values,
        ]

        # ------------------------------------------------------------
        # APPEND EXACTLY ONE NEW ROW
        # ------------------------------------------------------------
        try:
            last_row = get_last_actual_row(history_sheet, "A", 1)
            next_row = max(2, last_row + 1)
        except Exception:
            next_row = tick_counter + 2

        # Force all three time columns to TEXT before writing. This prevents
        # Excel from converting 20:03:46 into the 1900 date system.
        history_sheet.range(f"A{next_row}:C{next_row}").number_format = "@"
        history_sheet.range(f"A{next_row}:C{next_row}").value = [[
            str(row_data[0] or ""),
            str(row_data[1] or ""),
            str(row_data[2] or ""),
        ]]
        history_sheet.range(f"D{next_row}:AE{next_row}").number_format = "#,##0.00"
        history_sheet.range(f"D{next_row}:AE{next_row}").value = [[*row_data[3:]]]

        tick_counter += 1
        last_tick_signature = signature

        # Autofit only occasionally; never every refresh.
        if tick_counter == 1 or tick_counter % 50 == 0:
            try:
                history_sheet.autofit()
            except Exception:
                pass

        print(
            f"✅ Tick {tick_counter} written at row {next_row} | "
            f"{now_dt.strftime('%H:%M:%S')} | Spot={market_values[0]:.2f}"
        )

    except Exception as e:
        print(f"⚠️ Error storing tick data: {e}")
        import traceback
        traceback.print_exc()


# ----------------------------------------------------------------------
# Main option chain loop
# ----------------------------------------------------------------------

def get_nse_strike_classification(df_full, spot_ltp):
    try:
        spot = float(spot_ltp or 0)
    except Exception:
        spot = 0.0

    strikes = []
    try:
        col = next((c for c in ["Strike", "Strike Price", "strike", "strprc"]
                    if c in df_full.columns), None)
        if col:
            strikes = sorted({float(x) for x in df_full[col].tolist()
                              if x is not None and str(x).strip() != ""})
    except Exception:
        pass

    if not strikes:
        return {"atm": None, "ce_itm": set(), "ce_otm": set(),
                "pe_itm": set(), "pe_otm": set()}

    atm = min(strikes, key=lambda x: abs(x - spot))
    return {
        "atm": atm,
        "ce_itm": {x for x in strikes if x < atm},
        "ce_otm": {x for x in strikes if x > atm},
        "pe_itm": {x for x in strikes if x > atm},
        "pe_otm": {x for x in strikes if x < atm},
    }

def run_option_chain(wb, oc_sheet, history_sheet, candle_sheet):
    global tick_counter, feed_time, request_time, last_traded_time, last_aggregation_time, AGGREGATION_INTERVAL, POST_MARKET_ENABLED, POST_MARKET_END
        
    fut_token = get_index_future_token()
    if fut_token is not None:
        subscribe_token(EXCHANGE, fut_token)
    subscribe_token(SPOT_EXCHANGE, NIFTY_SPOT_TOKEN)

    build_option_chain_template()
    dump_available_expiries(oc_sheet)

    # Auto-fill Expiry (B1) with the nearest upcoming expiry if it's blank,
    # and report which of the core config cells (B1-B4) still need attention.
    _auto_select_expiry_if_blank(oc_sheet)
    _report_blank_config_cells(oc_sheet)

    pre_expiry = None
    pre_no_of_strike = None
    refresh_rate = 1
    last_aggregation_time = None
    
    # Check Tick_History status at startup
    check_tick_history_status(history_sheet)

    while True:
        try:
            # Check for updated aggregation interval from Excel
            agg_interval = oc_sheet.range("B4").value
            if agg_interval and isinstance(agg_interval, (int, float)) and agg_interval > 0:
                if int(agg_interval) != AGGREGATION_INTERVAL:
                    AGGREGATION_INTERVAL = int(agg_interval)
                    print(f"✅ Aggregation interval updated to {AGGREGATION_INTERVAL} minutes")
                    # No title row to update anymore - interval is shown on
                    # OptionChain!B4 and headers now sit in Candel row 1.
            
            expiry_str = oc_sheet.range("B1").value
            if not expiry_str:
                expiry_str = _auto_select_expiry_if_blank(oc_sheet)

            try:
                expiry_input = parse_date(expiry_str)
            except ValueError as e:
                oc_sheet.range("C1").value = f"Invalid date format: {expiry_str}"
                print(f"❌ Date parsing error: {e}")
                time.sleep(2)
                continue

            no_of_strike = int(oc_sheet.range("B2").value or NUMBER_OF_STRIKES)
            refresh_rate = max(0.05, float(oc_sheet.range("B3").value or 0.10))
            iv_underlying_mode = str(oc_sheet.range("B7").value or "SPOT").strip().upper()

            # Optional post-market window
            pm_value = str(oc_sheet.range("B5").value or "ON").strip().upper()
            POST_MARKET_ENABLED = pm_value in ("ON", "YES", "TRUE", "1")
            end_value = str(oc_sheet.range("B6").value or "23:59").strip()
            try:
                hh, mm = [int(x) for x in end_value.split(":")[:2]]
                POST_MARKET_END = hh * 60 + mm
            except Exception:
                POST_MARKET_END = 23 * 60 + 59

            if not expiry_input:
                time.sleep(1)
                continue

            template = [t for t in OptionChain_template if t["symbol"] == SYMBOL][0]["Expiry_Strike_token"]
            match = [e for e in template if e["Expiry"] == expiry_input]
            if not match:
                oc_sheet.range("C1").value = f"Expiry {expiry_input} not found - pick from D column"
                time.sleep(1)
                continue

            lot_size = match[0]["LotSize"]
            strikes = match[0]["Strike_list"]

            if SYMBOL not in subs_lst:
                print(f"Subscribing {len(strikes)} strikes for {SYMBOL} {expiry_input}...")
                for s in strikes:
                    if s["PE_Token"] != "NA":
                        subscribe_token(EXCHANGE, s["PE_Token"])
                    if s["CE_Token"] != "NA":
                        subscribe_token(EXCHANGE, s["CE_Token"])
                subs_lst.append(SYMBOL)

            # Spot/future tokens are already subscribed over the websocket at
            # the top of run_option_chain, so pull them from live_data (like
            # every other field) instead of blocking on a REST round trip
            # every single loop iteration. Only fall back to REST if the
            # websocket hasn't delivered a tick for that token yet (cold start).
            spot_key = f"{SPOT_EXCHANGE}|{NIFTY_SPOT_TOKEN}"
            spot_ltp = convert_to_float(get_field(spot_key, "lp", 0))
            if spot_ltp == 0:
                spot_ltp = convert_to_float(api.get_quotes(SPOT_EXCHANGE, str(NIFTY_SPOT_TOKEN)).get("lp"))

            if fut_token:
                fut_key = f"{EXCHANGE}|{fut_token}"
                future_ltp = convert_to_float(get_field(fut_key, "lp", 0))
                if future_ltp == 0:
                    future_ltp = convert_to_float(api.get_quotes(EXCHANGE, str(fut_token)).get("lp"))
            else:
                future_ltp = spot_ltp

            rows = []
            for s in strikes:
                strike = convert_to_float(s["strike"])
                ce_key = f"{EXCHANGE}|{s['CE_Token']}"
                pe_key = f"{EXCHANGE}|{s['PE_Token']}"

                ce_oi = convert_to_float(get_field(ce_key, "oi", 0))
                ce_poi = convert_to_float(get_field(ce_key, "poi", 0))
                pe_oi = convert_to_float(get_field(pe_key, "oi", 0))
                pe_poi = convert_to_float(get_field(pe_key, "poi", 0))

                ce_total_buy = convert_to_float(get_field(ce_key, "tbq", 0))
                ce_total_sell = convert_to_float(get_field(ce_key, "tsq", 0))
                pe_total_buy = convert_to_float(get_field(pe_key, "tbq", 0))
                pe_total_sell = convert_to_float(get_field(pe_key, "tsq", 0))
                
                ce_bp1 = convert_to_float(get_field(ce_key, "bp1", 0))
                ce_sp1 = convert_to_float(get_field(ce_key, "sp1", 0))
                pe_bp1 = convert_to_float(get_field(pe_key, "bp1", 0))
                pe_sp1 = convert_to_float(get_field(pe_key, "sp1", 0))
                
                ce_lp = convert_to_float(get_field(ce_key, "lp", 0))
                pe_lp = convert_to_float(get_field(pe_key, "lp", 0))

                rows.append({
                    "strike": strike,
                    "CE_Token": s['CE_Token'],
                    "PE_Token": s['PE_Token'],
                    "CE_oi": ce_oi / lot_size if lot_size > 0 else 0,
                    "CE_coi": (ce_oi - ce_poi) / lot_size if lot_size > 0 else 0,
                    "CE_v": convert_to_float(get_field(ce_key, "v", 0)) / lot_size if lot_size > 0 else 0,
                    "CE_lp": ce_lp,
                    "CE_pc": get_field(ce_key, "pc", "-"),
                    "CE_total_buy": ce_total_buy,
                    "CE_total_sell": ce_total_sell,
                    "CE_bp1": ce_bp1,
                    "CE_sp1": ce_sp1,
                    "PE_total_buy": pe_total_buy,
                    "PE_total_sell": pe_total_sell,
                    "PE_bp1": pe_bp1,
                    "PE_sp1": pe_sp1,
                    "PE_pc": get_field(pe_key, "pc", "-"),
                    "PE_lp": pe_lp,
                    "PE_v": convert_to_float(get_field(pe_key, "v", 0)) / lot_size if lot_size > 0 else 0,
                    "PE_coi": (pe_oi - pe_poi) / lot_size if lot_size > 0 else 0,
                    "PE_oi": pe_oi / lot_size if lot_size > 0 else 0,
                })

            df_full = pd.DataFrame(rows).sort_values(by="strike").reset_index(drop=True)

            # Find the index of the strike closest to spot
            df_full["strike_diff"] = abs(df_full["strike"] - spot_ltp)
            atm_idx = df_full["strike_diff"].idxmin()
            atm_strike = df_full.loc[atm_idx, "strike"]

            # Get ATM data
            atm_ce_price = df_full.loc[atm_idx, "CE_lp"]
            atm_pe_price = df_full.loc[atm_idx, "PE_lp"]

            # IMPORTANT: Refresh IV calculator on EVERY cycle.
            # Spot, future, ATM prices and time change continuously. Keeping
            # the calculator only for the whole expiry makes IV stale.
            print("🔄 Updating IV calculator with current market snapshot...")
            if not init_iv_calculator(
                spot_ltp, future_ltp, atm_strike,
                atm_ce_price, atm_pe_price, expiry_input,
                underlying_mode=iv_underlying_mode
            ):
                print("⚠️ IV calculator update failed; IV will be 0 for this refresh.")
            pre_expiry = expiry_input

            # Get first OTM Call and Put strikes from FULL dataframe
            otm_call_strike, otm_put_strike = get_first_otm_strikes(df_full, spot_ltp, atm_strike, STRIKE_STEP)

            # Trim to N strikes each side of ATM for display
            lo = max(0, atm_idx - no_of_strike)
            hi = min(len(df_full), atm_idx + no_of_strike + 1)
            df_display = df_full.iloc[lo:hi].reset_index(drop=True)

            request_time = dt.now().strftime('%H:%M:%S')
            
            total_call_buy = df_display["CE_total_buy"].sum()
            total_call_sell = df_display["CE_total_sell"].sum()
            total_put_buy = df_display["PE_total_buy"].sum()
            total_put_sell = df_display["PE_total_sell"].sum()
            
            call_bid_ask_diff = (df_display["CE_sp1"] - df_display["CE_bp1"]).mean() if len(df_display) > 0 else 0
            put_bid_ask_diff = (df_display["PE_sp1"] - df_display["PE_bp1"]).mean() if len(df_display) > 0 else 0
            
            call_buy_sell_diff = total_call_buy - total_call_sell
            put_buy_sell_diff = total_put_buy - total_put_sell
            
            # Calculate IV for each strike in display
            call_iv_list = []
            put_iv_list = []
            
            for idx, row in df_display.iterrows():
                strike = row["strike"]
                call_price = row["CE_lp"]
                put_price = row["PE_lp"]
                call_iv, put_iv = calculate_iv_for_strike(strike, call_price, put_price)
                call_iv_list.append(call_iv)
                put_iv_list.append(put_iv)
            
            df_final = pd.DataFrame({
                "Feed Time": [feed_time] * len(df_display),
                "Request Time": [request_time] * len(df_display),
                "Last Traded Time": [last_traded_time] * len(df_display),
                "Spot": [spot_ltp] * len(df_display),
                "ATM Strike": [atm_strike] * len(df_display),
                "Call Strike": df_display["strike"],
                "Call LTP": df_display["CE_lp"],
                "Call Volume": df_display["CE_v"],
                "Call IV": call_iv_list,
                "Call OI": df_display["CE_oi"],
                "Call Change OI": df_display["CE_coi"],
                "Call Total Buy": df_display["CE_total_buy"],
                "Call Total Sell": df_display["CE_total_sell"],
                "Call Buy-Sell Diff": df_display["CE_total_buy"] - df_display["CE_total_sell"],
                "Call Bid": df_display["CE_bp1"],
                "Call Ask": df_display["CE_sp1"],
                "Call Bid-Ask Diff": df_display["CE_sp1"] - df_display["CE_bp1"],
                "Put Strike": df_display["strike"],
                "Put LTP": df_display["PE_lp"],
                "Put Volume": df_display["PE_v"],
                "Put IV": put_iv_list,
                "Put OI": df_display["PE_oi"],
                "Put Change OI": df_display["PE_coi"],
                "Put Total Buy": df_display["PE_total_buy"],
                "Put Total Sell": df_display["PE_total_sell"],
                "Put Buy-Sell Diff": df_display["PE_total_buy"] - df_display["PE_total_sell"],
                "Put Bid": df_display["PE_bp1"],
                "Put Ask": df_display["PE_sp1"],
                "Put Bid-Ask Diff": df_display["PE_sp1"] - df_display["PE_bp1"],
                "PCR OI": np.where(
                    pd.to_numeric(df_display["CE_oi"], errors="coerce").fillna(0) != 0,
                    pd.to_numeric(df_display["PE_oi"], errors="coerce").fillna(0) /
                    pd.to_numeric(df_display["CE_oi"], errors="coerce").fillna(0),
                    0.0
                ),
                "PCR Change OI": np.where(
                    pd.to_numeric(df_display["CE_coi"], errors="coerce").fillna(0) != 0,
                    pd.to_numeric(df_display["PE_coi"], errors="coerce").fillna(0) /
                    pd.to_numeric(df_display["CE_coi"], errors="coerce").fillna(0),
                    0.0
                ),
            })

            if pre_expiry != expiry_input or pre_no_of_strike != no_of_strike:
                oc_sheet.range("A11:AA1000").value = None
                oc_sheet.range("A11:AA1000").color = None
                global _last_atm_highlight_row
                _last_atm_highlight_row = None
                pre_expiry, pre_no_of_strike = expiry_input, no_of_strike

            oc_sheet.range("A10").options(index=False, header=True).value = df_final

            # atm_idx is the position in df_full; lo is where df_display was
            # sliced from - so this is the exact displayed row without any
            # Excel read-back.
            apply_atm_highlight(oc_sheet, 11 + (atm_idx - lo))
            
            # Store tick data using FULL dataframe with IV calculation
            store_tick_data(history_sheet, df_full, spot_ltp, atm_strike, otm_call_strike, otm_put_strike, expiry_input, future_ltp, atm_ce_price, atm_pe_price)
            
            # ============================================================
            # AGGREGATE CANDLES - APPEND new candles
            # ============================================================
            current_time = dt.now()

            # ULTRA-FAST CANDLE UPDATE:
            # Update the current 1-minute candle immediately after each
            # NEW market snapshot. Do not wait another 60 seconds.
            global last_candle_tick_counter
            if tick_counter >= 1 and tick_counter != last_candle_tick_counter:
                aggregate_candles(
                    history_sheet,
                    candle_sheet,
                    AGGREGATION_INTERVAL
                )
                last_candle_tick_counter = tick_counter
            
            oc_sheet.range("C1").value = (
                f"{SYMBOL} Spot={spot_ltp:.1f}  ATM={atm_strike}  "
                f"OTM Call={otm_call_strike} (ABOVE)  OTM Put={otm_put_strike} (BELOW)  "
                f"Ticks={tick_counter}  Aggregation={AGGREGATION_INTERVAL}min  "
                f"Feed={feed_time}  Req={request_time}  LTT={last_traded_time}  "
                f"IV=LIVE({current_iv_underlying_mode}={current_iv_underlying:.2f}) "
                f"DTE={current_iv_dte_days} T={current_iv_T:.8f} r=10%"
            )
            wb.save()

        except Exception as e:
            print(f"Loop exception: {e}")
            import traceback
            traceback.print_exc()

        time.sleep(refresh_rate)


# ----------------------------------------------------------------------
if __name__ == "__main__":
    print(f"Starting {SYMBOL} option chain tool...")
    print(f"Excel file: {WORKBOOK_NAME}")
    print(f"Symbol: {SYMBOL}")
    print(f"Number of strikes each side: {NUMBER_OF_STRIKES}")
    print(f"Default candle interval: {AGGREGATION_INTERVAL} minutes")
    print("ULTRA-FAST: Candel updates on every NEW tick; refresh >= 0.05 sec")
    print("-" * 50)
    
    wb, login_sheet, oc_sheet, history_sheet, candle_sheet = get_or_create_workbook()

    if not shoonya_login(login_sheet):
        print("❌ Login failed. Check the Login sheet.")
        sys.exit(1)

    load_instruments()

    # Start WebSocket. If the saved token cannot open it, use the saved
    # AUTH CODE first. Only if getAccessToken(AUTH_CODE) returns None do
    # we generate a new browser OAuth code.
    def start_ws_and_wait(seconds=15):
        global feed_opened
        feed_opened = False
        api.start_websocket(
            order_update_callback=event_handler_order_update,
            subscribe_callback=event_handler_quote_update,
            socket_open_callback=open_callback,
            socket_close_callback=event_handler_socket_closed,
        )
        for _ in range(seconds * 10):
            if feed_opened:
                return True
            time.sleep(0.1)
        return feed_opened

    print("🔌 Starting WebSocket with current session...")
    if not start_ws_and_wait(15):
        print("⚠️ Saved-token WebSocket did not open.")
        login_data = get_all_login_data(login_sheet)
        print("🔐 Trying SAVED AUTH CODE before any browser login...")
        if login_data.get("auth_code") and _login_from_auth_code(login_sheet, login_data, login_data["auth_code"]):
            if not start_ws_and_wait(15):
                print("❌ WebSocket still did not open after saved AUTH CODE.")
                print("   Browser login is NOT started unless AUTH CODE returned None.")
        elif login_data.get("auth_code"):
            # IMPORTANT: getAccessToken() returned None.  Do NOT call shoonya_login()
            # here because shoonya_login() is intentionally TOKEN-FIRST and would
            # simply reuse the same rejected token again.
            print("🚨 SAVED AUTH CODE WAS REJECTED (getAccessToken returned None).")
            print("🔄 Generating ONE NEW OAuth AUTH CODE now...")
            print("--------------------------------------------------")
            try:
                if not login_data.get("password") or not login_data.get("totp_secret"):
                    print("❌ Password/TOTP missing; cannot generate new OAuth code.")
                    sys.exit(1)

                new_api = _make_api()
                auth_code_new = get_auth_code_via_selenium(
                    login_data["client_id"],
                    login_data["user_id"],
                    login_data["password"],
                    login_data["totp_secret"],
                )
                if not auth_code_new:
                    print("❌ Browser did not return a new AUTH CODE.")
                    sys.exit(1)

                print(f"   Auth code obtained: {auth_code_new[:20]}...")
                # Exchange the NEW code on the NEW API object.
                login_data = get_all_login_data(login_sheet)
                if not _login_from_auth_code(login_sheet, login_data, auth_code_new):
                    print("❌ New AUTH CODE could not produce an access token.")
                    sys.exit(1)

                print("🔌 Starting WebSocket with NEW OAuth session...")
                if not start_ws_and_wait(15):
                    print("❌ WebSocket did not open after NEW OAuth login.")
                    sys.exit(1)
            except Exception as e:
                print(f"❌ Fresh OAuth recovery failed: {e}")
                import traceback
                traceback.print_exc()
                sys.exit(1)
        else:
            # No saved AUTH CODE: generate one fresh code immediately.
            print("⚠️ No saved AUTH CODE available.")
            print("🔄 Generating ONE NEW OAuth AUTH CODE now...")
            login_data = get_all_login_data(login_sheet)
            if not login_data.get("password") or not login_data.get("totp_secret"):
                print("❌ Password/TOTP missing; cannot generate new OAuth code.")
                sys.exit(1)
            auth_code_new = get_auth_code_via_selenium(
                login_data["client_id"], login_data["user_id"],
                login_data["password"], login_data["totp_secret"]
            )
            if not auth_code_new:
                print("❌ Browser did not return a new AUTH CODE.")
                sys.exit(1)
            login_data = get_all_login_data(login_sheet)
            if not _login_from_auth_code(login_sheet, login_data, auth_code_new):
                print("❌ New AUTH CODE could not produce an access token.")
                sys.exit(1)
            if not start_ws_and_wait(15):
                print("❌ WebSocket did not open after NEW OAuth login.")
                sys.exit(1)

    if feed_opened:
        print("✅ WebSocket connected. Enter expiry/strike count in the OptionChain sheet.")
    else:
        print("⚠️ WebSocket is not open yet; continuing without creating another sheet/login.")
    print("-" * 50)

    run_option_chain(wb, oc_sheet, history_sheet, candle_sheet)
