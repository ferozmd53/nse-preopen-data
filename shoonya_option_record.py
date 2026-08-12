
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
            "selenium", "webdriver_manager", "openpyxl"]:
    try:
        __import__(pkg)
    except ImportError:
        os.system(f"{sys.executable} -m pip install -U {pkg}")

import pandas as pd
import numpy as np
import pyotp
import xlwings as xw
import requests

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
WORKBOOK_NAME = "STRIKE_OPTION.xlsx"
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
MARKET_START = 9 * 60 + 14  # 9:15 AM in minutes
MARKET_END = 15 * 60 + 30   # 3:30 PM in minutes

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

# Global variables for NSE IV calculation
current_expiry_date = None
current_spot = 0.0
current_future = 0.0
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
    """Format timestamp to HH:MM:SS format"""
    try:
        if ts_value is None or ts_value == "":
            return ""
        
        if isinstance(ts_value, (int, float)):
            ist_dt = epoch_to_ist(ts_value)
            if ist_dt:
                return ist_dt.strftime('%H:%M:%S')
        
        if isinstance(ts_value, str):
            try:
                epoch_val = float(ts_value)
                ist_dt = epoch_to_ist(epoch_val)
                if ist_dt:
                    return ist_dt.strftime('%H:%M:%S')
            except:
                pass
            
            if ts_value.startswith('NIFTY'):
                return ""
            
            try:
                if ':' in ts_value and '.' in ts_value:
                    parts = ts_value.split(':')
                    if len(parts) == 3:
                        return f"{int(parts[0]):02d}:{int(parts[1]):02d}:{int(parts[2].split('.')[0]):02d}"
                    elif len(parts) == 2:
                        minutes = int(parts[0])
                        seconds = int(parts[1].split('.')[0])
                        hours = minutes // 60
                        minutes = minutes % 60
                        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                elif ':' in ts_value:
                    parts = ts_value.split(':')
                    if len(parts) == 3:
                        return f"{int(parts[0]):02d}:{int(parts[1]):02d}:{int(parts[2]):02d}"
                    elif len(parts) == 2:
                        minutes = int(parts[0])
                        seconds = int(parts[1])
                        hours = minutes // 60
                        minutes = minutes % 60
                        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                return ts_value
            except:
                pass
        
        return ""
    except Exception as e:
        print(f"⚠️ Error formatting timestamp {ts_value}: {e}")
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


def _solve_iv(option_price, pricing_function, lower_bound, upper_bound, max_iter=120):
    """Solve IV by bisection without requiring scipy."""
    if option_price <= 0:
        return 0.0

    lo = 1e-8
    hi = 5.0  # 500% annualized volatility is a safe numerical ceiling.
    p_lo = pricing_function(lo)
    p_hi = pricing_function(hi)

    # Market price must be inside the model's no-arbitrage range.
    if option_price < lower_bound - 1e-7 or option_price > upper_bound + 1e-7:
        return 0.0

    # If the option is essentially at its lower bound, IV is not stable.
    if abs(option_price - p_lo) < 1e-9:
        return 0.0

    if p_lo > option_price or p_hi < option_price:
        return 0.0

    for _ in range(max_iter):
        mid = (lo + hi) / 2.0
        p_mid = pricing_function(mid)
        if abs(p_mid - option_price) < 1e-9:
            return mid * 100.0
        if p_mid < option_price:
            lo = mid
        else:
            hi = mid

    return ((lo + hi) / 2.0) * 100.0


def init_iv_calculator(spot_ltp, future_ltp, atm_strike, atm_call_price, atm_put_price, expiry_date):
    """Refresh the NSE IV engine with the current market snapshot."""
    global current_expiry_date, current_spot, current_future
    global current_atm_strike, current_atm_call_price, current_atm_put_price
    global current_iv_dte_days, current_iv_T, current_iv_timestamp

    current_spot = convert_to_float(spot_ltp)
    current_future = convert_to_float(future_ltp)
    current_atm_strike = convert_to_float(atm_strike)
    current_atm_call_price = convert_to_float(atm_call_price)
    current_atm_put_price = convert_to_float(atm_put_price)
    current_expiry_date = expiry_date
    current_iv_timestamp = dt.now(IST).replace(tzinfo=None)

    current_iv_dte_days = _nse_calendar_dte(expiry_date, current_iv_timestamp)
    current_iv_T = current_iv_dte_days / NSE_IV_DAYS_PER_YEAR

    if current_spot <= 0 or current_iv_dte_days <= 0:
        return False

    print(
        f"🔄 NSE IV engine: Spot={current_spot:.2f} "
        f"Expiry={expiry_date} DTE={current_iv_dte_days} "
        f"T={current_iv_T:.8f} r={NSE_IV_RATE:.2%}"
    )
    return True


def calculate_iv_for_strike(strike_price, call_price, put_price):
    """Calculate independent NSE-style CE IV and PE IV for one strike."""
    strike = convert_to_float(strike_price)
    call = convert_to_float(call_price)
    put = convert_to_float(put_price)

    if (
        current_spot <= 0 or strike <= 0 or current_iv_T <= 0
        or (call <= 0 and put <= 0)
    ):
        return 0.0, 0.0

    discount_factor = math.exp(-NSE_IV_RATE * current_iv_T)

    # European Black-Scholes no-arbitrage bounds using NSE's 10% rate.
    call_lower = max(current_spot - strike * discount_factor, 0.0)
    call_upper = current_spot
    put_lower = max(strike * discount_factor - current_spot, 0.0)
    put_upper = strike * discount_factor

    call_iv = 0.0
    put_iv = 0.0

    if call > 0:
        call_iv = _solve_iv(
            call,
            lambda sigma: _bs_call_price(
                current_spot, strike, NSE_IV_RATE, current_iv_T, sigma
            ),
            call_lower,
            call_upper,
        )

    if put > 0:
        put_iv = _solve_iv(
            put,
            lambda sigma: _bs_put_price(
                current_spot, strike, NSE_IV_RATE, current_iv_T, sigma
            ),
            put_lower,
            put_upper,
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
        oc_sheet.range("C1").value = "Available expiries -->"
        oc_sheet.range("G1").value = f"Symbol: {SYMBOL}"
        oc_sheet.range("H1").value = f"Token: {NIFTY_SPOT_TOKEN}"
        oc_sheet.range("I4").value = "Enter minutes (1, 2, 3, 5, 10, 15, 30, 60)"
    else:
        oc_sheet = wb.sheets["OptionChain"]
        oc_sheet.range("G1").value = f"Symbol: {SYMBOL}"
        oc_sheet.range("H1").value = f"Token: {NIFTY_SPOT_TOKEN}"
        oc_sheet.range("B2").value = NUMBER_OF_STRIKES
        # Read aggregation interval from Excel
        agg_interval = oc_sheet.range("B4").value
        if agg_interval and isinstance(agg_interval, (int, float)) and agg_interval > 0:
            AGGREGATION_INTERVAL = int(agg_interval)
            print(f"✅ Aggregation interval set to {AGGREGATION_INTERVAL} minutes")

    history_sheet_name = "Tick_History"
    if history_sheet_name.lower() not in sheet_names:
        history_sheet = wb.sheets.add(history_sheet_name)
        # Create headers for new sheet with IV
        headers = [
            "Feed Time", "Request Time", "Last Traded Time",
            "Spot", "ATM Strike",
            "OTM Call Strike", "Call IV", "Call OI", "Call Change OI",
            "Call Total Buy", "Call Total Sell", "Call Buy-Sell Diff",
            "Call Bid", "Call Ask", "Call Bid-Ask Diff",
            "OTM Put Strike", "Put IV", "Put OI", "Put Change OI",
            "Put Total Buy", "Put Total Sell", "Put Buy-Sell Diff",
            "Put Bid", "Put Ask", "Put Bid-Ask Diff",
            "PCR OI", "PCR Change OI"
        ]
        history_sheet.range("A1:AA1").value = headers
        history_sheet.range("B:B").number_format = "yyyy-mm-dd hh:mm:ss"
        tick_counter = 0
    else:
        history_sheet = wb.sheets[history_sheet_name]
        # Check if headers exist, if not create them
        try:
            header_check = history_sheet.range("A1").value
            if header_check is None or header_check == "" or header_check != "Feed Time":
                print("📝 Creating missing headers in Tick_History...")
                headers = [
                    "Feed Time", "Request Time", "Last Traded Time",
                    "Spot", "ATM Strike",
                    "OTM Call Strike", "Call IV", "Call OI", "Call Change OI",
                    "Call Total Buy", "Call Total Sell", "Call Buy-Sell Diff",
                    "Call Bid", "Call Ask", "Call Bid-Ask Diff",
                    "OTM Put Strike", "Put IV", "Put OI", "Put Change OI",
                    "Put Total Buy", "Put Total Sell", "Put Buy-Sell Diff",
                    "Put Bid", "Put Ask", "Put Bid-Ask Diff",
                    "PCR OI", "PCR Change OI"
                ]
                history_sheet.range("A1:AA1").value = headers
                history_sheet.range("B:B").number_format = "yyyy-mm-dd hh:mm:ss"
        except Exception as e:
            print(f"⚠️ Error checking headers: {e}")

    # Always keep the Tick_History header structure current.
    # This does not rewrite or delete existing tick data.
    try:
        history_sheet.range("A1:AA1").value = headers
        history_sheet.range("B:B").number_format = "yyyy-mm-dd hh:mm:ss"
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
    candle_sheet_name = "Candles"
    if candle_sheet_name.lower() not in sheet_names:
        candle_sheet = wb.sheets.add(candle_sheet_name)
        # Create headers for new candle sheet with IV
        candle_headers = [
            "Time", "Close",
            "Call Strike", "Call IV", "Call OI", "Call Change OI", "Call Bid-Ask Avg",
            "Call Total Buy", "Call Total Sell", "Call Buy-Sell Diff",
            "Put Strike", "Put IV", "Put OI", "Put Change OI", "Put Bid-Ask Avg",
            "Put Total Buy", "Put Total Sell", "Put Buy-Sell Diff",
            "PCR OI", "PCR Change OI",
            "Ticks"
        ]
        candle_sheet.range("A2").value = candle_headers
        candle_sheet.range("A1").value = f"{AGGREGATION_INTERVAL}-Minute Candle Data"
    else:
        candle_sheet = wb.sheets[candle_sheet_name]
        # Check if headers exist
        try:
            header_check = candle_sheet.range("A2").value
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
                candle_sheet.range("A2:U2").value = [candle_headers]
        except Exception:
            pass
        candle_sheet.range("A1").value = f"{AGGREGATION_INTERVAL}-Minute Candle Data"

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
        f"Real Candles last row: {get_last_actual_row(candle_sheet, 'A', 2)}"
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

        values = history_sheet.range(f"A1:AA{last_row}").value
        if not values or len(values) < 2:
            return 0

        headers = values[0]
        data_rows = values[1:]
        df = pd.DataFrame(data_rows, columns=headers)

        required = [
            "Request Time", "Spot", "OTM Call Strike", "Call IV", "Call OI", "Call Change OI",
            "Call Total Buy", "Call Total Sell", "Call Bid", "Call Ask",
            "OTM Put Strike", "Put IV", "Put OI", "Put Change OI", "Put Total Buy",
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
            "Spot", "OTM Call Strike", "Call IV", "Call OI", "Call Change OI",
            "Call Total Buy", "Call Total Sell", "Call Bid", "Call Ask",
            "OTM Put Strike", "Put IV", "Put OI", "Put Change OI", "Put Total Buy",
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
        # MARKET HOURS
        # ------------------------------------------------------------
        minute_of_day = (
            df["datetime"].dt.hour * 60
            + df["datetime"].dt.minute
        )

        df = df[
            (minute_of_day >= MARKET_START) &
            (minute_of_day <= MARKET_END)
        ].copy()

        if df.empty:
            return 0

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
            float(last["Call IV"]),
            float(current_group["Call OI"].mean()),
            float(last["Call Change OI"]),
            float(call_bid_ask),
            float(call_buy),
            float(call_sell),
            float(call_buy - call_sell),
            float(last["OTM Put Strike"]),
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
        # ------------------------------------------------------------
        candle_title = f"{interval_minutes}-Minute Candle Data"

        try:
            current_title = candle_sheet.range("A1").value
        except Exception:
            current_title = None

        if str(current_title or "").strip() != candle_title:
            candle_sheet.range("A1").value = candle_title

        candle_headers = [
"Time", "Close",
            "Call Strike", "Call IV", "Call OI", "Call Change OI", "Call Bid-Ask Avg",
            "Call Total Buy", "Call Total Sell", "Call Buy-Sell Diff",
            "Put Strike", "Put IV", "Put OI", "Put Change OI", "Put Bid-Ask Avg",
            "Put Total Buy", "Put Total Sell", "Put Buy-Sell Diff",
            "PCR OI", "PCR Change OI",
            "Ticks"
        ]

        try:
            current_headers = candle_sheet.range("A2:U2").value
            if current_headers and isinstance(current_headers, list):
                if len(current_headers) == 1 and isinstance(
                    current_headers[0], list
                ):
                    current_headers = current_headers[0]
        except Exception:
            current_headers = None

        if current_headers != candle_headers:
            candle_sheet.range("A2:U2").value = [candle_headers]

        # ------------------------------------------------------------
        # FIND ONLY THE LAST CANDLE ROW
        #
        # We intentionally DO NOT read rows 3:N.
        # Only the latest row matters.
        # ------------------------------------------------------------
        try:
            candle_last_row = get_last_actual_row(candle_sheet, "A", 2)
        except Exception:
            candle_last_row = 2

        if candle_last_row < 3:
            # No candle exists yet -> add the first one.
            target_row = 3

            candle_sheet.range(
                f"A{target_row}:U{target_row}"
            ).value = [new_values]

            candle_sheet.range(
                f"B{target_row}:S{target_row}"
            ).number_format = "#,##0.00"

            print(
                f"🆕 Candle created: {candle_time}"
            )

            return 1

        # ------------------------------------------------------------
        # READ ONLY THE LAST ROW
        # ------------------------------------------------------------
        last_row_values = candle_sheet.range(
            f"A{candle_last_row}:U{candle_last_row}"
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
                f"A{candle_last_row}:U{candle_last_row}"
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
            f"A{target_row}:U{target_row}"
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
    """
    IMPORTANT:
    Older versions stored Call Change OI in Q and Put Change OI in R.

    Old layout:
    A Time
    B Close
    C Call Strike
    D Call IV
    E Call OI
    F Call Bid-Ask Avg
    G Call Total Buy
    H Call Total Sell
    I Call Buy-Sell Diff
    J Put Strike
    K Put IV
    L Put OI
    M Put Bid-Ask Avg
    N Put Total Buy
    O Put Total Sell
    P Put Buy-Sell Diff
    Q Call Change OI
    R Put Change OI
    S PCR OI
    T PCR Change OI
    U Ticks

    New layout:
    A Time
    B Close
    C Call Strike
    D Call IV
    E Call OI
    F Call Change OI
    G Call Bid-Ask Avg
    H Call Total Buy
    I Call Total Sell
    J Call Buy-Sell Diff
    K Put Strike
    L Put IV
    M Put OI
    N Put Change OI
    O Put Bid-Ask Avg
    P Put Total Buy
    Q Put Total Sell
    R Put Buy-Sell Diff
    S PCR OI
    T PCR Change OI
    U Ticks

    This physically moves the EXISTING rows. Changing only the header
    does not fix old data, which was the source of the user's problem.
    """
    headers = [
        "Time", "Close",
        "Call Strike", "Call IV", "Call OI", "Call Change OI",
        "Call Bid-Ask Avg", "Call Total Buy", "Call Total Sell",
        "Call Buy-Sell Diff",
        "Put Strike", "Put IV", "Put OI", "Put Change OI",
        "Put Bid-Ask Avg", "Put Total Buy", "Put Total Sell",
        "Put Buy-Sell Diff", "PCR OI", "PCR Change OI", "Ticks"
    ]

    try:
        last_row = get_last_actual_row(candle_sheet, "A", 2)
        if last_row < 2:
            candle_sheet.range("A2:U2").value = [headers]
            return

        raw = candle_sheet.range(f"A2:U{last_row}").value
        if not raw:
            candle_sheet.range("A2:U2").value = [headers]
            return

        # Flatten one-row return if xlwings gives it that way.
        if isinstance(raw, list) and raw and not isinstance(raw[0], list):
            raw = [raw]

        first = raw[0] if raw else []
        first_norm = [str(x).strip() if x is not None else "" for x in first]

        new_norm = [str(x) for x in headers]

        # Detect the OLD layout by its exact identifying positions.
        old_layout = (
            len(first_norm) >= 21 and
            first_norm[4] == "Call OI" and
            first_norm[5] == "Call Bid-Ask Avg" and
            first_norm[15] == "Put Buy-Sell Diff" and
            first_norm[16] == "Call Change OI" and
            first_norm[17] == "Put Change OI"
        )

        # Detect whether row 2 is already the new header.
        already_new = first_norm[:21] == new_norm

        if old_layout:
            print("🔧 Migrating existing Candles columns: Q→F and R→N...")

            old_data = raw[1:]  # exclude old header

            migrated = []
            for row in old_data:
                row = list(row) + [None] * (21 - len(row))
                migrated.append([
                    row[0],   # A Time
                    row[1],   # B Close
                    row[2],   # C Call Strike
                    row[3],   # D Call IV
                    row[4],   # E Call OI
                    row[16],  # F Call Change OI  <-- OLD Q
                    row[5],   # G Call Bid-Ask Avg
                    row[6],   # H Call Total Buy
                    row[7],   # I Call Total Sell
                    row[8],   # J Call Buy-Sell Diff
                    row[9],   # K Put Strike
                    row[10],  # L Put IV
                    row[11],  # M Put OI
                    row[17],  # N Put Change OI  <-- OLD R
                    row[12],  # O Put Bid-Ask Avg
                    row[13],  # P Put Total Buy
                    row[14],  # Q Put Total Sell
                    row[15],  # R Put Buy-Sell Diff
                    row[18],  # S PCR OI
                    row[19],  # T PCR Change OI
                    row[20],  # U Ticks
                ])

            # Clear and rewrite ONLY the Candles table once, during startup
            # migration. Normal live refreshes never rewrite old rows.
            candle_sheet.range(f"A2:U{last_row}").clear_contents()
            candle_sheet.range("A2:U2").value = [headers]

            if migrated:
                candle_sheet.range(
                    f"A3:U{2 + len(migrated)}"
                ).value = migrated

            print(
                f"✅ Candles migration complete: {len(migrated)} historical rows fixed."
            )
            return

        if not already_new:
            # New sheet or unusual header: set the correct header.
            candle_sheet.range("A2:U2").value = [headers]
            print("✅ Candles header set to corrected order.")
        else:
            print("✅ Candles column order already correct.")

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
    """Create a clean Shoonya OAuth API object."""
    class ShoonyaApiPy(NorenApi):
        def __init__(self):
            super().__init__(
                host="https://api.shoonya.com/NorenWClientAPI/",
                websocket="wss://api.shoonya.com/NorenWSAPI/",
            )
    return ShoonyaApiPy()


def _install_oauth_session(api_obj, uid, token, actid):
    """
    Restore BOTH the normal Noren identity fields and the OAuth HTTP
    Authorization header.  The latter is important for the OAuth
    WebSocket flow.
    """
    api_obj.uid = uid
    api_obj.token = token
    api_obj.actid = actid

    try:
        header = api_obj.injectOAuthHeader(token)
        print("   ✅ OAuth HTTP header/session restored")
        return header
    except Exception as e:
        print(f"   ⚠️ injectOAuthHeader() failed: {e}")
        return None


def _apply_access_token_result(api_obj, result, login_sheet, auth_code):
    """Save and install a successful getAccessToken() result."""
    if not result or not isinstance(result, (tuple, list)) or len(result) < 4:
        return False

    acc_tok, usrid, ref_tok, actid = result

    if not acc_tok or not usrid:
        return False

    _install_oauth_session(api_obj, usrid, acc_tok, actid)

    update_login_data(login_sheet, auth_code, acc_tok, usrid)

    login_sheet.range("C2").value = f"Login OK - {usrid}"
    print(f"✅ Login successful! User: {usrid}")
    print(f"   Auth Code saved: {str(auth_code)[:20]}...")
    print(f"   Token saved: {str(acc_tok)[:20]}...")
    print(f"   UserToken saved: {str(usrid)[:20]}...")
    print("   Token timestamp saved to Excel")
    return True


def shoonya_login(login_sheet):
    """
    Login/re-login policy:

    1. Try SAVED ACCESS TOKEN first.
    2. If saved token cannot be used, try the AUTH CODE already stored
       in Login!  NO browser login at this stage.
    3. If getAccessToken(auth_code, ...) returns None, treat that saved
       AUTH CODE as expired/invalid and ONLY THEN perform a fresh browser
       OAuth login.
    4. If fresh browser OAuth returns an auth code, exchange it for a token.
    5. Never stop merely because the saved auth code returned None.
    """
    global api

    login_data = get_all_login_data(login_sheet)

    user_id = login_data["user_id"]
    client_id = login_data["client_id"] or f"{user_id}_U"
    secret_code = login_data["secret_code"]
    saved_auth_code = login_data["auth_code"]
    saved_token = login_data["token"]
    saved_usertoken = login_data["usertoken"] or user_id

    print("🔐 LOGIN POLICY:")
    print("   1. Saved TOKEN")
    print("   2. Saved AUTH CODE")
    print("   3. Fresh browser OAuth ONLY if saved AUTH CODE returns None")

    if not user_id or not secret_code:
        login_sheet.range("C2").value = "Missing User ID / Secret Code"
        print("❌ Missing User ID or Secret Code in Login sheet.")
        return False

    # ==============================================================
    # STEP 1 — SAVED TOKEN FIRST
    # ==============================================================
    if saved_token and saved_usertoken:
        print("--------------------------------------------------")
        print("🔐 STEP 1: Trying SAVED ACCESS TOKEN")
        print(f"   User ID : {user_id}")
        print(f"   Token   : {str(saved_token)[:20]}...")
        print(f"   ACTID   : {user_id}")

        try:
            api = _make_api()
            _install_oauth_session(api, saved_usertoken, saved_token, user_id)

            # Small HTTP test.  This is NOT used to create a new token.
            # It only tells us whether the saved session can currently
            # talk to the broker.
            test = api.get_quotes(SPOT_EXCHANGE, str(NIFTY_SPOT_TOKEN))

            if test and test.get("lp") not in (None, "", "0", 0):
                login_sheet.range("C2").value = f"Saved token reused - {user_id}"
                print(f"✅ REUSING SAVED TOKEN for user: {user_id}")
                print(f"   NIFTY spot response: {test.get('lp')}")
                print("   🔐 No OAuth browser login")
                print("   🔐 No new auth code")
                return True

            print("❌ Saved token did not return usable market data.")
        except Exception as e:
            print(f"❌ Saved token/session failed: {e}")

    # ==============================================================
    # STEP 2 — SAVED AUTH CODE
    # ==============================================================
    if saved_auth_code:
        print("--------------------------------------------------")
        print("🔐 STEP 2: Trying AUTH CODE already stored in Login sheet")
        print(f"   Auth Code: {str(saved_auth_code)[:25]}...")
        print("   🚫 No browser login yet")
        print("   🔄 Calling EXACT getAccessToken(auth_code, secret, client_id, uid)...")

        try:
            api = _make_api()

            result = api.getAccessToken(
                saved_auth_code,
                secret_code,
                client_id,
                user_id
            )

            print("getAccessToken() raw result:", result)

            if _apply_access_token_result(
                api, result, login_sheet, saved_auth_code
            ):
                print("✅ Saved AUTH CODE produced a fresh access token.")
                return True

            # IMPORTANT:
            # For this program, None/invalid tuple means the saved AUTH
            # CODE is no longer usable.  NOW, and only now, do browser OAuth.
            print("❌ getAccessToken() returned None/invalid result.")
            print("   ➜ Treating saved AUTH CODE as expired/invalid.")
        except Exception as e:
            print(f"❌ Saved AUTH CODE exchange raised an error: {e}")
            print("   ➜ Treating saved AUTH CODE as unusable.")

    else:
        print("--------------------------------------------------")
        print("⚠️ No saved AUTH CODE in Login sheet.")

    # ==============================================================
    # STEP 3 — FRESH BROWSER OAUTH
    # ==============================================================
    print("--------------------------------------------------")
    print("🔄 STEP 3: FRESH OAuth browser login")
    print("   This happens ONLY after saved TOKEN + saved AUTH CODE fail.")

    if not login_data["password"] or not login_data["totp_secret"]:
        login_sheet.range("C2").value = (
            "Saved token/auth code failed; Password/TOTP required for fresh OAuth"
        )
        print("❌ Password or TOTP Secret missing; cannot perform fresh OAuth.")
        return False

    try:
        api = _make_api()

        login_sheet.range("C2").value = "Opening browser for fresh OAuth..."
        print(f"   User ID  : {user_id}")
        print(f"   Client ID: {client_id}")
        print("   Opening browser for OAuth login...")

        fresh_auth_code = get_auth_code_via_selenium(
            client_id,
            user_id,
            login_data["password"],
            login_data["totp_secret"]
        )

        if not fresh_auth_code:
            print("❌ Browser OAuth did not return an AUTH CODE.")
            login_sheet.range("C2").value = "Fresh OAuth auth code not received"
            return False

        print(f"   OAuth authorization code received: {str(fresh_auth_code)[:25]}...")
        print("   Getting access token...")

        fresh_result = api.getAccessToken(
            fresh_auth_code,
            secret_code,
            client_id,
            user_id
        )

        print("Fresh getAccessToken() raw result:", fresh_result)

        if not _apply_access_token_result(
            api, fresh_result, login_sheet, fresh_auth_code
        ):
            print("❌ Fresh AUTH CODE also failed to produce a token.")
            login_sheet.range("C2").value = "Fresh OAuth token exchange failed"
            return False

        print("✅ Fresh OAuth login completed.")
        return True

    except Exception as e:
        print(f"❌ Fresh OAuth login error: {e}")
        import traceback
        traceback.print_exc()
        login_sheet.range("C2").value = f"OAuth error: {str(e)[:120]}"
        return False


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
    
    # Extract Feed Time (ft) - format as HH:MM:SS
    if 'ft' in msg:
        feed_time = format_timestamp(msg['ft'])
    elif 'exch_tm' in msg:
        feed_time = format_timestamp(msg['exch_tm'])
    elif 'ts' in msg:
        feed_time = format_timestamp(msg['ts'])
    
    # Extract Last Traded Time (ltt) - format as HH:MM:SS
    if 'ltt' in msg:
        last_traded_time = format_timestamp(msg['ltt'])
    
    # If we have feed time but not last traded time, use feed time
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


# ============================================================
# ATM HIGHLIGHT - GREEN BACKGROUND
# ============================================================
def apply_atm_highlight(oc_sheet, atm_strike):
    """Highlight the ATM row in the OptionChain data area.

    OptionChain headers are on row 10 and data starts on row 11.
    The highlight extends across all 27 displayed columns (A:AA).
    """
    try:
        # Clear only the displayed OptionChain data area.
        oc_sheet.range("A11:AA1000").color = None

        data = oc_sheet.range("A11:AA1000").value
        if not data or not isinstance(data, list):
            return

        for i, row in enumerate(data):
            if not row:
                continue

            # OptionChain columns:
            # F = Call Strike, P = Put Strike (1-based Excel columns)
            call_strike = convert_to_float(row[5]) if len(row) > 5 else None
            put_strike = convert_to_float(row[15]) if len(row) > 15 else None

            if ((call_strike is not None and abs(call_strike - atm_strike) < 0.01) or
                (put_strike is not None and abs(put_strike - atm_strike) < 0.01)):
                row_num = i + 11
                oc_sheet.range(f"A{row_num}:AA{row_num}").color = (0, 255, 0)
                print(f"✅ ATM Highlight applied to row {row_num} (ATM Strike: {atm_strike})")
                break

    except Exception as e:
        print(f"⚠️ ATM highlight error: {e}")

# ----------------------------------------------------------------------
# Function to get first OTM Call and Put strikes
# ----------------------------------------------------------------------
def get_nse_strike_classification(df_full, spot_ltp):
    """
    NSE-style strike classification using the ACTUAL strikes present in
    the NFO instrument list.

    Definitions:
      CE:
        ITM = strike < spot
        ATM = nearest actual strike to spot
        OTM = strike > spot

      PE:
        ITM = strike > spot
        ATM = nearest actual strike to spot
        OTM = strike < spot

    The function NEVER invents a strike.  It selects from the actual
    available NIFTY option strikes.
    """
    strikes = sorted(
        pd.to_numeric(df_full["strike"], errors="coerce")
        .dropna()
        .unique()
        .tolist()
    )

    if not strikes:
        return {
            "atm": None,
            "ce_itm": None,
            "ce_otm": None,
            "pe_itm": None,
            "pe_otm": None,
        }

    spot = float(spot_ltp)

    # NSE option-chain ATM = nearest available strike.
    atm = min(strikes, key=lambda x: (abs(x - spot), x))

    below = [s for s in strikes if s < spot]
    above = [s for s in strikes if s > spot]

    # Call:
    #   nearest strike below spot = first ITM CE
    #   nearest strike above spot = first OTM CE
    ce_itm = max(below) if below else None
    ce_otm = min(above) if above else None

    # Put:
    #   nearest strike above spot = first ITM PE
    #   nearest strike below spot = first OTM PE
    pe_itm = min(above) if above else None
    pe_otm = max(below) if below else None

    return {
        "atm": atm,
        "ce_itm": ce_itm,
        "ce_otm": ce_otm,
        "pe_itm": pe_itm,
        "pe_otm": pe_otm,
    }


def get_first_otm_strikes(df_full, spot_ltp, atm_strike=None, strike_step=50):
    """
    Compatibility wrapper used by the existing option-chain code.

    IMPORTANT:
    It uses the actual NSE/NFO strikes in df_full.
    It does NOT calculate OTM as ATM +/- a hard-coded 50 if an actual
    strike is available.
    """
    cls = get_nse_strike_classification(df_full, spot_ltp)

    otm_call_strike = cls["ce_otm"]
    otm_put_strike = cls["pe_otm"]

    # Only as a final defensive fallback.
    if otm_call_strike is None and atm_strike is not None:
        otm_call_strike = atm_strike + strike_step

    if otm_put_strike is None and atm_strike is not None:
        otm_put_strike = atm_strike - strike_step

    return otm_call_strike, otm_put_strike


# ----------------------------------------------------------------------
# Function to store data in history sheet - APPEND new data WITH IV
# ----------------------------------------------------------------------
def store_tick_data(history_sheet, df_full, spot_ltp, atm_strike, otm_call_strike, otm_put_strike, expiry_input, future_ltp, atm_call_price, atm_put_price):
    """
    Store a market snapshot in Tick_History.

    IMPORTANT:
    1. Never write after 15:30 IST.
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
        market_end_minutes = MARKET_END

        # Allow the final 15:30 snapshot, but NOTHING after 15:30:00.
        if market_minutes > market_end_minutes:
            return

        if market_minutes == market_end_minutes and now_ist.second > 0:
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
        if not last_traded_time:
            last_traded_time = feed_time

        # Use current exchange/feed timestamp when available; otherwise
        # keep the existing value maintained by the main loop.
        headers = [
            "Feed Time", "Request Time", "Last Traded Time",
            "Spot", "ATM Strike",
            "OTM Call Strike", "Call IV", "Call OI", "Call Change OI",
            "Call Total Buy", "Call Total Sell", "Call Buy-Sell Diff",
            "Call Bid", "Call Ask", "Call Bid-Ask Diff",
            "OTM Put Strike", "Put IV", "Put OI", "Put Change OI",
            "Put Total Buy", "Put Total Sell", "Put Buy-Sell Diff",
            "Put Bid", "Put Ask", "Put Bid-Ask Diff",
            "PCR OI", "PCR Change OI"
        ]

        # Ensure headers exist.
        try:
            if history_sheet.range("A1").value != "Feed Time":
                history_sheet.range("A1:AA1").value = headers
                history_sheet.range("B:B").number_format = "yyyy-mm-dd hh:mm:ss"
        except Exception:
            history_sheet.range("A1:Z1").value = headers

        row_data = [
            feed_time,
            now_dt,
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

        history_sheet.range(f"A{next_row}:AA{next_row}").value = row_data
        history_sheet.range(f"B{next_row}").number_format = "yyyy-mm-dd hh:mm:ss"
        history_sheet.range(f"D{next_row}:AA{next_row}").number_format = "#,##0.00"

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
def run_option_chain(wb, oc_sheet, history_sheet, candle_sheet):
    global tick_counter, feed_time, request_time, last_traded_time, last_aggregation_time, AGGREGATION_INTERVAL
        
    fut_token = get_index_future_token()
    if fut_token is not None:
        subscribe_token(EXCHANGE, fut_token)
    subscribe_token(SPOT_EXCHANGE, NIFTY_SPOT_TOKEN)

    build_option_chain_template()
    dump_available_expiries(oc_sheet)

    pre_expiry = None
    pre_no_of_strike = None
    refresh_rate = 0.5  #  1
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
                    candle_sheet.range("A1").value = f"{AGGREGATION_INTERVAL}-Minute Candle Data"
            
            expiry_str = oc_sheet.range("B1").value
            
            try:
                expiry_input = parse_date(expiry_str)
            except ValueError as e:
                oc_sheet.range("C1").value = f"Invalid date format: {expiry_str}"
                print(f"❌ Date parsing error: {e}")
                time.sleep(2)
                continue

            no_of_strike = int(oc_sheet.range("B2").value or NUMBER_OF_STRIKES)
            refresh_rate = int(oc_sheet.range("B3").value or 1)

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

            spot_ltp = convert_to_float(api.get_quotes(SPOT_EXCHANGE, str(NIFTY_SPOT_TOKEN)).get("lp"))
            future_ltp = convert_to_float(api.get_quotes(EXCHANGE, str(fut_token)).get("lp")) if fut_token else spot_ltp

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

            # ============================================================
            # NSE-STYLE ACTUAL STRIKE CLASSIFICATION
            # ============================================================
            strike_class = get_nse_strike_classification(df_full, spot_ltp)

            atm_strike = strike_class["atm"]
            ce_itm_strike = strike_class["ce_itm"]
            ce_otm_strike = strike_class["ce_otm"]
            pe_itm_strike = strike_class["pe_itm"]
            pe_otm_strike = strike_class["pe_otm"]

            if atm_strike is None:
                print("⚠️ No actual NSE strike available for current spot.")
                time.sleep(refresh_rate)
                continue

            # Locate the actual ATM row.
            atm_matches = df_full[df_full["strike"] == atm_strike]
            if atm_matches.empty:
                print("⚠️ ATM strike row not found.")
                time.sleep(refresh_rate)
                continue

            atm_idx = atm_matches.index[0]

            # Get ATM data
            atm_ce_price = df_full.loc[atm_idx, "CE_lp"]
            atm_pe_price = df_full.loc[atm_idx, "PE_lp"]

            # IMPORTANT: Refresh IV calculator on EVERY cycle.
            # Spot, future, ATM prices and time change continuously. Keeping
            # the calculator only for the whole expiry makes IV stale.
            print("🔄 Updating IV calculator with current market snapshot...")
            if not init_iv_calculator(
                spot_ltp, future_ltp, atm_strike,
                atm_ce_price, atm_pe_price, expiry_input
            ):
                print("⚠️ IV calculator update failed; IV will be 0 for this refresh.")
            pre_expiry = expiry_input

            # First OTM strikes from ACTUAL NSE/NFO strike list.
            otm_call_strike = ce_otm_strike
            otm_put_strike = pe_otm_strike

            # Defensive fallback only if the instrument list has no
            # strike on one side.
            if otm_call_strike is None:
                otm_call_strike = get_first_otm_strikes(
                    df_full, spot_ltp, atm_strike, STRIKE_STEP
                )[0]

            if otm_put_strike is None:
                otm_put_strike = get_first_otm_strikes(
                    df_full, spot_ltp, atm_strike, STRIKE_STEP
                )[1]

            print(
                f"🎯 NSE strikes | Spot={spot_ltp:.2f} | "
                f"ATM={atm_strike} | "
                f"CE ITM={ce_itm_strike} | CE OTM={ce_otm_strike} | "
                f"PE ITM={pe_itm_strike} | PE OTM={pe_otm_strike}"
            )

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
                pre_expiry, pre_no_of_strike = expiry_input, no_of_strike

            oc_sheet.range("A10").options(index=False, header=True).value = df_final
            
            apply_atm_highlight(oc_sheet, atm_strike)
            
            # Store tick data using FULL dataframe with IV calculation
            store_tick_data(history_sheet, df_full, spot_ltp, atm_strike, otm_call_strike, otm_put_strike, expiry_input, future_ltp, atm_ce_price, atm_pe_price)
            # ============================================================
            # REAL-TIME CANDLE UPDATE - NO STARTUP-TIMER DELAY
            # ============================================================
            # IMPORTANT:
            # The old code waited AGGREGATION_INTERVAL*60 seconds from
            # script startup. If the script started at 10:00:30, the first
            # aggregation happened at 10:01:30, creating a ~30 second delay.
            #
            # Now aggregation is checked EVERY refresh cycle.
            # aggregate_candles() itself writes only the current candle row
            # when its values changed, so old candle rows are NOT rewritten.
            # ============================================================
            if tick_counter >= 2:
                aggregate_candles(
                    history_sheet,
                    candle_sheet,
                    AGGREGATION_INTERVAL
                )
            
            oc_sheet.range("C1").value = (
                f"{SYMBOL} Spot={spot_ltp:.1f}  ATM={atm_strike}  "
                f"CE ITM={ce_itm_strike}  CE OTM={otm_call_strike}  "
                f"PE ITM={pe_itm_strike}  PE OTM={otm_put_strike}  "
                f"Ticks={tick_counter}  Aggregation={AGGREGATION_INTERVAL}min  "
                f"Feed={feed_time}  Req={request_time}  LTT={last_traded_time}  "
                f"IV=LIVE DTE={current_iv_dte_days} T={current_iv_T:.8f} r=10%"
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
    print("-" * 50)
    
    wb, login_sheet, oc_sheet, history_sheet, candle_sheet = get_or_create_workbook()

    if not shoonya_login(login_sheet):
        print("❌ Login failed. Check the Login sheet.")
        sys.exit(1)

    load_instruments()

    # --------------------------------------------------------------
    # Start WebSocket with the CURRENT OAuth session.
    # Authentication recovery has already happened inside
    # shoonya_login().
    # --------------------------------------------------------------
    feed_opened = False
    print("🔌 Starting WebSocket with current OAuth session...")
    print(f"   WebSocket UID : {getattr(api, 'uid', '')}")
    print(f"   WebSocket ACTID: {getattr(api, 'actid', '')}")
    print(f"   Token loaded  : {str(getattr(api, 'token', ''))[:20]}...")

    api.start_websocket(
        order_update_callback=event_handler_order_update,
        subscribe_callback=event_handler_quote_update,
        socket_open_callback=open_callback,
        socket_close_callback=event_handler_socket_closed,
    )

    # IMPORTANT:
    # A WebSocket timeout is NOT automatically treated as a market-data
    # timeout.  Noren may reconnect asynchronously.  Wait patiently.
    websocket_wait = 0
    while not feed_opened and websocket_wait < 60:
        time.sleep(0.5)
        websocket_wait += 0.5

    if feed_opened:
        print("✅ WebSocket connected. Enter expiry/strike count in the OptionChain sheet.")
    else:
        print("⚠️ WebSocket has not opened after 60 seconds.")
        print("   Continuing; Noren's websocket may reconnect asynchronously.")
        print("   No new OAuth code will be generated merely because no tick arrived.")

    print("-" * 50)

    # NEVER stop here merely because the first websocket/tick was delayed.
    # The option-chain loop itself continues and handles market-data gaps.
    run_option_chain(wb, oc_sheet, history_sheet, candle_sheet)
