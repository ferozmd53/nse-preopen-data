import requests
import datetime
import json
import os
import csv
import traceback
import sys
import glob

URL = "https://www.nseindia.com/api/market-data-pre-open?key=FO"
HOME_URL = "https://www.nseindia.com/report-detail/fo_eq_security"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/report-detail/fo_eq_security",
}

def get_next_filename(base_name):
    """
    Get the next available filename with number suffix.
    Example: if base_name.csv exists, returns base_name (1).csv
    """
    # First, check if the base file exists
    base_file = f"data/{base_name}.csv"
    if not os.path.exists(base_file):
        return base_file
    
    # If it exists, find the next number
    pattern = f"data/{base_name} (*).csv"
    existing_files = glob.glob(pattern)
    
    # Extract numbers from existing files
    numbers = []
    for f in existing_files:
        # Extract number between ( and )
        try:
            num = int(f.split('(')[1].split(')')[0])
            numbers.append(num)
        except:
            continue
    
    # Find the next number
    next_num = max(numbers) + 1 if numbers else 1
    
    return f"data/{base_name} ({next_num}).csv"

def fetch_json_data():
    """Fetch JSON data from NSE API"""
    print("🔍 Debug: Starting fetch_json_data...")
    print(f"🔍 Python version: {sys.version}")
    
    session = requests.Session()
    session.headers.update(HEADERS)
    
    print("🌐 Connecting to NSE...")
    try:
        home_response = session.get(HOME_URL, timeout=10)
        print(f"   Homepage status: {home_response.status_code}")
        print(f"   Cookies: {len(session.cookies)} cookies")
    except Exception as e:
        print(f"❌ Homepage error: {e}")
        raise
    
    print("📊 Fetching F&O pre-open data...")
    try:
        response = session.get(URL, timeout=10)
        print(f"   Response status: {response.status_code}")
        response.raise_for_status()
    except Exception as e:
        print(f"❌ API error: {e}")
        print(f"   Response: {response.text[:200] if 'response' in locals() else 'No response'}")
        raise
    
    print("✅ Data fetched successfully")
    return response.json()

def parse_and_save(json_data):
    """Parse JSON and save as CSV with number suffix"""
    print("🔍 Debug: Starting parse_and_save...")
    
    # Create data folder if it doesn't exist
    os.makedirs("data", exist_ok=True)
    
    # Base filename (without extension)
    today = datetime.date.today().isoformat()
    base_name = f"preopen_fo_{today}"
    
    # Get the next available filename
    filename = get_next_filename(base_name)
    print(f"📁 Saving to: {filename}")
    
    records = []
    data_list = json_data.get('data', [])
    print(f"📊 Found {len(data_list)} items in data")
    
    for item in data_list:
        metadata = item.get('metadata', {})
        detail = item.get('detail', {})
        preopen = detail.get('preOpenMarket', {})
        
        # Extract IEP price from order book
        preopen_list = preopen.get('preopen', [])
        iep_price = None
        for entry in preopen_list:
            if entry.get('iep') == True:
                iep_price = entry.get('price')
                break
        
        if iep_price is None and preopen_list:
            iep_price = preopen_list[0].get('price')
        
        buy_qty = preopen.get('totalBuyQuantity', 0)
        sell_qty = preopen.get('totalSellQuantity', 0)
        
        # Determine ADV/DECL
        if buy_qty > sell_qty:
            adv_decl = 'BUY'
        elif buy_qty < sell_qty:
            adv_decl = 'SELL'
        else:
            adv_decl = 'NEUTRAL'
        
        record = {
            'SYMBOL': metadata.get('symbol', ''),
            'PREOPEN': iep_price,
            'FINAL_PRICE': preopen.get('finalPrice', ''),
            'FINAL_QUANTITY': preopen.get('finalQuantity', ''),
            'LAST_UPDATE': preopen.get('lastUpdateTime', ''),
            'TOTAL_BUY_QTY': buy_qty,
            'TOTAL_SELL_QTY': sell_qty,
            'ADV_DECL': adv_decl
        }
        records.append(record)
    
    # Write to CSV
    try:
        with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
            fieldnames = ['SYMBOL', 'PREOPEN', 'FINAL_PRICE', 'FINAL_QUANTITY', 
                         'LAST_UPDATE', 'TOTAL_BUY_QTY', 'TOTAL_SELL_QTY', 'ADV_DECL']
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(records)
        
        # Verify file was created
        if os.path.exists(filename):
            file_size = os.path.getsize(filename)
            print(f"✅ File created: {filename} ({file_size} bytes)")
        else:
            print(f"❌ File NOT created at {filename}")
            return []
            
    except Exception as e:
        print(f"❌ Error saving file: {e}")
        raise
    
    print(f"✅ Saved: {filename}")
    print(f"📊 Records: {len(records)}")
    return records

if __name__ == "__main__":
    print("=" * 60)
    print("🚀 Starting NSE data fetch...")
    print(f"📅 Time: {datetime.datetime.now()}")
    print(f"📁 Current directory: {os.getcwd()}")
    print("=" * 60)
    
    try:
        json_data = fetch_json_data()
        records = parse_and_save(json_data)
        
        if records:
            print(f"\n✅ Success! {len(records)} records saved.")
            print("\n📋 Sample Data (first 5 rows):")
            print("-" * 80)
            for i, record in enumerate(records[:5]):
                print(f"{i+1}. {record['SYMBOL']:12} | "
                      f"IEP: {record['PREOPEN']:>8} | "
                      f"Buy: {record['TOTAL_BUY_QTY']:>6} | "
                      f"Sell: {record['TOTAL_SELL_QTY']:>6} | "
                      f"{record['ADV_DECL']}")
        else:
            print("❌ No records saved")
            
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        print("\n📋 Full traceback:")
        traceback.print_exc()
    
    print("=" * 60)
