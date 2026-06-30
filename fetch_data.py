import requests
import datetime
import json
import os
import csv

URL = "https://www.nseindia.com/api/market-data-pre-open?key=FO"
HOME_URL = "https://www.nseindia.com"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/market-data/pre-open-market-cm-and-emerge-market",
}

def fetch_json_data():
    """Fetch JSON data from NSE API"""
    session = requests.Session()
    session.headers.update(HEADERS)
    
    print("🌐 Connecting to NSE...")
    session.get(HOME_URL, timeout=10)
    
    print("📊 Fetching F&O pre-open data...")
    response = session.get(URL, timeout=10)
    response.raise_for_status()
    
    return response.json()

def parse_and_save(json_data):
    """Parse JSON and save as CSV"""
    today = datetime.date.today().isoformat()
    os.makedirs("data", exist_ok=True)
    filename = f"data/preopen_fo_{today}.csv"
    
    records = []
    
    for item in json_data.get('data', []):
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
    with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
        fieldnames = ['SYMBOL', 'PREOPEN', 'FINAL_PRICE', 'FINAL_QUANTITY', 
                     'LAST_UPDATE', 'TOTAL_BUY_QTY', 'TOTAL_SELL_QTY', 'ADV_DECL']
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(records)
    
    print(f"✅ Saved: {filename}")
    print(f"📊 Records: {len(records)}")
    
    return records

if __name__ == "__main__":
    try:
        json_data = fetch_json_data()
        records = parse_and_save(json_data)
        
        if records:
            print("\n📋 Sample Data (first 5 rows):")
            print("-" * 80)
            for i, record in enumerate(records[:5]):
                print(f"{i+1}. {record['SYMBOL']:12} | "
                      f"IEP: {record['PREOPEN']:>8} | "
                      f"Buy: {record['TOTAL_BUY_QTY']:>6} | "
                      f"Sell: {record['TOTAL_SELL_QTY']:>6} | "
                      f"{record['ADV_DECL']}")
    
    except Exception as e:
        print(f"❌ Error: {e}")
