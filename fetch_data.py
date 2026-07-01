import requests
import datetime
import json
import os
import csv
import traceback
import sys
import glob
import re
import time

# Define both URLs
URLS = [
    "https://www.nseindia.com/api/market-data-pre-open?key=NIFTY",
    "https://www.nseindia.com/api/market-data-pre-open?key=FO"
]

HOME_URL = "https://www.nseindia.com/report-detail/fo_eq_security"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/report-detail/fo_eq_security",
}

def get_filename_from_url(url):
    """
    Get filename based on URL parameters
    URL: https://www.nseindia.com/api/market-data-pre-open?key=NIFTY
    Returns: pre-open-NIFTY-DD-MMM-YYYY.csv
    """
    # Extract key from URL
    key_match = re.search(r'key=([^&]+)', url)
    key = key_match.group(1) if key_match else 'UNKNOWN'
    
    today = datetime.date.today()
    # Format: DD-MMM-YYYY (e.g., 01-JUL-2026)
    date_str = today.strftime("%d-%b-%Y").upper()
    
    return f"pre-open-{key}-{date_str}.csv"

def get_next_filename(base_name):
    """
    Get the next available filename with number suffix.
    Example: if base_name.csv exists, returns base_name (1).csv
    """
    # Create data folder if it doesn't exist
    os.makedirs("data", exist_ok=True)
    
    # First, check if the base file exists
    base_file = f"data/{base_name}"
    if not os.path.exists(base_file):
        return base_file
    
    # If it exists, find the next number
    base_without_ext = os.path.splitext(base_name)[0]
    ext = os.path.splitext(base_name)[1]
    pattern = f"data/{base_without_ext} (*){ext}"
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
    
    return f"data/{base_without_ext} ({next_num}){ext}"

def fetch_json_data(url):
    """Fetch JSON data from NSE API for a specific URL"""
    print(f"🔍 Fetching data from: {url}")
    
    session = requests.Session()
    session.headers.update(HEADERS)
    
    # First visit homepage to get cookies
    try:
        home_response = session.get(HOME_URL, timeout=10)
        print(f"   Homepage status: {home_response.status_code}")
    except Exception as e:
        print(f"❌ Homepage error: {e}")
        raise
    
    # Fetch API data
    print(f"📊 Fetching pre-open data for {url.split('key=')[1]}...")
    try:
        response = session.get(url, timeout=10)
        print(f"   Response status: {response.status_code}")
        response.raise_for_status()
    except Exception as e:
        print(f"❌ API error: {e}")
        print(f"   Response: {response.text[:200] if 'response' in locals() else 'No response'}")
        raise
    
    print("✅ Data fetched successfully")
    return response.json()

def parse_and_save(json_data, url):
    """Parse JSON and save as CSV with filename from URL key"""
    print("🔍 Debug: Starting parse_and_save...")
    
    # Create data folder if it doesn't exist
    os.makedirs("data", exist_ok=True)
    
    # Get the filename based on URL key
    url_filename = get_filename_from_url(url)
    print(f"📁 Filename from URL: {url_filename}")
    
    # Get the next available filename (handles duplicates)
    filename = get_next_filename(url_filename)
    print(f"📁 Saving to: {filename}")
    
    records = []
    data_list = json_data.get('data', [])
    print(f"📊 Found {len(data_list)} items in data")
    
    if not data_list:
        print("⚠️ No data received from API")
        return []
    
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

def download_all():
    """Download data for all URLs"""
    print("=" * 60)
    print("🚀 Starting NSE data fetch for multiple URLs...")
    print(f"📅 Time: {datetime.datetime.now()}")
    print(f"📁 Current directory: {os.getcwd()}")
    print("=" * 60)
    
    all_results = []
    total_records = 0
    
    for i, url in enumerate(URLS, 1):
        print(f"\n{'='*60}")
        print(f"📌 Processing URL {i}/{len(URLS)}: {url}")
        print('='*60)
        
        try:
            # Add a small delay between requests to avoid rate limiting
            if i > 1:
                print("⏳ Waiting 2 seconds before next request...")
                time.sleep(2)
            
            # Fetch data
            json_data = fetch_json_data(url)
            
            # Parse and save
            records = parse_and_save(json_data, url)
            
            if records:
                total_records += len(records)
                all_results.append({
                    'url': url,
                    'key': url.split('key=')[1],
                    'records': len(records),
                    'success': True
                })
                print(f"✅ Successfully processed {url.split('key=')[1]}")
            else:
                all_results.append({
                    'url': url,
                    'key': url.split('key=')[1],
                    'records': 0,
                    'success': False
                })
                print(f"❌ Failed to process {url.split('key=')[1]}")
                
        except Exception as e:
            print(f"\n❌ ERROR processing {url}: {e}")
            print("\n📋 Full traceback:")
            traceback.print_exc()
            all_results.append({
                'url': url,
                'key': url.split('key=')[1],
                'records': 0,
                'success': False,
                'error': str(e)
            })
    
    # Summary
    print("\n" + "=" * 60)
    print("📊 SUMMARY REPORT")
    print("=" * 60)
    successful = [r for r in all_results if r['success']]
    failed = [r for r in all_results if not r['success']]
    
    print(f"✅ Successful downloads: {len(successful)}")
    print(f"❌ Failed downloads: {len(failed)}")
    print(f"📊 Total records saved: {total_records}")
    
    if successful:
        print("\n📁 Downloaded files:")
        for result in successful:
            filename = get_filename_from_url(result['url'])
            print(f"   - {filename} ({result['records']} records)")
    
    if failed:
        print("\n❌ Failed URLs:")
        for result in failed:
            error_msg = result.get('error', 'Unknown error')
            print(f"   - {result['url']} ({error_msg})")
    
    print("=" * 60)
    return all_results

if __name__ == "__main__":
    try:
        download_all()
    except Exception as e:
        print(f"\n❌ FATAL ERROR: {e}")
        print("\n📋 Full traceback:")
        traceback.print_exc()
