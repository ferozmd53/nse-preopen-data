import requests
import datetime
import os

URL = "https://www.nseindia.com/api/market-data-pre-open?key=NIFTY&csv=true&selectValFormat=crores"
HOME_URL = "https://www.nseindia.com"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/market-data/pre-open-market-cm-and-emerge-market",
}

def fetch_nse_csv():
    session = requests.Session()
    session.headers.update(HEADERS)

    # Step 1: Hit homepage first to get cookies (NSE blocks direct API calls)
    session.get(HOME_URL, timeout=10)

    # Step 2: Now call the actual data endpoint
    response = session.get(URL, timeout=10)
    response.raise_for_status()

    return response.text

def save_csv(content):
    today = datetime.date.today().isoformat()
    os.makedirs("data", exist_ok=True)
    filepath = f"data/preopen_nifty_{today}.csv"
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"Saved: {filepath}")

if __name__ == "__main__":
    csv_content = fetch_nse_csv()
    save_csv(csv_content)