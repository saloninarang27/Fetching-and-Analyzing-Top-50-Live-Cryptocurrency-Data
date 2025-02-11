import requests
import pandas as pd
import time
from openpyxl import Workbook

def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json()
    else:
        print("Error fetching data:", response.status_code)
        return []

def analyze_data(data):
    df = pd.DataFrame(data)
    df = df[["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"]]
    
    top_5 = df.nlargest(5, "market_cap")
    avg_price = df["current_price"].mean()
    highest_change = df.loc[df["price_change_percentage_24h"].idxmax()]
    lowest_change = df.loc[df["price_change_percentage_24h"].idxmin()]
    
    return df, top_5, avg_price, highest_change, lowest_change

def write_to_excel(df):
    with pd.ExcelWriter("crypto_data.xlsx", engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Crypto Data", index=False)
    print("Excel file updated.")

def main():
    while True:
        data = fetch_crypto_data()
        if data:
            df, top_5, avg_price, highest_change, lowest_change = analyze_data(data)
            print("Updating Excel with live data...")
            write_to_excel(df)
            time.sleep(300)  # Update every 5 minutes
        else:
            print("Retrying in 5 minutes...")
            time.sleep(300)

if __name__ == "__main__":
    main()
