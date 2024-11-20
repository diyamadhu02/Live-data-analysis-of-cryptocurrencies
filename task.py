import requests
import pandas as pd

def fetch_top_cryptocurrencies():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    
    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        data = response.json()
        
        df = pd.DataFrame(data, columns=[
            'name',                
            'symbol',             
            'current_price',      
            'market_cap',          
            'total_volume',        
            'price_change_percentage_24h'  
        ])
        
        return df
    
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data: {e}")
        return None

def analyze_data(df):
    print("\nTop 50 Cryptocurrencies:")
    print(df)
    
    top_5 = df.nlargest(5, 'market_cap')[['name', 'market_cap']]
    print("\nTop 5 Cryptocurrencies by Market Cap:")
    print(top_5)
    
    avg_price = df['current_price'].mean()
    print(f"\nAverage Price of Top 50 Cryptocurrencies: ${avg_price:.2f}")
    
    highest_change = df.loc[df['price_change_percentage_24h'].idxmax()]
    lowest_change = df.loc[df['price_change_percentage_24h'].idxmin()]
    
    print("\nHighest 24-hour Percentage Price Change:")
    print(f"{highest_change['name']} ({highest_change['symbol']}): {highest_change['price_change_percentage_24h']:.2f}%")
    
    print("\nLowest 24-hour Percentage Price Change:")
    print(f"{lowest_change['name']} ({lowest_change['symbol']}): {lowest_change['price_change_percentage_24h']:.2f}%")
    
    return top_5, avg_price, highest_change, lowest_change

def export_to_excel(df, filename="live_crypto_data.xlsx"):
    df.to_excel(filename, index=False)
    print(f"Live data exported to {filename}")

if __name__ == "__main__":
    df = fetch_top_cryptocurrencies()
    if df is not None:
        analyze_data(df)
        export_to_excel(df)
