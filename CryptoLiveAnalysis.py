import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from time import sleep
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

# Function to fetch cryptocurrency data
def fetch_crypto_data():
    """
    Fetches live cryptocurrency data for the top 50 coins by market capitalization.
    Returns:
        pd.DataFrame: DataFrame containing cryptocurrency data.
    """
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "price_change_percentage": "24h"
    }
    try:
        response = requests.get(url, params=params)
        if response.status_code == 200:
            data = response.json()
            crypto_data = [
                {
                    "Name": coin['name'],
                    "Symbol": coin['symbol'],
                    "Current Price (USD)": coin['current_price'],
                    "Market Cap": coin['market_cap'],
                    "24h Trading Volume": coin['total_volume'],
                    "24h % Change": coin['price_change_percentage_24h']
                }
                for coin in data
            ]
            return pd.DataFrame(crypto_data)
        else:
            logging.error(f"Error fetching data: {response.status_code}")
            return pd.DataFrame()
    except Exception as e:
        logging.error(f"An exception occurred: {e}")
        return pd.DataFrame()

# Save live data to Excel
def save_to_excel(df, filename="crypto_data.xlsx"):
    """
    Saves DataFrame data to an Excel file.
    Args:
        df (pd.DataFrame): DataFrame containing the data to save.
        filename (str): Name of the Excel file.
    """
    try:
        with pd.ExcelWriter(filename, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, index=False, sheet_name="Crypto Data")
        logging.info(f"Data saved to {filename}")
    except Exception as e:
        logging.error(f"Error saving to Excel: {e}")

# Perform basic analysis
def analyze_data(df):
    """
    Performs basic analysis on cryptocurrency data.
    Args:
        df (pd.DataFrame): DataFrame containing cryptocurrency data.
    Returns:
        dict: Analysis results.
    """
    if df.empty:
        return {"Error": "No data available for analysis."}

    top_5_by_market_cap = df.nlargest(5, "Market Cap")
    avg_price = df["Current Price (USD)"].mean()
    highest_change = df["24h % Change"].max()
    lowest_change = df["24h % Change"].min()

    analysis = {
        "Top 5 Cryptocurrencies by Market Cap": top_5_by_market_cap.to_dict('records'),
        "Average Price": avg_price,
        "Highest 24h Change (%)": highest_change,
        "Lowest 24h Change (%)": lowest_change
    }
    return analysis

# Main function to run continuously
def main(update_frequency=300):
    """
    Fetches, saves, and analyzes cryptocurrency data continuously.
    Args:
        update_frequency (int): Frequency of updates in seconds.
    """
    while True:
        logging.info("Fetching live cryptocurrency data...")
        df = fetch_crypto_data()
        if not df.empty:
            save_to_excel(df)
            analysis = analyze_data(df)
            logging.info("Analysis Report:")
            logging.info(analysis)
        logging.info(f"Waiting for the next update in {update_frequency} seconds...\n")
        sleep(update_frequency)

if __name__ == "__main__":
    main()
