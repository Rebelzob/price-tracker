import pandas as pd
import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse
from datetime import datetime
import re
import os
import logging
import matplotlib.pyplot as plt
from openpyxl import Workbook


logging.basicConfig(filename='price_scraper.log', level=logging.INFO, 
                    format='%(asctime)s:%(levelname)s:%(message)s')

def get_amazon_prices(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Connection": "keep-alive",
        "DNT": "1"
    }
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            price = (
                soup.find("span", {"class": "a-price-whole"}) or
                soup.find("div", {"class": "productdetail_product_inclvat"})
            )
            if price:
                raw_price = price.get_text().strip()
                cleaned_price = re.sub(r'[^\d]', '', raw_price)
                return cleaned_price
            else:
                logging.error(f"Price not found for {url}")
                return "Price Not Found"
        else:
            logging.error(f"HTTP error {response.status_code} for {url}")
            return f"HTTP Error: {response.status_code}"
    except requests.exceptions.RequestException as e:
        logging.error(f"Request error for {url}: {e}")
        return f"Request Error: {e}"
    
def scrape_prices():
    df = pd.read_excel("pc_build_components.xlsx")
    current_date = datetime.now().strftime("%Y-%m-%d")
    price_history_file = "price_history.xlsx"

    if not os.path.exists(price_history_file):
        with pd.ExcelWriter(price_history_file, engine='openpyxl') as writer:
            pd.DataFrame({'Initial': [0]}).to_excel(writer, sheet_name='Temp')

    for index, row in df.iterrows():
        url = str(row['URL'])
        if is_valid_url(url):
            price = get_amazon_prices(url)
        else:
            price = "Invalid URL"
        
        price_data = {
            "Date": [current_date],
            "Price": [price]
        }

        item_name = row['Model']
        append_price_history(price_history_file, item_name, price_data)

    with pd.ExcelWriter(price_history_file, mode='a', engine='openpyxl') as writer:
        book = writer.book
        if 'Temp' in book.sheetnames:
            del book['Temp']

    logging.info("Prices updated and saved to price_history.xlsx")

def is_valid_url(url):
    try:
        result = urlparse(url)
        return all([result.scheme, result.netloc])
    except ValueError:
        logging.error(f"Invalid URL format: {url}")
        return False

def append_price_history(price_history_file, item_name, price_data):
    item_name = item_name[:31]
    try:
        item_df = pd.read_excel(price_history_file, sheet_name=item_name, engine='openpyxl')
    except ValueError:  # Sheet does not exist
        item_df = pd.DataFrame()
    except Exception as e:
        logging.error(f"Error reading sheet {item_name} in {price_history_file}: {e}")
        return

    new_data = pd.DataFrame(price_data)
    updated_df = pd.concat([item_df, new_data], ignore_index=True)

    try:
        with pd.ExcelWriter(price_history_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            updated_df.to_excel(writer, sheet_name=item_name, index=False)
    except Exception as e:
        logging.error(f"Error writing to sheet {item_name} in {price_history_file}: {e}")

    generate_graph(updated_df, item_name)

def generate_graph(df, item_name):
    plt.figure(figsize=(10, 5))
    plt.plot(df['Date'], df['Price'], marker='o')
    plt.title(f'Price Trend for {item_name}')
    plt.xlabel('Date')
    plt.ylabel('Price')
    plt.xticks(rotation=45)
    plt.grid(True)
    plt.tight_layout()
    plt.savefig(f'static/graphs/{item_name}_price_trend.png')
    plt.close()

def create_summary(price_history_file):
    try:
        xls = pd.ExcelFile(price_history_file, engine='openpyxl')
        summary = pd.DataFrame()
        for sheet_name in xls.sheet_names:
            if sheet_name == 'Summary':
                continue
            df = pd.read_excel(price_history_file, sheet_name=sheet_name, engine='openpyxl')
            df['Price'] = pd.to_numeric(df['Price'], errors='coerce').fillna(0)
            if summary.empty:
                summary = df[['Date', 'Price']].copy()
            else:
                summary = pd.merge(summary, df[['Date', 'Price']], on='Date', how='outer', suffixes=(None, f'_{sheet_name}'))
                summary['Price'] = summary.filter(like='Price').sum(axis=1)
        summary = summary[['Date', 'Price']]
        summary = summary.groupby('Date', as_index=False).sum()
        with pd.ExcelWriter(price_history_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            summary.to_excel(writer, sheet_name='Summary', index=False)
        generate_graph(summary, 'Summary')
    except Exception as e:
        logging.error(f"Error creating summary: {e}")