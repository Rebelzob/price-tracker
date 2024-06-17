from flask import Flask, render_template, request, redirect, url_for
import os
import pandas as pd
from datetime import datetime
from price_scraper import scrape_prices, create_summary, generate_graph

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/scrape')
def scrape():
    current_date = datetime.now().strftime('%Y-%m-%d')
    price_history_file = "price_history.xlsx"

    if not os.path.exists(price_history_file) or not is_scraped_today(price_history_file, current_date):
        scrape_prices()
        create_summary(price_history_file)
    return redirect(url_for('index'))

@app.route('/graphs')
def graphs():
    graphs_dir = 'static/graphs'
    graphs = [f for f in os.listdir(graphs_dir) if os.path.isfile(os.path.join(graphs_dir, f))]
    return render_template('graphs.html', graphs=graphs)

def is_scraped_today(file_path, date):
    try:
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
            if date in df['Date'].values:
                return True
    except Exception as e:
        print(f"Error checking if scraped today: {e}")
    return False

if __name__ == "__main__":
    app.run(debug=True)
