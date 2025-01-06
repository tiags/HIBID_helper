import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl.styles import Alignment, Font
import openpyxl
import re
import os
import time
from concurrent.futures import ThreadPoolExecutor
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry

''' 
This program scrapes auction item data from HIBID online auctions, searches each item on eBay and Yahoo to determine resale value.
'''

# Initialize session and retry strategy
session = requests.Session()
retry_strategy = Retry(
    total=3,  # Retry 3 times
    backoff_factor=1,  # Wait between retries
    status_forcelist=[429, 500, 502, 503, 504],  # Retry on these status codes
)
adapter = HTTPAdapter(max_retries=retry_strategy)
session.mount("https://", adapter)

# Set up paths
folder_path = "HIBID_AUCTIONS"
os.makedirs(folder_path, exist_ok=True)

def scrape_prices(title):
    ebay_price = search_ebay(title)
    yahoo_price = search_yahoo(title)
    return ebay_price, yahoo_price

def search_yahoo(title):
    try:
        search_url = f"https://shopping.yahoo.com/search?p={title.replace(' ', '+')}"
        response = session.get(search_url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        class_pattern = re.compile(r'FluidProductCell__PriceText-sc-fsx0f7-9.*')
        prices = []
        for price_element in soup.find_all(class_=class_pattern):
            price_text = price_element.text
            price_numeric = re.search(r'\d+(\.\d+)?', price_text)
            if price_numeric:
                prices.append(float(price_numeric.group()))
        return round(sum(prices) / len(prices), 2) if prices else None
    except requests.exceptions.RequestException as e:
        print(f"Error searching Yahoo for item: {title}, {e}")
        return None

def search_ebay(title):
    try:
        search_url = f"https://www.ebay.ca/sch/i.html?_nkw={title.replace(' ', '+')}"
        response = requests.get(search_url, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        prices = []
        price_elements = soup.find_all(class_='s-item__price')
        for price_element in price_elements:
            price_text = price_element.text
            price_numeric = re.search(r'\d+(\.\d+)?', price_text)
            if price_numeric:
                price = float(price_numeric.group())
                prices.append(price)
            elif not prices:
                return None
        return round(sum(prices) / len(prices), 2) if prices else None
    except requests.exceptions.RequestException as e:
        print("Error searching eBay for item:", title)
        print("Error:", e)
        return None

def scrape_auction_data(base_url):
    page_number = 1
    items_data = []

    while True:
        url = f"{base_url}?apage={page_number}&s=HOT_RANK"
        print("Requesting URL:", url)
        response = requests.get(url)
        print("Response status code:", response.status_code)
        
        if response.status_code != 200:
            print("Error fetching page:", page_number)
            break
        
        soup = BeautifulSoup(response.text, 'html.parser')
        company_name_element = soup.find('h2', class_='company-name').find('a')
        company_name = company_name_element.get_text().strip()
        
        date_of_auction_element = soup.select_one("p")
        raw_text = date_of_auction_element.get_text().strip()
        date_match = re.search(r'Date\(s\)\s*(\d{1,2}/\d{1,2}/\d{4})\s*-\s*(\d{1,2}/\d{1,2}/\d{4})', raw_text)
        end_date = date_match.group(2)
        end_date = end_date.replace('/', '_') 
        workbook_title = f"{company_name}_{end_date}.xlsx"
        save_path = os.path.join(folder_path, workbook_title)
        
        titles = soup.find_all(class_='lot-title')
        if not titles:
            print("No more items on page", page_number)
            break
        
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = [executor.submit(scrape_prices, title.get_text().strip()) for title in titles]
            for title, future in zip(titles, futures):
                ebay_price, yahoo_price = future.result()
                item_title = title.get_text().strip()
                weight_yahoo, weight_ebay = 0.4, 0.6
                
                # Calculate weighted average
                weighted_average_price = None
                if ebay_price is not None and yahoo_price is not None:
                    weighted_average_price = round(yahoo_price * weight_yahoo + (ebay_price * weight_ebay), 2)
                elif ebay_price is not None:
                    weighted_average_price = ebay_price
                elif yahoo_price is not None:
                    weighted_average_price = yahoo_price
                
                items_data.append([item_title, 
                                ebay_price if ebay_price is not None else None, 
                                yahoo_price if yahoo_price is not None else None, 
                                weighted_average_price if weighted_average_price is not None else None])
                time.sleep(1)
                
        # save workbook every 3 pages
        if page_number % 3 == 0:  
            save_items_to_excel(items_data, save_path)
            
        del soup
        del response
        page_number += 1

    save_items_to_excel(items_data, save_path)
    print(f"Workbook saved at {save_path}")
    
def save_items_to_excel(items_data, save_path):
    # Sorting data by Weighted Average Price
    items_data.sort(key=lambda x: float('-inf') if x[3] is None else -x[3], reverse=True)

    # Create pandas DataFrame
    df = pd.DataFrame(items_data, columns=["Item Title", "Ebay Price", "Yahoo Price", "Weighted Average Price"])
    df.sort_values(by="Weighted Average Price", ascending=False, na_position='last', inplace=True)
    df.to_excel(save_path, index=False)

    # Save to Excel with openpyxl formatting
    workbook = openpyxl.load_workbook(save_path)
    worksheet = workbook.active
    bold_font = Font(bold=True)
    worksheet.append(["Item Title", "Ebay Price", "Yahoo Price", "Weighted Average Price"])
    
    # Center align columns
    center_alignment = Alignment(horizontal='center')
    for col in worksheet.columns:
        for cell in col:
            cell.alignment = center_alignment
    worksheet.column_dimensions['A'].width = 80
    worksheet.column_dimensions['B'].width = 25
    worksheet.column_dimensions['C'].width = 25
    worksheet.column_dimensions['D'].width = 25

    # Save the workbook with formatting
    workbook.save(save_path)

base_url = input("Enter Auction URL: ")
scrape_auction_data(base_url)
