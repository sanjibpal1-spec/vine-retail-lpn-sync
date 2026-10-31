import os
import time
import json
from datetime import datetime, timedelta
import pandas as pd
import requests
from playwright.sync_api import sync_playwright

WEBHOOK_URL = os.environ["GSHEET_WEBHOOK_URL"]

def upload_to_gsheet(df):
    # Convert DataFrame to a list of lists (Headers + Data)
    data_payload = [df.columns.values.tolist()] + df.fillna("").values.tolist()
    
    response = requests.post(
        WEBHOOK_URL,
        data=json.dumps({"data": data_payload}),
        headers={"Content-Type": "application/json"}
    )
    print("Sheet upload response:", response.text)

def get_target_dates():
    now = datetime.now()
    current_hour = now.hour
    today_str = now.strftime("%d/%m/%Y")
    yesterday_str = (now - timedelta(days=1)).strftime("%d/%m/%Y")
    
    if 0 <= current_hour < 20:
        return [yesterday_str, today_str]
    else:
        return [today_str]

def download_vine_report(page, download_dir, target_date):
    page.goto("https://multistorewh.vineretail.com/eRetailWeb/eRetailLogin")
    page.fill("#username", os.environ["VINE_USER"])
    page.fill("#password", os.environ["VINE_PASS"])
    page.click("button:has-text('Login')")
    
    page.click("text=Reports")
    page.click("text=MIS Report 2.0")
    
    page.select_option("select[name='reportGroup']", label="Sales")
    page.select_option("select[name='reportName']", label="PICK IPP Report")
    
    page.fill("input[name='dateRange']", f"{target_date} - {target_date}")
    
    page.click(".select2-selection")
    page.click("li:has-text('JCT-CBE IM1 Temp')")
    
    page.click("button:has-text('Export')")
    
    print(f"Waiting 3 minutes for report ({target_date})...")
    time.sleep(180)
    
    page.click("img[title='Refresh'], .blue-refresh-icon")
    time.sleep(5)
    
    with page.expect_download() as download_info:
        page.click("img[src*='download'], .download-icon")
    download = download_info.value
    filepath = os.path.join(download_dir, f"report_{target_date.replace('/', '_')}.xlsx")
    download.save_as(filepath)
    return filepath

def main():
    target_dates = get_target_dates()
    all_filtered_dfs = []
    download_dir = os.path.abspath("./downloads")
    os.makedirs(download_dir, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        for d in target_dates:
            file_path = download_vine_report(page, download_dir, d)
            df = pd.read_excel(file_path, header=10)
            
            if 'Order_Type' in df.columns:
                filtered_df = df[df['Order_Type'].astype(str).str.strip().str.upper() == 'PREPAID']
            else:
                filtered_df = df[df.iloc[:, 6].astype(str).str.strip().str.upper() == 'PREPAID']
                
            all_filtered_dfs.append(filtered_df)
        
        browser.close()

    final_df = pd.concat(all_filtered_dfs, ignore_index=True)
    upload_to_gsheet(final_df)

if __name__ == "__main__":
    main()
