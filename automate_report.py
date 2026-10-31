import os
import time
import json
from datetime import datetime, timedelta
import pandas as pd
import requests
from playwright.sync_api import sync_playwright

WEBHOOK_URL = os.environ["GSHEET_WEBHOOK_URL"]

def upload_to_gsheet(df):
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
    
    # After 12 AM (00:00 to 19:59): fetch BOTH yesterday and today (2 separate runs)
    if 0 <= current_hour < 20:
        return [yesterday_str, today_str]
    # After 8 PM (20:00 to 23:59): fetch TODAY only
    else:
        return [today_str]

def download_vine_report(page, download_dir, target_date):
    print(f"Navigating to login page for target date: {target_date}")
    page.goto("https://multistorewh.vineretail.com/eRetailWeb/eRetailLogin", wait_until="networkidle")
    
    # 1. Fill login credentials
    page.fill("input[name='username'], #username", os.environ["VINE_USER"])
    page.fill("input[name='password'], #password", os.environ["VINE_PASS"])
    
    # 2. Click Login button
    login_selector = "input[type='submit'], input[value='Login'], button[type='submit'], #loginBtn, .btn-primary"
    page.click(login_selector)
    page.wait_for_load_state("networkidle")
    print("Logged in successfully.")
    
    # 3. Open Sidebar Menu -> Reports -> MIS Report 2.0
    # Hover sidebar reports menu icon (book icon)
    sidebar_reports = page.locator("i.fa-book, a[title='Reports'], .sidebar-menu li:has-text('Reports')").first
    sidebar_reports.hover()
    page.wait_for_selector("text=MIS Report 2.0", state="visible", timeout=10000)
    page.click("text=MIS Report 2.0")
    page.wait_for_load_state("networkidle")
    
    # 4. Select Dropdowns (Report Group & Report Name)
    page.select_option("select[name='reportGroup'], #reportGroup", label="Sales")
    page.select_option("select[name='reportName'], #reportName", label="PICK IPP Report")
    
    # 5. Fill Date Range (1 single day per export)
    date_input = page.locator("input[name='dateRange'], #dateRange, .daterangepicker-input").first
    date_input.fill(f"{target_date} - {target_date}")
    
    # 6. Select Location "JCT-CBE IM1 Temp"
    page.click(".select2-selection, select[name='location']")
    page.wait_for_selector("li:has-text('JCT-CBE IM1 Temp')", state="visible")
    page.click("li:has-text('JCT-CBE IM1 Temp')")
    
    # 7. Click Export (moves to Pending Report tab)
    page.click("button:has-text('Export'), input[value='Export'], .btn-warning")
    page.wait_for_load_state("networkidle")
    
    print(f"Export requested. Waiting 3 minutes for report ({target_date}) to generate...")
    time.sleep(180)
    
    # 8. Refresh Pending Report Page via Top-Right Blue/Green Refresh Icon
    refresh_icon = page.locator(".fa-sync, .fa-refresh, img[title*='Refresh'], #pull_refresh, a.refresh_btn").first
    if refresh_icon.is_visible():
        refresh_icon.click()
    else:
        page.reload()
    time.sleep(5)
    
    # 9. Trigger & Save File Download
    with page.expect_download(timeout=60000) as download_info:
        # Download icon under Download column (cylinder/database icon)
        page.click("img[src*='download'], img[src*='db'], a[title='Download'], .fa-download")
    
    download = download_info.value
    filepath = os.path.join(download_dir, f"report_{target_date.replace('/', '_')}.xlsx")
    download.save_as(filepath)
    print(f"File downloaded successfully to {filepath}")
    return filepath

def main():
    target_dates = get_target_dates()
    print(f"Target execution dates: {target_dates}")
    all_filtered_dfs = []
    download_dir = os.path.abspath("./downloads")
    os.makedirs(download_dir, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        for d in target_dates:
            file_path = download_vine_report(page, download_dir, d)
            
            # Read Excel starting from Row 11 (Header is at row 11 -> header=10)
            df = pd.read_excel(file_path, header=10)
            
            # Filter Column G (Order_Type) for "PREPAID"
            if 'Order_Type' in df.columns:
                filtered_df = df[df['Order_Type'].astype(str).str.strip().str.upper() == 'PREPAID']
            else:
                # Fallback to column index 6 (Column G) if header name has trailing spaces
                filtered_df = df[df.iloc[:, 6].astype(str).str.strip().str.upper() == 'PREPAID']
                
            all_filtered_dfs.append(filtered_df)
        
        browser.close()

    # Combine dataframes if two days were downloaded
    final_df = pd.concat(all_filtered_dfs, ignore_index=True)
    
    # Send filtered Prepaid orders to Google Sheet Webhook
    upload_to_gsheet(final_df)

if __name__ == "__main__":
    main()
