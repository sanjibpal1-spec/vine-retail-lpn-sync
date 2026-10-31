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
    
    # After 12 AM (00:00 to 19:59): fetch BOTH yesterday and today
    if 0 <= current_hour < 20:
        return [yesterday_str, today_str]
    # After 8 PM (20:00 to 23:59): fetch TODAY only
    else:
        return [today_str]

def download_vine_report(page, download_dir, target_date):
    print(f"Navigating to login page for target date: {target_date}")
    page.goto("https://multistorewh.vineretail.com/eRetailWeb/eRetailLogin", wait_until="networkidle")
    
    # 1. Login
    page.fill("input[name='username'], #username", os.environ["VINE_USER"])
    page.fill("input[name='password'], #password", os.environ["VINE_PASS"])
    
    login_selector = "input[type='submit'], input[value='Login'], button[type='submit'], #loginBtn, .btn-primary"
    page.click(login_selector)
    page.wait_for_load_state("networkidle")
    print("Logged in successfully.")
    
    # 2. Click Reports icon on left sidebar (Book Icon - 8th menu item down)
    print("Opening Reports menu...")
    page.wait_for_timeout(3000)
    
    # Use exact visual locator matching "Reports" icon/tooltip or left dark panel item
    reports_icon = page.locator("a:has-text('Reports'), [title='Reports'], i.fa-book, .left-side div, .sidebar div").filter(has_text="Reports").first
    if not reports_icon.is_visible():
        # Fallback to clicking the 8th icon link in the left dark sidebar navigation bar
        reports_icon = page.locator("nav a, div.sidebar a, ul.nav a").nth(7)
    
    reports_icon.click()
    page.wait_for_timeout(2000)
    
    # 3. Click "MIS Report 2.0"
    print("Selecting MIS Report 2.0...")
    mis_report_link = page.locator("text=MIS Report 2.0").first
    mis_report_link.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(3000)

    # Resolve active frame/page context
    report_container = page
    if len(page.frames) > 1:
        for frame in page.frames:
            if "Report" in frame.url or frame.name == "mainFrame":
                report_container = frame
                break
    
    # 4. Target VISIBLE dropdown selects only
    print("Setting report filters...")
    visible_selects = report_container.locator("select:visible")
    
    # Select Report Group = "Sales"
    visible_selects.nth(0).select_option(label="Sales")
    page.wait_for_timeout(1000)

    # Select Report Name = "PICK IPP Report"
    visible_selects.nth(1).select_option(label="PICK IPP Report")
    page.wait_for_timeout(1000)
    
    # 5. Fill Date Range (Format: DD/MM/YYYY - DD/MM/YYYY)
    date_input = report_container.locator("input:visible[type='text']").first
    date_input.fill(f"{target_date} - {target_date}")
    
    # 6. Select Location "JCT-CBE IM1 Temp"
    print("Selecting location JCT-CBE IM1 Temp...")
    select2_box = report_container.locator(".select2-selection, .select2-choices").first
    select2_box.click()
    page.wait_for_selector("li:has-text('JCT-CBE IM1 Temp')", state="visible")
    page.click("li:has-text('JCT-CBE IM1 Temp')")
    
    # 7. Click Export
    print("Clicking Export...")
    export_btn = report_container.locator("button:has-text('Export'), .btn-warning, input[value='Export']").first
    export_btn.click()
    page.wait_for_load_state("networkidle")
    
    print(f"Export requested. Waiting 3 minutes for report ({target_date}) to generate...")
    time.sleep(180)
    
    # 8. Refresh Pending Report tab
    print("Refreshing Pending Report screen...")
    refresh_icon = page.locator("img[src*='refresh'], .fa-sync, .fa-refresh, #pull_refresh, a.refresh_btn").first
    if refresh_icon.is_visible():
        refresh_icon.click()
    else:
        page.reload()
    time.sleep(5)
    
    # 9. Download file
    print("Downloading Excel file...")
    with page.expect_download(timeout=60000) as download_info:
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
            
            # Read Excel starting from Row 11 (Header is row 11 -> header=10)
            df = pd.read_excel(file_path, header=10)
            
            # Filter Column G (Order_Type) for "PREPAID"
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
