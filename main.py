from playwright.sync_api import Playwright, sync_playwright
import playwright.async_api
from datetime import datetime, timedelta
import pandas as pd
import time
import os

current_date = datetime.now()

# BASE_URL = "https://asap.dataforce.com.au/~em/login.php?client_code=em"
BASE_URL = "https://asap.dataforce.com.au/~em/main.php?&iDashboardId=&msg="

data = open("config.txt", "r")
for x in data:
    if 'username' in x:
        username = x.replace('username = ', '').replace('\n', '')
    if 'password' in x:
        password = x.replace('password = ', '').replace('\n', '')
    if 'download_path' in x:
        download_path = x.replace('download_path = ', '').replace('\n', '')  


def login(page, context):
    time.sleep(1)
    if page.title() == 'Dataforce ASAP Login':
        print('Logging in ...')
        page.get_by_placeholder("Username").fill(username)
        page.get_by_placeholder("Password").fill(password)
        page.get_by_role("button", name="Login").click()
        context.storage_state(path="auth.json")
        
        
with sync_playwright() as playwright:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(storage_state="auth.json")
    page = context.new_page()
    page.goto(BASE_URL)
    
    login(page, context)
        
    one_year_ago = current_date - timedelta(days=365)
    one_year_ago = one_year_ago.strftime("%d-%m-%Y")
        
    page.locator("span").first.click()
    page.get_by_role("link", name="Energy Makeovers VEU").click() # This report name is different in each iteration.
    page.frame_locator("iframe >> nth=1").get_by_text("Reports").click()
    page.frame_locator("iframe >> nth=1").frame_locator("#mainContent").locator("#report_single_67").get_by_text("Customer Invoice Summary").click()
    
    
    page.frame_locator("iframe >> nth=1").frame_locator("#mainContent").get_by_title("Scheduled Date (Install)").click()
    page.frame_locator("iframe >> nth=1").frame_locator("#mainContent").get_by_role("treeitem", name="Actual Completed Date (Install)").click()
    page.frame_locator("iframe >> nth=1").frame_locator("#mainContent").locator("#dFromDate").fill(one_year_ago)


    page.frame_locator("iframe >> nth=1").frame_locator("#mainContent").get_by_role("button", name="Show Summary").click()
    
    with page.expect_download() as download_info:
        page.frame_locator("iframe >> nth=1").frame_locator("#mainContent").get_by_role("button", name="Export To CSV").click()
    download = download_info.value
    downloaded_file_name = download_path + "/downloaded.csv"
    download.save_as(downloaded_file_name)
    
    downloaded_file = pd.read_csv(download_path + '/downloaded.csv')

    current_date = current_date.strftime("%Y%m%d%H%M%S")
    file_name = download_path + '/EMVIC-VCUSTOMER-INVOICE-SUMMARY-REPORT-1' + current_date + '.csv' # This final file name is different in each iteration.
    downloaded_file.to_csv(file_name, sep='|', index=False)
    print("File Downloaded successfully")
    os.remove(downloaded_file_name)           
       
    time.sleep(2)
    context.close()
    browser.close()
