from playwright.sync_api import Playwright, sync_playwright
import playwright.async_api
from datetime import datetime, timedelta
import pandas as pd
import time
import os



# BASE_URL = "https://asap.dataforce.com.au/~em/login.php?client_code=em"
BASE_URL = "https://asap.dataforce.com.au/~em/main.php?&iDashboardId=&msg="

data = open("config.txt", "r")
for x in data:
    if 'username' in x:
        username = x.replace('username = ', '').replace('\n', '')
    if 'password' in x:
        password = x.replace('password = ', '').replace('\n', '')
    if 'destination_path' in x:
        destination_path = x.replace('destination_path = ', '').replace('\n', '')  


def login(page, context):
    time.sleep(1)
    if page.title() == 'Dataforce ASAP Login':
        print('Logging in ...')
        page.get_by_placeholder("Username").fill(username)
        page.get_by_placeholder("Password").fill(password)
        page.get_by_role("button", name="Login").click()
        context.storage_state(path="auth.json")
        
        
def save_file(download, file_name):
    
    downloaded_file_name = "downloaded.csv"
    download.save_as(downloaded_file_name)
    
    downloaded_file = pd.read_csv('downloaded.csv')
    current_date_today = datetime.now()
    current_date_today = current_date_today.strftime("%Y%m%d%H%M%S")
    file_name = destination_path + '/' +file_name + current_date_today + '.csv' 
    downloaded_file.to_csv(file_name, sep='|', index=False)
    print(f"File {file_name} downloaded successfully")
    os.remove(downloaded_file_name)    
    time.sleep(2)    
        
with sync_playwright() as playwright:

    current_date1 = datetime.now()
    one_year_ago = (current_date1 - timedelta(days=365)).strftime("%d-%m-%Y")
    current_date2 = datetime.now()
    next_day = (current_date2 + timedelta(days=1)).strftime("%d-%m-%Y")

    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(storage_state="auth.json")
    page = context.new_page()
    page.goto(BASE_URL)
    
    login(page, context)
        

    # -------------------------- File 1 -------------------
    page.locator("span").first.click()
    page.get_by_role("link", name="Energy Makeovers VEU").click() 
    page.frame_locator("iframe >> nth=1").get_by_text("Reports").click()
    page.frame_locator("iframe >> nth=1").frame_locator("#mainContent").locator("#report_single_67").get_by_text("Customer Invoice Summary").click()
    page.frame_locator("iframe >> nth=1").frame_locator("#mainContent").get_by_title("Scheduled Date (Install)").click()
    page.frame_locator("iframe >> nth=1").frame_locator("#mainContent").get_by_role("treeitem", name="Actual Completed Date (Install)").click()
    
    
    page.frame_locator("iframe >> nth=1").frame_locator("#mainContent").locator("#dFromDate").fill(one_year_ago)
    page.frame_locator("iframe >> nth=1").frame_locator("#mainContent").locator("#dToDate").fill(next_day)

    page.frame_locator("iframe >> nth=1").frame_locator("#mainContent").get_by_role("button", name="Show Summary").click()
    
    with page.expect_download() as download_info:
        page.frame_locator("iframe >> nth=1").frame_locator("#mainContent").get_by_role("button", name="Export To CSV").click()
    download = download_info.value
    save_file(download, "EMVIC-VCUSTOMER-INVOICE-SUMMARY-REPORT-1")

    # -------------------------- File 2 -------------------
    page.locator(".start-button").click()
    page.get_by_role("link", name="Energy Makeovers Solar PV").click()
    time.sleep(3)
    page.frame_locator("iframe >> nth=2").get_by_text("Reports").click()
    page.frame_locator("iframe >> nth=2").frame_locator("#mainContent").locator("#report_single_67").get_by_text("Customer Invoice Summary").click()
    page.frame_locator("iframe >> nth=2").frame_locator("#mainContent").get_by_label("Scheduled Date (Install)").locator("span").nth(1).click()
    page.frame_locator("iframe >> nth=2").frame_locator("#mainContent").get_by_role("treeitem", name="Actual Completed Date (Install)").click()
    page.frame_locator("iframe >> nth=2").frame_locator("#mainContent").locator("#dFromDate").fill(one_year_ago)
    page.frame_locator("iframe >> nth=2").frame_locator("#mainContent").locator("#dToDate").fill(next_day)

    page.frame_locator("iframe >> nth=2").frame_locator("#mainContent").get_by_role("button", name="Show Summary").click()
    with page.expect_download() as download1_info:
        page.frame_locator("iframe >> nth=2").frame_locator("#mainContent").get_by_role("button", name="Export To CSV").click()
    download1 = download1_info.value  
     
    save_file(download, "EMSOLAR-VCUSTOMER-INVOICE-SUMMARY-REPORT-1") 
   
    # -------------------------- File 3 -------------------
    page.locator(".start-button").click()
    page.get_by_role("link", name="Energy Makeovers (Aggregation) - VEU").click()
    time.sleep(3)
    page.frame_locator("iframe >> nth=3").get_by_text("Reports").click()
    page.frame_locator("iframe >> nth=3").frame_locator("#mainContent").locator("#report_single_67").get_by_text("Customer Invoice Summary").click()
    page.frame_locator("iframe >> nth=3").frame_locator("#mainContent").get_by_label("Scheduled Date (Install)").locator("span").nth(1).click()
    page.frame_locator("iframe >> nth=3").frame_locator("#mainContent").get_by_role("treeitem", name="Actual Completed Date (Install)").click()
    page.frame_locator("iframe >> nth=3").frame_locator("#mainContent").locator("#dFromDate").fill(one_year_ago)
    page.frame_locator("iframe >> nth=3").frame_locator("#mainContent").locator("#dToDate").fill(next_day)

    page.frame_locator("iframe >> nth=3").frame_locator("#mainContent").get_by_role("button", name="Show Summary").click()
    with page.expect_download() as download2_info:
        page.frame_locator("iframe >> nth=3").frame_locator("#mainContent").get_by_role("button", name="Export To CSV").click()
    download2 = download2_info.value 
     
    save_file(download2, "EMVEUAGG-VCUSTOMER-INVOICE-SUMMARY-REPORT-1")  


    # -------------------------- File 4 -------------------
    page.locator(".start-button").click()
    page.get_by_role("link", name="Energy Makeovers ESS").click()
    time.sleep(3)
    page.frame_locator("iframe >> nth=4").get_by_text("Reports").click()
    page.frame_locator("iframe >> nth=4").frame_locator("#mainContent").locator("#report_single_67").get_by_text("Customer Invoice Summary").click()
    page.frame_locator("iframe >> nth=4").frame_locator("#mainContent").get_by_label("Scheduled Date (Install)").locator("span").nth(1).click()
    page.frame_locator("iframe >> nth=4").frame_locator("#mainContent").get_by_role("treeitem", name="Actual Completed Date (Install)").click()
    page.frame_locator("iframe >> nth=4").frame_locator("#mainContent").locator("#dFromDate").fill(one_year_ago)
    page.frame_locator("iframe >> nth=4").frame_locator("#mainContent").locator("#dToDate").fill(next_day)

    page.frame_locator("iframe >> nth=4").frame_locator("#mainContent").get_by_role("button", name="Show Summary").click()
    with page.expect_download() as download3_info:
        page.frame_locator("iframe >> nth=4").frame_locator("#mainContent").get_by_role("button", name="Export To CSV").click()
    download3 = download3_info.value 
     
    save_file(download3, "EMNSW-VCUSTOMER-INVOICE-SUMMARY-REPORT-1")     
       
    time.sleep(2)
    context.close()
    browser.close()
