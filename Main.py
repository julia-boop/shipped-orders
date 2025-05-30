#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#IMPORTS ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
import pandas as pd
import smtplib
import os
import gspread
import openpyxl
import tempfile
import time
import base64
import re
import json
from datetime import datetime
from dateutil.relativedelta import relativedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from google.oauth2.service_account import Credentials
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from io import BytesIO
from dotenv import load_dotenv
from datetime import datetime, timedelta


#For development
# service = Service("/usr/local/bin/chromedriver") 

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#SETTINGS~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#For production
import chromedriver_autoinstaller
chromedriver_path = chromedriver_autoinstaller.install()
service = Service(chromedriver_path)

user_data_dir = tempfile.mkdtemp()
script_dir = os.path.dirname(os.path.abspath(__file__))
download_path = os.path.join(script_dir, 'LogiwaOrders')
os.makedirs(download_path, exist_ok=True)


chrome_options = Options()
chrome_options.add_argument("--headless")  
chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
#For develpment
#chrome_options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

#For production
chrome_options.binary_location = "/usr/bin/chromium"
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_path,  
    "download.prompt_for_download": False,       
    "download.directory_upgrade": True,        
    "safebrowsing.enabled": True               
})

driver = webdriver.Chrome(service=service, options=chrome_options)

load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#GET LOGIWA FILE~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


def wait_for_download_to_finish(download_path, timeout=1200):
    """
    Waits until the download folder has no .crdownload or .part files (indicating download in progress).
    """
    seconds = 0
    while seconds < timeout:
        files = os.listdir(download_path)
        if any(file.endswith(('.crdownload', '.part')) for file in files):
            time.sleep(1)
            seconds += 1
        else:
            return True
    raise TimeoutError("Download did not complete within timeout.")

def get_latest_file(directory):
    files = [os.path.join(directory, f) for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]
    if not files:
        return None  
    latest_file = max(files, key=os.path.getmtime)  
    return latest_file

def get_logiwa_file(date_entry=None):
    driver.get("https://app.logiwa.com/en/Login")

    username_field = driver.find_element(By.ID, "UserName")
    password_field = driver.find_element(By.ID, "Password")

    print(username_field)
    print(password_field)

    time.sleep(3)

    user = os.getenv("LOGIWA_USERNAME")
    password = os.getenv("LOGIWA_PASSWORD")

    username_field.send_keys(user)
    password_field.send_keys(password)

    login_button = driver.find_element(By.ID, "LoginButton")
    login_button.click()

    time.sleep(3)

    login_handle = None

    try:
        login_handle = driver.find_element(By.CSS_SELECTOR, ".bootbox-body")
    except Exception as e:
        print("No login handle needed")

    if login_handle:
        buttons = driver.find_elements(By.CLASS_NAME, "btn-success")
        for b in buttons:
            text = b.text
            if text == "Ok":
                b.click()
                time.sleep(3)
    else:
        print("No login handle needed")
    
    driver.get("https://app.logiwa.com/en/WMS/WarehouseOrder")

    time.sleep(20)

    dropdown_btn = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div[3]/form/div/div[1]/div[2]/div[7]/div[2]/div/button")
    wait = WebDriverWait(driver, 20)
    dropdown_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div[3]/form/div/div[1]/div[2]/div[7]/div[2]/div/button")))
    driver.execute_script("arguments[0].scrollIntoView(true);", dropdown_btn)
    driver.execute_script("arguments[0].click();", dropdown_btn)
    for position in [2, 8]:
        option = driver.find_element(By.XPATH, f"/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div[3]/form/div/div[1]/div[2]/div[7]/div[2]/div/ul/li[{position}]/a/label")
        option.click()

    time.sleep(10)

    dropdown_btn_2 = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div[3]/form/div/div[1]/div[2]/div[16]/div[2]/div/button")
    wait = WebDriverWait(driver, 20)
    dropdown_btn_2 = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div[3]/form/div/div[1]/div[2]/div[16]/div[2]/div/button")))
    driver.execute_script("arguments[0].scrollIntoView(true);", dropdown_btn_2)
    driver.execute_script("arguments[0].click();", dropdown_btn_2)

    for position in [2, 17]:
        option = driver.find_element(By.XPATH, f"/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div[3]/form/div/div[1]/div[2]/div[16]/div[2]/div/ul/li[{position}]/a/label")
        option.click()
    
    time.sleep(10)

    date_input = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div[3]/form/div/div[2]/div/div[5]/div[2]/div/input")
    first_day = datetime.today().replace(day=1)
    today = datetime.today()
    first_day_prev_month = today.replace(day=1) - relativedelta(months=1)
    date_range = date_entry if date_entry else f"01.01.2025 00:00:00 - {today.strftime('%m.%d.%Y')} 00:00:00"
    print(date_range) 
    date_input.send_keys(date_range)

    time.sleep(10)
    button_search = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div/div/div[1]/button[1]")))
    driver.execute_script("arguments[0].scrollIntoView(true);", button_search)
    driver.execute_script("arguments[0].click();", button_search)

    time.sleep(10)

    button_excel = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div[3]/div/div[5]/div/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/span")))
    driver.execute_script("arguments[0].scrollIntoView(true);", button_excel)
    driver.execute_script("arguments[0].click();", button_excel)


    time.sleep(200)

    driver.quit() 

    latest_file = get_latest_file(download_path)

    if latest_file:
        print(f"Latest download file: {latest_file}")
        return latest_file
    else:
        print(f"No files found in the directory {download_path}")
        return



#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#GET GOOGLE FILE ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


def get_googlesheets_file():
    service_account_json = base64.b64decode(os.getenv("SERVICE_ACCOUNT_FILE")).decode("utf-8")
    SERVICE_ACCOUNT_FILE = json.loads(service_account_json)

    SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 
            'https://www.googleapis.com/auth/drive']

    creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    client = gspread.authorize(creds)

    spreadsheet = client.open("Order History")

    sheet = spreadsheet.worksheet("Order History")  

    data = sheet.get_all_values()

    output_folder = os.path.join(os.path.dirname(__file__), 'GoogleSheetsFile')
    os.makedirs(output_folder, exist_ok=True)

    file_name = os.path.join(output_folder, f"AllOrders-{datetime.today()}.xlsx")

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Ordenes"

    for row in data:
        worksheet.append(row)

    workbook.save(file_name)
    print(f" File '{file_name}' saved successfully!")
    return file_name

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#COMPARE FILES ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

def compare_files(date_entry=None):
    file_gs = get_googlesheets_file()
    file_lw = get_logiwa_file(date_entry if date_entry else None)

    df_gs = pd.read_excel(file_gs)
    df_lw = pd.read_excel(file_lw)

    df_gs = df_gs[df_gs['Status'] == 'Shipped']
    df_lw = df_lw[(df_lw['Order Status'] != 'Shipped') & (df_lw['Operation Status'] != 'Shipped')]

    for col in ['Order', 'PO#', 'DC/Store']:
        df_gs[col] = df_gs[col].astype(str).str.strip()
        df_gs[col] = df_gs[col].astype(str).str.replace('#', '', regex=False).str.strip()

    for col in ['Logiwa Order #', 'Customer Order #']:
        df_lw[col] = df_lw[col].astype(str).str.strip()
        df_lw[col] = df_lw[col].astype(str).str.replace('#', '', regex=False).str.strip()

    matches = []
    no_matches = []

    matched_logiwa_indices = set() 

    for lw_idx, lw_row in df_lw.iterrows():
        if lw_idx in matched_logiwa_indices:
            continue

        logiwa_order = lw_row['Logiwa Order #']
        customer_order = lw_row['Customer Order #']
        match_found = False

        df_gs['order_dc'] = df_gs['Order'] + '-' + df_gs['DC/Store']
        df_gs['po_dc'] = df_gs['PO#'] + '-' + df_gs['DC/Store']
        df_gs['hashpo_dc'] = df_gs['PO#'] + '-' + df_gs['DC/Store']
        df_gs['order_dc_cus'] = df_gs['Order'] + '-' + df_gs['DC/Store']


        match_df = df_gs[df_gs['Order'].str.contains(re.escape(str(logiwa_order)), case=False, na=False)]
        if len(match_df) == 1:
            match = match_df.iloc[0]
            if str(lw_row['Client']).strip().lower() not in str(match['Client']).strip().lower():
                continue
            combined_row = {
                'Client': match['Client'],  
                'Customer': match['Customer'],
                'Order': match['Order'],
                '#PO': match['PO#'],
                'Tracker Status': match['Status'],
                'Tracker Units': match['Units Shipped'],
                'Logiwa Client': lw_row['Client'],
                'Logiwa Order #': lw_row['Logiwa Order #'],
                'Customer Order #': lw_row['Customer Order #'],
                'Logiwa Status': lw_row['Order Status'],
                'Logiwa Units': lw_row['Nof Products'],
            }
            matches.append(pd.DataFrame([combined_row]))
            print(f"Matched at 1 {combined_row}")
            matched_logiwa_indices.add(lw_idx)
            match_found = True

        elif len(match_df) > 1:
            match_df2 = df_gs[df_gs['order_dc'].str.contains(re.escape(str(logiwa_order)), case=False, na=False)]
            if len(match_df2) == 1:
                match = match_df2.iloc[0]
                if str(lw_row['Client']).strip().lower() not in str(match['Client']).strip().lower():
                    continue
                combined_row = {
                    'Client': match['Client'],  
                    'Customer': match['Customer'],
                    'Order': match['Order'],
                    '#PO': match['PO#'],
                    'Tracker Status': match['Status'],
                    'Tracker Units': match['Units Shipped'],
                    'Logiwa Client': lw_row['Client'],
                    'Logiwa Order #': lw_row['Logiwa Order #'],
                    'Customer Order #': lw_row['Customer Order #'],
                    'Logiwa Status': lw_row['Order Status'],
                    'Logiwa Units': lw_row['Nof Products'],
                }
                matches.append(pd.DataFrame([combined_row]))
                print(f"Matched at 1.1 {combined_row}")
                matched_logiwa_indices.add(lw_idx)
                match_found = True

        if not match_found:
            match_df = df_gs[df_gs['PO#'].str.contains(re.escape(str(logiwa_order)), case=False, na=False)]
            if len(match_df) == 1:
                match = match_df.iloc[0]
                if str(lw_row['Client']).strip().lower() not in str(match['Client']).strip().lower():
                    continue
                combined_row = {
                    'Client': match['Client'], 
                    'Customer': match['Customer'],
                    'Order': match['Order'],
                    '#PO': match['PO#'],
                    'Tracker Status': match['Status'],
                    'Tracker Units': match['Units Shipped'],
                    'Logiwa Client': lw_row['Client'],
                    'Logiwa Order #': lw_row['Logiwa Order #'],
                    'Customer Order #': lw_row['Customer Order #'],
                    'Logiwa Status': lw_row['Order Status'],
                    'Logiwa Units': lw_row['Nof Products'],
                }
                matches.append(pd.DataFrame([combined_row]))
                print(f"Matched at 2 {combined_row}")
                matched_logiwa_indices.add(lw_idx)
                match_found = True

            elif len(match_df) > 1:
                match_df2 = df_gs[df_gs['po_dc'].str.contains(re.escape(str(logiwa_order)), case=False, na=False)]
                if len(match_df2) == 1:
                    match = match_df2.iloc[0]
                    if str(lw_row['Client']).strip().lower() not in str(match['Client']).strip().lower():
                        continue
                    combined_row = {
                        'Client': match['Client'],  
                        'Customer': match['Customer'],
                        'Order': match['Order'],
                        '#PO': match['PO#'],
                        'Tracker Status': match['Status'],
                        'Tracker Units': match['Units Shipped'],
                        'Logiwa Client': lw_row['Client'],
                        'Logiwa Order #': lw_row['Logiwa Order #'],
                        'Customer Order #': lw_row['Customer Order #'],
                        'Logiwa Status': lw_row['Order Status'],
                        'Logiwa Units': lw_row['Nof Products'],
                    }
                    matches.append(pd.DataFrame([combined_row]))
                    print(f"Matched at 2.1 {combined_row}")
                    matched_logiwa_indices.add(lw_idx)
                    match_found = True

        if not match_found:
            match_df = df_gs[df_gs['PO#'].str.contains(re.escape(str(customer_order)), case=False, na=False)]
            if len(match_df) == 1:
                match = match_df.iloc[0]
                if str(lw_row['Client']).strip().lower() not in str(match['Client']).strip().lower():
                    continue
                combined_row = {
                    'Client': match['Client'], 
                    'Customer': match['Customer'],
                    'Order': match['Order'],
                    '#PO': match['PO#'],
                    'Tracker Status': match['Status'],
                    'Tracker Units': match['Units Shipped'],
                    'Logiwa Client': lw_row['Client'],
                    'Logiwa Order #': lw_row['Logiwa Order #'],
                    'Customer Order #': lw_row['Customer Order #'],
                    'Logiwa Status': lw_row['Order Status'],
                    'Logiwa Units': lw_row['Nof Products'],
                }
                matches.append(pd.DataFrame([combined_row]))
                print(f"Matched at 3 {combined_row}")
                matched_logiwa_indices.add(lw_idx)
                match_found = True

            elif len(match_df) > 1:
                match_df2 = df_gs[df_gs['hashpo_dc'].str.contains(re.escape(str(customer_order)), case=False, na=False)]
                if len(match_df2) == 1:
                    match = match_df2.iloc[0]
                    if str(lw_row['Client']).strip().lower() not in str(match['Client']).strip().lower():
                        continue
                    combined_row = {
                        'Client': match['Client'],  
                        'Customer': match['Customer'],
                        'Order': match['Order'],
                        '#PO': match['PO#'],
                        'Tracker Status': match['Status'],
                        'Tracker Units': match['Units Shipped'],
                        'Logiwa Client': lw_row['Client'],
                        'Logiwa Order #': lw_row['Logiwa Order #'],
                        'Customer Order #': lw_row['Customer Order #'],
                        'Logiwa Status': lw_row['Order Status'],
                        'Logiwa Units': lw_row['Nof Products'],
                    }
                    matches.append(pd.DataFrame([combined_row]))
                    print(f"Matched at 3.1 {combined_row}")
                    matched_logiwa_indices.add(lw_idx)
                    match_found = True

        if not match_found:
            match_df = df_gs[df_gs['Order'].str.contains(re.escape(str(customer_order)), case=False, na=False)]
            if len(match_df) == 1:
                match = match_df.iloc[0]
                if str(lw_row['Client']).strip().lower() not in str(match['Client']).strip().lower():
                    continue
                combined_row = {
                    'Client': match['Client'],  
                    'Customer': match['Customer'],
                    'Order': match['Order'],
                    '#PO': match['PO#'],
                    'Tracker Status': match['Status'],
                    'Tracker Units': match['Units Shipped'],
                    'Logiwa Client': lw_row['Client'],
                    'Logiwa Order #': lw_row['Logiwa Order #'],
                    'Customer Order #': lw_row['Customer Order #'],
                    'Logiwa Status': lw_row['Order Status'],
                    'Logiwa Units': lw_row['Nof Products'],
                }
                matches.append(pd.DataFrame([combined_row]))
                print(f"Matched 4 {combined_row}")
                matched_logiwa_indices.add(lw_idx)
                match_found = True
        
        elif len(match_df) > 1:
            match_df2 = df_gs[df_gs['order_dc_cus'].str.contains(re.escape(str(customer_order)), case=False, na=False)]
            if len(match_df2) == 1:
                match = match_df2.iloc[0]
                if str(lw_row['Client']).strip().lower() not in str(match['Client']).strip().lower():
                    continue
                combined_row = {
                    'Client': match['Client'],  
                    'Customer': match['Customer'],
                    'Order': match['Order'],
                    '#PO': match['PO#'],
                    'Tracker Status': match['Status'],
                    'Tracker Units': match['Units Shipped'],
                    'Logiwa Client': lw_row['Client'],
                    'Logiwa Order #': lw_row['Logiwa Order #'],
                    'Customer Order #': lw_row['Customer Order #'],
                    'Logiwa Status': lw_row['Order Status'],
                    'Logiwa Units': lw_row['Nof Products'],
                }
                matches.append(pd.DataFrame([combined_row]))
                print(f"Matched at 4.1 {combined_row}")
                matched_logiwa_indices.add(lw_idx)
                match_found = True


        if not match_found:
            no_matches.append(lw_row)

    if matches:
        final_match = pd.concat(matches, ignore_index=True)
    else:
        final_match = pd.DataFrame()
    
    final_match = final_match.sort_values(by="Client", ascending=True)
    final_match['Tracker Units'] = pd.to_numeric(final_match['Tracker Units'], errors='coerce')
    final_match['Logiwa Units'] = pd.to_numeric(final_match['Logiwa Units'], errors='coerce')
    final_match['Difference in Units'] = final_match['Tracker Units'] - final_match['Logiwa Units']
    final_match['Difference in Units'] = final_match['Difference in Units'].abs()
    print(final_match.head(20))

    return final_match

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#SEND EMAIL ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

def send_email_with_matches(matched_orders):
    row_count = len(matched_orders)
    html_table = matched_orders.to_html(index=False, escape=False, border=0)

    html_content = f"""
    <html>
        <body style="font-family: Arial, sans-serif; background-color: #0c1c24; padding: 20px;">
            <h1 style="color: white; font-size: 24px; text-align: center;"> {row_count} Pending orders:</h1>
            <table style="width: 80%; max-width: 600px; margin: 0 auto; border-collapse: collapse; background: #fff; box-shadow: 0px 4px 8px rgba(0, 0, 0, 0.1); border-radius: 10px; overflow: hidden;">
                <tr style="background-color: #182937; color: white;">
                    <th style="padding: 12px; text-align: left; border-bottom: 1px solid #ddd;">Client</th>
                    <th style="padding: 12px; text-align: left; border-bottom: 1px solid #ddd;">Customer</th>
                    <th style="padding: 12px; text-align: left; border-bottom: 1px solid #ddd;">Order</th>
                    <th style="padding: 12px; text-align: left; border-bottom: 1px solid #ddd;">#PO</th>
                    <th style="padding: 12px; text-align: left; border-bottom: 1px solid #ddd;">Logiwa Order</th>
                    <th style="padding: 12px; text-align: left; border-bottom: 1px solid #ddd;">Logiwa Units</th>
                    <th style="padding: 12px; text-align: left; border-bottom: 1px solid #ddd;">Tracker Units</th>
                    <th style="padding: 12px; text-align: left; border-bottom: 1px solid #ddd;">Difference</th>
                </tr>
                {''.join(
                    f"<tr style='background-color: #f9f9f9;'>"
                    + "".join(
                        f"<td style='padding: 12px; border-bottom: 1px solid #ddd;'>{'' if pd.isna(row.get(col)) or str(row.get(col)).strip() == '' else row.get(col)}</td>"
                        for col in ['Client', 'Customer', 'Order', '#PO', 'Logiwa Order #']
                    )
                    + f"<td style='padding: 12px; border-bottom: 1px solid #ddd;'>{'' if pd.isna(row.get('Logiwa Units')) or str(row.get('Logiwa Units')).strip() == '' else int(row.get('Logiwa Units'))}</td>"
                    + f"<td style='padding: 12px; border-bottom: 1px solid #ddd;'>{'' if pd.isna(row.get('Tracker Units')) or str(row.get('Tracker Units')).strip() == '' else int(row.get('Tracker Units'))}</td>"
                    + (
                        f"<td style='padding: 12px; border-bottom: 1px solid #ddd; color: #DAA520;'>Incomplete</td>"
                        if pd.isna(row.get('Difference in Units')) or str(row.get('Difference in Units')).strip() == ''
                        else f"<td style='padding: 12px; border-bottom: 1px solid #ddd; color: {'red' if row['Difference in Units'] != 0 else 'black'};'>{int(row['Difference in Units'])}</td>"
                    )
                    + "</tr>"
                    for _, row in matched_orders.iterrows()
                )}
            </table>
            <img src="https://www.the5411.com/wp-content/uploads/2024/03/5411-Distribution-Logo-1.png" 
                style="display: block; margin: 20px auto; max-width: 150px; height: auto;">
            <p style="text-align: center; font-size: 12px; color: white; margin-top: 8px;">Dallas, Texas, United States</p>
        </body>
    </html>
    """

    #receiver_email = ["buenos-aires@the5411.com", "jcordero@the5411.com", "isalazar@the5411.com"]
    #receiver_email = ", ".join(receiver_email)
    receiver_email = "jcordero@the5411.com"
    subject = "🚀 Pending orders to ship in Logiwa:"

    message = MIMEMultipart()
    message["From"] = os.getenv("SENDER_EMAIL")
    message["To"] = receiver_email
    message["Subject"] = subject

    message.attach(MIMEText(html_content, "html"))

    excel_buffer = BytesIO()
    matched_orders.to_excel(excel_buffer, index=False, engine='openpyxl')
    excel_buffer.seek(0)

    attachment = MIMEApplication(excel_buffer.read(), _subtype="xlsx")
    attachment.add_header(
        "Content-Disposition",
        "attachment",
        filename="PendingOrders.xlsx"
    )
    message.attach(attachment)

    with smtplib.SMTP_SSL(os.getenv("SMTP_SERVER"), int(os.getenv("SMTP_PORT"))) as server:
        server.login(os.getenv("SENDER_EMAIL"), os.getenv("EMAIL_PASSWORD"))  # Use App Password if 2FA is enabled
        server.sendmail(os.getenv("SENDER_EMAIL"), receiver_email, message.as_string())

    print("✅ Email sent successfully!")







#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#CONDITIONAL CALLING OF FUNCTIONS ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

matches = compare_files()
send_email_with_matches(matches)

# Uncomment the following lines to run the script only on specific days


# today = datetime.today()

# def is_last_weekday_of_month():
#     tomorrow = today + timedelta(days=1)
#     return today.weekday() < 5 and tomorrow.day == 1

# def is_first_weekday_of_month():
#     day = today.day
#     weekday = today.weekday()
#     return day <= 3 and weekday < 5

# if today.weekday() == 4 or is_last_weekday_of_month():
#     print("✅ Running the script (Friday or last weekday)...")
#     matches = compare_files()
#     send_email_with_matches(matches)
# elif is_first_weekday_of_month():
#     print("✅ Running the script (first weekday of the month)...")
#     last_day_of_last_month = today - timedelta(days=1)
#     first_day_of_last_month = last_day_of_last_month.replace(day=1)
#     start = first_day_of_last_month.strftime("%m.%d.%Y") + " 00:00:00"
#     end = last_day_of_last_month.strftime("%m.%d.%Y") + " 00:00:00"
#     date_entry = f"{start} - {end}"
#     matches = compare_files(date_entry)
#     send_email_with_matches(matches)
# else:
#     print("⏳ Not a trigger day. Exiting...")
