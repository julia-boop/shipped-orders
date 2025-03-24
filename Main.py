import pandas as pd
import smtplib
import os
import gspread
import openpyxl
import tempfile
import time
import base64
import json
from datetime import datetime
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
from dotenv import load_dotenv
from datetime import datetime

DEBUG = True

if DEBUG:
    service = Service("/usr/local/bin/chromedriver") 
else:
    import chromedriver_autoinstaller
    chromedriver_path = chromedriver_autoinstaller.install()
    service = Service(chromedriver_path)

user_data_dir = tempfile.mkdtemp()
script_dir = os.path.dirname(os.path.abspath(__file__))
download_path = os.path.join(script_dir, 'LogiwaOrders')
os.makedirs(download_path, exist_ok=True)


chrome_options = Options()
#chrome_options.add_argument("--headless")  
chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_path,  
    "download.prompt_for_download": False,       
    "download.directory_upgrade": True,        
    "safebrowsing.enabled": True               
})

service = Service("/usr/local/bin/chromedriver")
driver = webdriver.Chrome(service=service, options=chrome_options)

load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))

def get_latest_file(directory):
    files = [os.path.join(directory, f) for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]
    if not files:
        return None  
    latest_file = max(files, key=os.path.getmtime)  
    return latest_file

def get_logiwa_file():
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

    time.sleep(3)

    dropdown_btn = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div[3]/form/div/div[1]/div[2]/div[7]/div[2]/div/button")
    dropdown_btn.click()
    for position in [3, 4, 5, 6, 7]:
        option = driver.find_element(By.XPATH, f"/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div[3]/form/div/div[1]/div[2]/div[7]/div[2]/div/ul/li[{position}]/a/label")
        option.click()
    
    time.sleep(3)

    date_input = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div[3]/form/div/div[2]/div/div[5]/div[2]/div/input")
    first_day = datetime.today().replace(day=1)
    today = datetime.today()
    date_range = f"{first_day.strftime('%m.%d.%Y')} 00:00:00 - {today.strftime('%m.%d.%Y')} 00:00:00"
    print(date_range) 
    date_input.send_keys(date_range)

    time.sleep(3)
    
    button_search = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div/div[1]/button[1]")
    button_search.click()

    time.sleep(3)

    button_excel = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/div[3]/div/div[5]/div/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/span")
    button_excel.click()

    time.sleep(20)

    driver.quit() 

    latest_file = get_latest_file(download_path)

    if latest_file:
        print(f"Latest download file: {latest_file}")
        return latest_file
    else:
        print(f"No files found in the directory {download_path}")
        return



def get_googlesheets_file():
    service_account_json = base64.b64decode(os.getenv("SERVICE_ACCOUNT_FILE")).decode("utf-8")
    SERVICE_ACCOUNT_FILE = json.loads(service_account_json)

    SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 
            'https://www.googleapis.com/auth/drive']

    creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    client = gspread.authorize(creds)

    spreadsheet = client.open("Prioridad de Ordenes")

    sheet = spreadsheet.worksheet("Ordenes")  

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



def compare_files():
    file_gs = get_googlesheets_file()
    file_lw = get_logiwa_file()

    df_gs = pd.read_excel(file_gs)
    df_lw = pd.read_excel(file_lw)

    df_gs = df_gs[df_gs['Status'] == 'Shipped']
    df_gs['Order'] = df_gs['Order'].str.strip().replace('', None)       
    df_gs['PO#'] = df_gs['PO#'].str.strip().replace('', None)
    df_lw['Logiwa Order #'] = df_lw['Logiwa Order #'].astype(str).str.strip()
    df_lw['Customer Order #'] = df_lw['Customer Order #'].astype(str).str.strip()

    matches = []
    matches.append(df_gs.merge(df_lw, left_on='Order', right_on='Logiwa Order #', how='inner'))
    matches.append(df_gs.merge(df_lw, left_on='Order', right_on='Customer Order #', how='inner'))
    matches.append(df_gs.merge(df_lw, left_on='PO#', right_on='Logiwa Order #', how='inner'))
    matches.append(df_gs.merge(df_lw, left_on='PO#', right_on='Customer Order #', how='inner'))
    matches.append(df_gs.merge(df_lw, left_on='DC/Store', right_on='Logiwa Order #', how='inner'))
    matches.append(df_gs.merge(df_lw, left_on='DC/Store', right_on='Customer Order #', how='inner'))

    final_match = pd.concat(matches).drop_duplicates()
    final_match = final_match.sort_values(by="Client_x", ascending=True)

    print(final_match)
    return final_match



def send_email_with_matches(matched_orders):
    html_table = matched_orders.to_html(index=False, escape=False, border=0)

    html_content = f"""
    <html>
        <body style="font-family: Arial, sans-serif; background-color: #0c1c24; padding: 20px;">
            <h1 style="color: white; font-size: 24px; text-align: center;">Pending orders:</h1>
            <table style="width: 80%; max-width: 600px; margin: 0 auto; border-collapse: collapse; background: #fff; box-shadow: 0px 4px 8px rgba(0, 0, 0, 0.1); border-radius: 10px; overflow: hidden;">
                <tr style="background-color: #182937; color: white;">
                    <th style="padding: 12px; text-align: left; border-bottom: 1px solid #ddd;">Client</th>
                    <th style="padding: 12px; text-align: left; border-bottom: 1px solid #ddd;">Customer</th>
                    <th style="padding: 12px; text-align: left; border-bottom: 1px solid #ddd;">Order</th>
                    <th style="padding: 12px; text-align: left; border-bottom: 1px solid #ddd;">#PO</th>
                </tr>
                {''.join(
                    f"<tr style='background-color: #f9f9f9;'>"
                    f"<td style='padding: 12px; border-bottom: 1px solid #ddd;'>{row['Client_x']}</td>"
                    f"<td style='padding: 12px; border-bottom: 1px solid #ddd;'>{row['Customer']}</td>"
                    f"<td style='padding: 12px; border-bottom: 1px solid #ddd;'>{row['Order']}</td>"
                    f"<td style='padding: 12px; border-bottom: 1px solid #ddd;'>{row['PO#']}</td>"
                    f"</tr>" for _, row in matched_orders.iterrows()
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

    with smtplib.SMTP_SSL(os.getenv("SMTP_SERVER"), int(os.getenv("SMTP_PORT"))) as server:
        server.login(os.getenv("SENDER_EMAIL"), os.getenv("EMAIL_PASSWORD"))  # Use App Password if 2FA is enabled
        server.sendmail(os.getenv("SENDER_EMAIL"), receiver_email, message.as_string())

    print("✅ Email sent successfully!")

matches = compare_files()
send_email_with_matches(matches)