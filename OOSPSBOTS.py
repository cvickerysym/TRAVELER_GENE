import time
import os
import pandas as pd
import ctypes
import win32print
from PyPDFForm import PdfWrapper
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from Auth import username, passkey
from datetime import datetime


def setup_driver():
    options = Options()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def login_to_site(driver, url):
    driver.get(url)
    driver.maximize_window()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div/form/div/div/div/input"))).click()
    Domain_login = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="username"]')))
    Domain_login.send_keys(username)
    domain_signin = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="loginbutton"]')))
    domain_signin.click()
    user_login = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="usernameForm"]/div[2]/div/div[1]/input')))
    user_login.send_keys(username)
    signin = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="usernameForm"]/div/button')))
    signin.click()
    password = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="passwordForm"]/div/div/div/input')))
    password.send_keys(passkey)
    sign_on = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="passwordForm"]/div/button')))
    sign_on.click()
    authenticate = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mechanismSelectionForm"]/div/button')))
    authenticate.click()
    time.sleep(15)

def download_file(driver):
    try:
        hover_element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="grid"]/div[2]/div[1]/div/div[2]/div')))
        actions = ActionChains(driver)
        actions.move_to_element(hover_element).perform()
        more_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="grid"]/div[2]/div[1]/div/div[2]/div/button[2]')))
        more_button.click()
        download_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="download"]')))
        download_button.click()
        DL_table = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/div[3]/div/div[2]/div/button')))
        DL_table.click()
    except Exception as e:
        print("An error occurred:", e)
        driver.save_screenshot("error_screenshot.png")
    time.sleep(7)

def convert_to_csv(download_dir):
    files = os.listdir(download_dir)
    latest_file = max([download_dir + "/" + f for f in files], key=os.path.getctime)
    if latest_file.endswith('.xlsx'):
        df = pd.read_excel(latest_file)
        csv_file = latest_file.replace('.xlsx', '.csv')
        df.to_csv(csv_file, index=False)
        print(f"Converted {latest_file} to {csv_file}")
    else:
        print("Downloaded file is not an Excel file")

def extract_data(download_dir):
    csv_df = os.listdir(download_dir)
    latest_file = max([download_dir + "/" + f for f in csv_df], key=os.path.getctime)
    df = pd.read_csv(latest_file)
    data = df[['BOT_ID', 'PULL SCORE', 'ALARMS (COUNTS)']]
    return data

def fill_pdf(template_path, output_path, data_dict):
    try:
        filled_pdf = PdfWrapper(template_path).fill(data_dict)
        with open(output_path, "wb") as output_file:
            output_file.write(filled_pdf.read())
        print(f"PDF created and filled: {output_path}")
    except Exception as e:
        print(f"An error occurred while creating the PDF: {e}")

def create_pdfs(data, bot_list, alpha_c_template_path, one_point_zero_template_path):
    date_time = str(datetime.now().date())
    pdf_files = []
    for index, row in data.iterrows():
        bot_id = row['BOT_ID']
        pull_score = row['PULL SCORE']
        alarms = row['ALARMS (COUNTS)']
        data_dict = {
            'Maint needed': 'Yes',
            'Maintenance needed': 'Yes',
            'Bot ID': str(bot_id),
            'Pull score': str(pull_score),
            'Location': ' ',
            'Datetime of removal': date_time,
            'Inductions': ' ',
            'Removals': ' ',
            'Qlik reasons': str(alarms),
            'Non Qlik reasons': 'N/A',
        }
        if bot_id in bot_list:
            if bot_id >= 20000:
                output_path = f"C:/Users/cvickery/TRAVELER_PDFs{date_time}/OnePointZero_filled_{bot_id}_{date_time}.pdf"
                fill_pdf(one_point_zero_template_path, output_path, data_dict)
            else:
                output_path = f"C:/Users/cvickery/TRAVELER_PDFs{date_time}/AlphaC_filled_{bot_id}_{date_time}.pdf"
                fill_pdf(alpha_c_template_path, output_path, data_dict)
            pdf_files.append(output_path)
    return pdf_files

def print_files_to_printer(pdf_files, printer_name):
    # Set the printer name
    printer = win32print.OpenPrinter(printer_name)
    try:
        for pdf_file in pdf_files:
            # Use ShellExecute via ctypes to send the PDF to the default printer
            result = ctypes.windll.shell32.ShellExecuteW(
                None,
                "print",
                pdf_file,
                None,
                ".",
                0
            )
            if result > 32:
                print(f"Sent {pdf_file} to printer {printer_name}")
                time.sleep(3)
            else:
                print(f"Failed to send {pdf_file} to printer. Error code: {result}")
    except Exception as e:
        print(f"An error occurred while printing: {e}")
    finally:
        win32print.ClosePrinter(printer)

def run_oospsbots(missing_bots, output_dir):
    url = "https://qsbi-symbotic.us.qlikcloud.com/sense/app/ac49ed6b-f0b8-4837-b2e7-36ee0c1f19f6/sheet/FpgMpp/state/analysis/hubUrl/%2Fcatalog%3Fquick_search_filter%3DPAL%26space_filter%3D62f3ed488f9ea270826ce0c7"
    download_dir = 'C:/Users/cvickery/Downloads'
    alpha_c_template_path = "C:/Users/cvickery/PycharmProjects/PDFGENERATOR/SymBot AlphaC Traveler v16.pdf"
    one_point_zero_template_path = "C:/Users/cvickery/PycharmProjects/PDFGENERATOR/SymBot 1.0 Traveler vH.pdf"
    date_time = str(datetime.now().date())
    printer_name = "HP0A6E76.office.wmt06036-a.symbotic (HP Color LaserJet Pro M478f-9f)"

    driver = setup_driver()
    login_to_site(driver, url)
    download_file(driver)
    convert_to_csv(download_dir)
    data = extract_data(download_dir)
    driver.quit()
    os.makedirs(f"C:/Users/cvickery/TRAVELER_PDFs{date_time}", exist_ok=True)

    pdf_files = create_pdfs(data, missing_bots, alpha_c_template_path, one_point_zero_template_path)
    print_files_to_printer(pdf_files, printer_name)
