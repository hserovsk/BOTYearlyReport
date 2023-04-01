import time
import ssl
import glob as gb
import smtplib
import pandas as pd
import logging
import pyautogui
import selenium.common.exceptions
from openpyxl import Workbook
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import os
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from datetime import date
from pywinauto.application import Application
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import autoit

# Actual Date
today = date.today()
date_text_month = today.strftime("%d %B %Y")

# Logger
logger = logging.getLogger("JobAdvertScraper")
logging.basicConfig(level=logging.INFO)
log_filename = 'logfile_{}.log'.format(date_text_month)


file_handler = logging.FileHandler(log_filename)
formatter = logging.Formatter("%(asctime)s: %(levelname)s - %(message)s")
file_handler.setFormatter(formatter)
file_handler.setLevel(logging.INFO)
logger.addHandler(file_handler)

PATH = "C:\Program Files (x86)\chromedriver.exe"

chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("prefs", {
  "download.default_directory": r"C:\Users\serowskh\PycharmProjects\RPABOTYearlyInovice\downloads"
  })
browser = webdriver.Chrome(PATH, options=chrome_options)

def configure():
    logger.info("###################### SCRAPER STARTING EXECUTION ######################")
    logger.info("Loading configuration...")
    load_dotenv()
    logger.info("Loading configuration... Success")

def login_to_web_page():
    logger.info("Loggin in to ACME website...")
    browser.get("https://acme-test.uipath.com/login")
    browser.maximize_window()
    browser.find_element(By.XPATH, "//div//input[@id='email']").send_keys(f"{os.getenv('ACME_LOGIN')}")
    browser.find_element(By.XPATH, "//div//input[@id='password']").send_keys(f"{os.getenv('ACME_PASSWORD')}")
    browser.find_element(By.XPATH, "//div//button[@type='submit']").click()
    time.sleep(5)
    logger.info("Loggin in to ACME website... Success")

def download_monthly_report():
    logger.info("Downloading monthly reports...")
    achains = ActionChains(browser)
    hover_element = browser.find_element(By.XPATH, "//button[normalize-space()='Reports']")
    achains.move_to_element(hover_element).perform()
    browser.find_element(By.XPATH, "//a[normalize-space()='Download Monthly Report']").click()
    time.sleep(3)
    browser.find_element(By.XPATH, "//div[@class='control-group form-group']//input").send_keys("IT754893")
    #browser.find_element(By.XPATH, "//div[@class='dropdown']//a[@href='https://acme-test.uipath.com/reports/download']")

    dropdown_month = browser.find_element(By.XPATH, "//select[@name='reportMonth']")
    dropdown_year = browser.find_element(By.XPATH, "//select[@name='reportYear']")
    select_year = Select(dropdown_year)
    select_year.select_by_value("2022")
    select_month = Select(dropdown_month)
    for x in range(1, 13):
        try:
            select_month.select_by_index(x)
            browser.find_element(By.XPATH, "//button[@id='buttonDownload']").click()
            time.sleep(2)
        except Exception as e:
            try:
                print(f"{e} try in try except")
            except:
                print("except in try except")
    logger.info("Downloading monthly reports... Success")
def merge_excel_files():
    logger.info("Merging excel files...")
    path = r"C:\Users\serowskh\PycharmProjects\RPABOTYearlyInovice\downloads"
    filenames = gb.glob(path + r"\xlsx")
    outputxlsx = pd.DataFrame()

    for file in filenames:
        logger.info(f"Merging file {str(file)} ...")
        df = pd.concat(pd.read_excel(file, sheet_name=None,), ignore_index=True, sort=False)
        outputxlsx = outputxlsx.append(df, ignore_index=True)
    outputxlsx.to_excel(r"C:\Users\serowskh\PycharmProjects\RPABOTYearlyInovice\mergedFiles\mergedfile.xlsx")
    logger.info("Merging excel files... Success")

def redirect_to_work_items():
    logger.info("Redirecting to work files...")
    #browser.click_button("xpath://div//button[@class='btn btn-default btn-lg']")
    time.sleep(3)
    logger.info("Redirecting to work files... Success")

def clean_folders():
    logger.info("Deleting files in folders...")
    downloaded_files = gb.glob(r"C:\Users\serowskh\PycharmProjects\RPABOTYearlyInovice\downloads\*")
    merged_files = gb.glob(r"C:\Users\serowskh\PycharmProjects\RPABOTYearlyInovice\mergedFiles\*")
    for fd in downloaded_files:
        logger.info(f"Removing file {str(fd)} in download folder")
        os.remove(fd)
    for fm in merged_files:
        logger.info(f"Removing file {str(fm)} in merged folder")
        os.remove(fm)
    logger.info("Deleting files in folders... Success")

def uploading_merged_file_test():
    browser.get("https://acme-test.uipath.com")
    achains = ActionChains(browser)
    hover_element = browser.find_element(By.XPATH, "//button[normalize-space()='Reports']")
    achains.move_to_element(hover_element).perform()
    browser.find_element(By.XPATH, "//a[normalize-space()='Upload Yearly Report']").click()
    browser.find_element(By.XPATH, "//input[@id='vendorTaxID']").send_keys("IT754893")
    dropdown_year = browser.find_element(By.XPATH, "//select[@name='reportYear']")
    select_year = Select(dropdown_year)
    select_year.select_by_value("2022")
    time.sleep(4)
    browser.find_element(By.XPATH, "//label[@for='my-file-selector']").click()
    time.sleep(1)
    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!")

    autoit.win_active("Open")
    autoit.control_send("Open", "Edit1", r"C:\Users\uu\Desktop\TestUpload.txt")
    autoit.control_send("Open", "Edit1", "{ENTER}")


    #element = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//label[@for='my-file-selector']")))

    #browser.find_element(By.XPATH, "//label[@for='my-file-selector']").click()
    #pyautogui.write(r"C:\Users\serowskh\PycharmProjects\RPABOTYearlyInovice\mergedFiles\file.txt")
    #pyautogui.press('enter')

    #browser.find_element(By.XPATH, "//label[@for='my-file-selector']").send_keys(r"C:\Users\serowskh\PycharmProjects\RPABOTYearlyInovice\mergedFiles\file.txt")
    #browser.find_element(By.ID, "my-file-selector").send_keys(os.getcwd()+ r"\mergedFiles\file.txt")
    time.sleep(2)
    #achains.send_keys("elo").send_keys(Keys.ENTER).perform()
    #app = Application(backend='uia').connect(title=btitle, timeout=100)
    #app = Application(backend='uia').connect(title="Open", timeout=100)
    #filebox = app.btitle.child_window(title="File name:", auto_id="1148", control_type="Edit").wrapper_object()
    #filebox = app.BrowserWindowName.child_window(title="File name:", auto_id="1148", control_type="Edit").wrapper_object()
    #filebox.type_keys(r'C:\Users\serowskh\PycharmProjects\RPABOTYearlyInovice\mergedFiles\mergedfile.xlsx')
def uploading_merged_file():
    browser.get("https://acme-test.uipath.com")
    achains = ActionChains(browser)
    hover_element = browser.find_element(By.XPATH, "//button[normalize-space()='Reports']")
    achains.move_to_element(hover_element).perform()
    browser.find_element(By.XPATH, "//a[normalize-space()='Upload Yearly Report']").click()
    browser.find_element(By.XPATH, "//input[@id='vendorTaxID']").send_keys("IT754893")
    dropdown_year = browser.find_element(By.XPATH, "//select[@name='reportYear']")
    select_year = Select(dropdown_year)
    select_year.select_by_value("2022")
    time.sleep(10)
    btitle = browser.title
    browser.find_element(By.XPATH, "//label[@for='my-file-selector']").click()
    time.sleep(2)
    fileInput = browser.find_element(By.name, 'Open')
    fileInput.send_keys("C:/path/to/file.jpg")
    #app = Application(backend='uia').connect(title=btitle, timeout=100)
    #app = Application(backend='uia').connect(title="Open", timeout=100)
    #filebox = app.btitle.child_window(title="File name:", auto_id="1148", control_type="Edit").wrapper_object()
    #filebox = app.BrowserWindowName.child_window(title="File name:", auto_id="1148", control_type="Edit").wrapper_object()
    #filebox.type_keys(r'C:\Users\serowskh\PycharmProjects\RPABOTYearlyInovice\mergedFiles\mergedfile.xlsx')

def send_mail():
    smtp_port = 587
    smtp_server = "smtp.gmail.com"


    email_from = f"{os.getenv('GOOGLE_MAIL')}"
    email_to = f"{os.getenv('GOOGLE_MAIL')}"
    pswd = f"{os.getenv('GOOGLE_APP_CREDENTIAL')}"

    message = MIMEMultipart("alternative")
    message["Subject"] = "RPABOT Run"
    #message["From"] = f"{os.getenv('GOOGLE_MAIL')}"
    #message["To"] = f"{os.getenv('GOOGLE_MAIL')}"
    # write the text/plain part
    text = f"""\
    Hi,
    The bot run was successful
    Best Regards,
    RPA BOT
    """


    # write the HTML part
    html = """\
    <html>
      <body>
        <p>Hi,<br>
           The bot run was successful</p>
        <p>Best Regards,</p>
        <p> RPA BOT </p>
      </body>
    </html>
    """

    # convert both parts to MIMEText objects and add them to the MIMEMultipart message
    part1 = MIMEText(text, "plain")
    part2 = MIMEText(html, "html")
    message.attach(part1)
    message.attach(part2)

    simple_email_context = ssl.create_default_context()

def send_exception_mail(error):
    smtp_port = 587
    smtp_server = "smtp.gmail.com"
    email_from = f"{os.getenv('GOOGLE_MAIL')}"
    email_to = f"{os.getenv('GOOGLE_MAIL')}"
    pswd = f"{os.getenv('GOOGLE_APP_CREDENTIAL')}"

    message = MIMEMultipart("alternative")
    message["Subject"] = "RPA BOT Run"
    # message["From"] = f"{os.getenv('GOOGLE_MAIL')}"
    # message["To"] = f"{os.getenv('GOOGLE_MAIL')}"
    # write the text/plain part
    text = f"""\
    Hi,
    The bot run had following error:
    Best Regards,
    RPA BOT
    """

    # write the HTML part
    html = f"""\
    <html>
        <body>
        <p>Hi,<br>
            The bot run had following error: </p><br>
            <p> {error} <p>
        <p>Best Regards,</p>
        <p> RPA BOT </p>
        </body>
    </html>
    """

    # convert both parts to MIMEText objects and add them to the MIMEMultipart message
    part1 = MIMEText(text, "plain")
    part2 = MIMEText(html, "html")
    message.attach(part1)
    message.attach(part2)

    simple_email_context = ssl.create_default_context()

    try:
        print("Conecting to server...")
        TIE_server = smtplib.SMTP(smtp_server, smtp_port)
        TIE_server.starttls(context=simple_email_context)
        TIE_server.login(email_from, pswd)
        print("Connected to server :) ")

        print(f"Sending exception email to {email_to}")
        TIE_server.sendmail(email_from, email_to, message.as_string())
        print(f"Email send exception successfully to {email_to}")
    except Exception as e:
        print(e)
    finally:
        TIE_server.quit
    logger.info("###################### SCRAPER EXECUTION ENDED ######################")

def minimal_task():
    configure()
    login_to_web_page()
    uploading_merged_file_test()
    """    try:
        configure()
        login_to_web_page()
        #download_monthly_report()
        #merge_excel_files()
        #uploading_merged_file()
        uploading_merged_file_test()
        #clean_folders()
        #send_mail()
    except Exception as error:
        clean_folders()
        send_exception_mail(error)"""


if __name__ == "__main__":
    minimal_task()