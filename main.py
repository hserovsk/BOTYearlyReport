import time
import ssl
import glob as gb
import smtplib
import pandas as pd
import selenium.common.exceptions
from openpyxl import Workbook
from dotenv import load_dotenv
import os
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains

PATH = "C:\Program Files (x86)\chromedriver.exe"

chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("prefs", {
  "download.default_directory": r"C:\Users\serowskh\PycharmProjects\RPABOTYearlyInovice\downloads"
  })
browser = webdriver.Chrome(PATH, chrome_options=chrome_options)

def configure():
    load_dotenv()

def login_to_web_page():
    browser.get("https://acme-test.uipath.com/login")
    browser.maximize_window()
    browser.find_element(By.XPATH, "//div//input[@id='email']").send_keys(f"{os.getenv('ACME_LOGIN')}")
    browser.find_element(By.XPATH, "//div//input[@id='password']").send_keys(f"{os.getenv('ACME_PASSWORD')}")
    browser.find_element(By.XPATH, "//div//button[@type='submit']").click()
    time.sleep(5)

def download_monthly_report():
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
            time.sleep(3)
        except Exception as e:
            try:
                print(f"{e} try in try except")
            except:
                print("except in try except")
def merge_excel_files():
    path = r"C:\Users\serowskh\PycharmProjects\RPABOTYearlyInovice\downloads"
    filenames = gb.glob(path + r"\xlsx")
    outputxlsx = pd.DataFrame()

    for file in filenames:
        df = pd.concat(pd.read_excel(file, sheet_name=None,), ignore_index=True, sort=False)
        outputxlsx = outputxlsx.append(df, ignore_index=True)
    outputxlsx.to_excel(r"C:\Users\serowskh\PycharmProjects\RPABOTYearlyInovice\mergedFiles\mergedfile.xlsx")

def redirect_to_work_items():
    browser.click_button("xpath://div//button[@class='btn btn-default btn-lg']")
    time.sleep(3)
def clean_folders():
    downloaded_files = gb.glob(r"C:\Users\serowskh\PycharmProjects\RPABOTYearlyInovice\downloads\*")
    merged_files = gb.glob(r"C:\Users\serowskh\PycharmProjects\RPABOTYearlyInovice\mergedFiles\*")
    for fd in downloaded_files:
        os.remove(fd)
    for fm in merged_files:
        os.remove(fm)
def send_mail():
    smtp_port = 587
    smtp_server = "smtp.gmail.com"
    email_from = f"{os.getenv('GOOGLE_MAIL')}"
    email_to = f"{os.getenv('GOOGLE_MAIL')}"
    pswd = f"{os.getenv('GOOGLE_APP_CREDENTIAL')}"
    message = "RPAHashBot"

    simple_email_context = ssl.create_default_context()

    try:
        print("Conecting to server...")
        TIE_server = smtplib.SMTP(smtp_server, smtp_port)
        TIE_server.starttls(context=simple_email_context)
        TIE_server.login(email_from, pswd)
        print("Connected to server :) ")

        print(f"Sending email to {email_to}")
        TIE_server.sendmail(email_from, email_to, message)
        print(f"Email send successfully to {email_to}")
    except Exception as e:
        print(e)
    finally:
        TIE_server.quit

def minimal_task():
    configure()
    login_to_web_page()
    download_monthly_report()
    merge_excel_files()
    send_mail()

if __name__ == "__main__":
    minimal_task()