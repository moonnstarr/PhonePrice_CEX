import streamlit as st
import pandas as pd
import xlwings as xw
import openpyxl
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pythoncom

# Initialize the COM subsystem
pythoncom.CoInitialize()

# Read the phone model list to get the price
phone_list = pd.read_csv("stock.csv")
st.table(phone_list)

main_df = pd.DataFrame(columns=['model',
                                'category',
                                'grade',
                                'wesell',
                                'webuy_cash',
                                'webuy_voucher',
                                'low_margin %',
                                'mid_margin %',
                                'high_margin %'])

for i, row in phone_list.iterrows():
    website = f"https://uk.webuy.com/search?stext=+{row[0]}&Grade=C"

    # Selenium and Chrome webscraping configs
    options = webdriver.ChromeOptions()
    chrome_options = ChromeOptions()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    options.add_experimental_option("detach", False)
    driver = webdriver.Chrome(options=options)
    driver.get(website)

    # Wait for the cookie acceptance button to be visible
    wait = WebDriverWait(driver, 10)
    cookie_button = wait.until(EC.visibility_of_element_located((By.ID, "onetrust-accept-btn-handler")))

    # Click the acceptance button
    cookie_button.click()

    # Extracting the data from the website
    phone_models = driver.find_elements(By.XPATH, "//div[@class='desc']")

    # List to append the extracted data
    model = []
    device = []
    grade = []
    wesell = []
    webuy_cash = []
    webuy_voucher = []
    low_margin = []
    mid_margin = []
    high_margin = []

    # Looping through each phone search results to get the values
    for models in phone_models:
        model.append(models.find_element(By.XPATH, ".//span[@class='ais-Highlight']").text)
        device.append(models.find_element(By.XPATH, "./p").text)
        grade.append('C')
        wesell.append(models.find_element(By.XPATH, ".//div[starts-with(@class,'priceTxt') and starts-with(text(),'WeSell for')]").text)
        webuy_cash.append(models.find_element(By.XPATH, ".//div[starts-with(@class,'priceTxt') and starts-with(text(),'WeBuy for cash')]").text)
        webuy_voucher.append(models.find_element(By.XPATH, ".//div[starts-with(@class,'priceTxt') and starts-with(text(),'WeBuy for voucher')]").text)
        low_margin.append(0.20)
        mid_margin.append(0.25)
        high_margin.append(0.30)

    # Rearranging the data into a data frame
    df = pd.DataFrame({'model': model,
                       'category': device,
                       'grade': grade,
                       'wesell': wesell,
                       'webuy_cash': webuy_cash,
                       'webuy_voucher': webuy_voucher,
                       'low_margin %': low_margin,
                       'mid_margin %': mid_margin,
                       'high_margin %': high_margin})

    first_row = df.iloc[0]

    main_df = main_df.append(first_row, ignore_index=True)
    final_df = pd.concat([phone_list, main_df], axis=1)

    # Converting text into numbers and calculations
    final_df['wesell'] = final_df['wesell'].str.replace('WeSell for £', '').astype(float)
    final_df['webuy_cash'] = final_df['webuy_cash'].str.replace('WeBuy for cash £', '').astype(float)
    final_df['webuy_voucher'] = final_df['webuy_voucher'].str.replace('WeBuy for voucher £', '').astype(float)
    final_df['low_margin_cost'] = final_df['webuy_cash'] * 0.80
    final_df['mid_margin_cost'] = final_df['webuy_cash'] * 0.75
    final_df['high_margin_cost'] = final_df['webuy_cash'] * 0.70
    final_df['low_margin %'] = final_df['low_margin %'].map('{:.1%}'.format)
    final_df['mid_margin %'] = final_df['mid_margin %'].map('{:.1%}'.format)
    final_df['high_margin %'] = final_df['high_margin %'].map('{:.1%}'.format)

    # Rename columns
    final_df = final_df.rename(columns={'Name': 'Supplier_Name',
                                        'Qty': 'Supplier_Qty',
                                        'Cost': 'Supplier_Cost'})
    # Close the web browser
    driver.quit()

# export the file into preformatted excel
def save_to_excel():
    with xw.App(visible=False) as app:
        wb = app.books.open('Template.xlsx')
        wb.sheets['Sheet1'].range('A3').value = final_df
        wb.sheets['Sheet1'].range('3:3').delete()
        wb.save("output.xlsx")

save_to_excel()

st.table(final_df)

# Uninitialize the COM subsystem
pythoncom.CoUninitialize()

