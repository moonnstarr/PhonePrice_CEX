import streamlit as st
import pandas as pd
import xlwings as xw
from selenium import webdriver
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pythoncom

pythoncom.CoInitialize()

# STREAMLIT PAGE CONFIGURATION -----------------------------------------------------------------------------------------
st.set_page_config(page_title='CEX BOT', layout='wide')
st.subheader("CEX Price Updater")


run = st.button("update")

if run:
    # Read supplier price quotation from Google Sheet
    gsheet_id = '1AUqkl6Nd9gRUooxtWiVrKUnvV_aXNQ2sAbCdvbBcxnQ'
    sheet_name = 'price'
    gsheet_url = f'https://docs.google.com/spreadsheets/d/{gsheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}'

    phone_list = pd.read_csv(gsheet_url)

    # Scrape CEX website to get the product prices
    def retrieve_data():
        # Initialize the COM subsystem
        # pythoncom.CoInitialize()

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
            user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/110.0"

            options = webdriver.ChromeOptions()
            # options.headless = True
            options.add_argument('--headless')
            options.add_argument(f'user-agent={user_agent}')
            options.add_argument("--window-size=1920,1080")
            options.add_argument('--ignore-certificate-errors')
            options.add_argument('--allow-running-insecure-content')
            options.add_argument("--disable-extensions")
            options.add_argument("--proxy-server='direct://'")
            options.add_argument("--proxy-bypass-list=*")
            options.add_argument("--start-maximized")
            options.add_argument('--disable-gpu')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--no-sandbox')
            # driver = webdriver.Chrome(executable_path="chromedriver.exe", options=options)
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

            first_row = df[:1]

            main_df = main_df.append(first_row, ignore_index=True)

        return main_df


    # Clean data and arrange it to export as Excel
    def rearrange_data():

        final_df = pd.concat([phone_list, retrieve_data()], axis=1)

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
        st.table(final_df)
        return final_df

    # Save to Excel
    def save_to_excel():
        with xw.App(visible=False) as app:
            wb = app.books.open('Template.xlsx')
            wb.sheets['Sheet1'].range('A3').value = rearrange_data()
            wb.sheets['Sheet1'].range('3:3').delete()
            wb.save("CEX_Product_Price.xlsx")

    save_to_excel()

    # st.download_button(label="Download",
    #                    data=final_file,
    #                    file_name="cex_product_price.xlsx")
else:
    st.write("Please click the update button to run")