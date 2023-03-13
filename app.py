import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import scrolledtext
from tkinter import ttk
import pandas as pd
import xlwings as xw
from selenium import webdriver
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pythoncom


class CexBot:

    def __init__(self):

        self.root = tk.Tk()

        self.Label = tk.Label(self.root, text="CEX Price Updater", font=('Arial', 18))
        self.Label.pack(pady=10, padx=10)

        # input file path variable
        self.label_file_path = tk.Label(self.root, text='')

        self.select_button = tk.Button(self.root, text="Select File", font=("Arial", 18),
                                       command=self.select_file, bg='lightblue')
        self.select_button.pack(padx=10, pady=10)

        # output file save location variable
        self.output_save_location = tk.Label(self.root, text='')

        self.update_button = tk.Button(self.root, text="Update & Save", font=("Arial", 18),
                                       command=self.update_and_save, bg='lightblue')
        self.update_button.pack(padx=10, pady=10)

        # Progress bar
        self.my_progress = ttk.Progressbar(self.root, orient='horizontal',
                                           length=300, mode='indeterminate')
        self.my_progress.pack(pady=20)

        # Print results
        self.results_box = scrolledtext.ScrolledText(self.root, width=40, height=10)
        self.results_box.pack(padx=10, pady=10)

        self.root.mainloop()

    def select_file(self):
        file_path = filedialog.askopenfilename(title='Select Input File')
                                               # filetypes=(("xlsx files", "*.xlsx"), ("All Files", "*.*"))
                                               # )
        self.label_file_path["text"] = file_path
        return None

    def update_and_save(self):

        # progress bar starting
        self.my_progress.start(10)
        # self.root.update_idletasks()

        # Save location
        save_folder = filedialog.askdirectory(title="Select a folder to Save")
        self.output_save_location["text"] = save_folder
        save_location = self.output_save_location["text"]
        save_folder_location = r"{}".format(save_location)

        # Read the phone model list to get the price
        excel_file_path = self.label_file_path["text"]

        try:
            excel_file_name = r"{}".format(excel_file_path)
            phone_list = pd.read_csv(excel_file_name)
            # phone_list = pd.read_csv("stock.csv")
            print(phone_list)

            # Printing results on tkinter window
            uploaded_models = phone_list["Name"].to_list()
            self.results_box.insert(tk.END, uploaded_models)
        #     -----------------------------------------------------------------------------------------------------

        except ValueError:
            tk.messagebox.showerror("Information", "The file you have entered is invalid")
            self.my_progress.stop()
            return None
        except FileNotFoundError:
            tk.messagebox.showerror("Information", f"No such file as {excel_file_path}")
            self.my_progress.stop()
            return None

        def retrieve_data():
            # Initialize the COM subsystem
            pythoncom.CoInitialize()

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

                # headless chrome option
                user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/110.0"

                options = webdriver.ChromeOptions()
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
                    test = model.append(models.find_element(By.XPATH, ".//span[@class='ais-Highlight']").text)
                    print(test) # ========================================================================================
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
                print(first_row)

                main_df = main_df.append(first_row, ignore_index=True)
                print(main_df)

                # display the DataFrame in the scrolled text widget
                # self.results_box.insert(tk.END, main_df.to_string())

            def rearrange_data():

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
                # Close the web browser and return
                driver.quit()


                return final_df

            # export the file into preformatted excel
            def save_to_excel():
                with xw.App(visible=False) as app:
                    wb = app.books.open('Template.xlsx')
                    wb.sheets['Sheet1'].range('A3').value = rearrange_data()
                    wb.sheets['Sheet1'].range('3:3').delete()
                    wb.save(f"{save_folder_location}/CEX_Product_Price.xlsx")

            try:
                retrieve_data()
                rearrange_data()
                save_to_excel()
                self.my_progress.stop()
            except:
                tk.messagebox.showerror("Information", "Error occured during Excel conversion!")
                self.my_progress.stop()
                return None

            tk.messagebox.showinfo("Information", 'Updated & File Saved Successfully!')

            # Uninitialize the COM subsystem
            pythoncom.CoUninitialize()

CexBot()

