import tkinter as tk
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import json
import gspread
import threading
import requests
import json
from openpyxl import Workbook
import datetime

# ====================== WebScrapper ======================
# ==========================================================
# ====================== Version 1.0 ======================

root = tk.Tk()
# give dimension to the main window
root.geometry("700x400")
root.title("Web Scrapper")

label = tk.Label(root, text="Euronics Price Update", font=('Ariel', 24), fg="#166ABB")
label.pack()
emptyLabel = tk.Label(root, text="")
emptyLabel.pack()


# This function will be running as "Run" button clicked
def runBot():
    # running bot
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    # Creating Array to save SKUs
    sku_list = [] 
    # Getting Data from google sheet
    # This part needs a google service account
    sa = gspread.service_account(filename="service_account.json")
    sh = sa.open("GOOGLE+_WORKBOOK_NAME")
    wks = sh.worksheet("GOOGLE_SHEET_NAME")
    print("Waiting for import data from Google Sheet...")
 
    

    for x in range(1,wks.row_count):
        # Storing values (SKU)
        one_sku = wks.acell("A" + str(x)).value 
        #  Filtering blank Cells
        if type(one_sku) != type(None):
            # Adding SKU to skus array
            sku_list.append(one_sku)

    # =================================================
    # STARTING SCREEN READER
    # getting login page url
    driver.get("LOG_IN_PAGE")
    print("Typing Username")
    typeInt = driver.find_element("id", "j_username")
    typeInt.send_keys("SAMPLE_EMAIL")
    print("Typing password")
    # sampleText.insert(tk.INSERT, "Typing password \n")
    typeInt = driver.find_element("id", "j_password")
    typeInt.send_keys("SAMPLE_PASSWORD")
    # Google captcha (I tried to bypass it but didn't work)
    print("Waiting 20 Seconds to solve captha")
    time.sleep(20)

    # Creating json array
    json_array_list = []
    # empty dic (for each result)
    final_result = {}
    for sku in sku_list:
            #adding sku to the json obj
            final_result["sku"] = sku
            # logged in to the website
            print("Reloading Page")
            # chrome drive get the new url because we were at login page
            driver.get("HOME_PAGE_URL")
            # find search box
            findSku = driver.find_element("name", "text")
            findSku.send_keys(sku)

            # Wait for pop up suggetions
            print("Waiting for pop up suggetions...")
            # sampleText.insert(tk.INSERT, "Waiting for pop up suggetions... \n")
            wait = WebDriverWait(driver, 10)
            element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ui-id-2"]/div[3]/strong[1]')))
            if element:             
                # Retail Price
                findPrices = driver.find_elements(By.XPATH, value='//*[@id="ui-id-2"]/div[3]/strong[1]')
                for findPrice in findPrices:
                    # adding retail Price to json obj
                    costprice = findPrice.text.split("£")[1].replace(",", "")
                    final_result["costprice"] = float(costprice)                
                # Selling Price
                findRetailPrice = driver.find_elements(By.XPATH, value='//*[@id="ui-id-2"]/div[3]/strong[2]')
                for findPrice in findRetailPrice:
                    # adding selling price json obj
                    sellingprice = findPrice.text.split("£")[1].replace(",", "")
                    final_result["sellingprice"] = float(sellingprice)
                # addin latest prices to the json
                json_array_list.append(final_result)
                # clear final_result dec
                
                final_result = {}
                # -----------------
                # Clear search box
                findSku.clear()

    driver.quit()
    print("Creating Json File")

    with open("./result.json", "w") as _f:   
        json.dump(json_array_list, _f, indent=4)
    
    
    # ==========================
    # sending data to the server
    headers = {'Content-Type': 'application/json', 'Accept':'application/json'}

    r = requests.post("ENDPOINT", data=json.dumps(json_array_list), headers=headers)
    response = r.json()
    print(r.text)

    # saving to xlsx
    # empty array for display result in xlxs
    resultoutputlist = []
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S") + '.xlsx'

    
    # taking results from response and assigning to empty array created above
    for item in response:
        itemresult = [item['sku'], item['status']]
        resultoutputlist.append(itemresult) 

    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S") + '.xlsx'


    # create workbook and worksheet
    workbook = Workbook()
    worksheet = workbook.active

    # wrting data to the excel file
    for oneITem in resultoutputlist:
        worksheet.append([oneITem[0], oneITem[1]])
    
    workbook.save(current_time)
    print("Results saved on folder")
    time.sleep(2)
        

button = tk.Button(root, text="Run", relief=tk.FLAT, font=("Ariel", 15) , command=(threading.Thread(target=runBot).start), bg="#2BBB6A", pady=0, padx=70)
button.pack()

emptyLabel = tk.Label(root, text="")
emptyLabel.pack()





root.mainloop()