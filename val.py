from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd

# Functions

# Open Browser
def open_browser(PATH):
    driver = webdriver.Chrome(PATH)
    driver.set_window_size(1980,1080)
    return driver

# Close Browser
def quit_browser(driver):
    driver.quit()
    
# Recorrer las filas y obtener los datos de las celdas
def rows_loop(driver, companies:list):
    table = driver.find_element(By.CLASS_NAME, 'table-light')
    table_rows = table.find_elements(By.TAG_NAME, 'tr')

    # loop de las filas de la tabla
    for row in table_rows:
        companies_info = []
        if row.get_attribute("align") == "center":
            continue
        actual_row = row.find_elements(By.TAG_NAME, "td") # Encontrar las celdas de la fila actuaL
    
        for cell in actual_row:
            companies_info.append(cell.find_element(By.TAG_NAME, "a").text) # Appende del valor de la celda
        companies.append(companies_info)

# Run

# Open Browser
PATH = '/home/juanjm/Documents/ChromeWebDriver/chromedriver' 
driver = open_browser(PATH)


# Variables
sections = (111, 121, 161)
jump = 20
limit = 400 # real number of companies: 8521
companies_overview = []
companies_valuation = []
companies_financial = []

overview = ("No.", "Ticker", "Company", "Sector", "Industry", "Country", "Market_Cap", "P/E", "Price", "Change", "Volume")
valuation = ("No.", "Ticker", "Market_Cap", "P/E", "Fwd_P/E", "PEG", "P/S", "P/B", "P/C", "P/FCF",
                     "EPS_this_Y", "EPS_next_Y", "EPS_past_5Y", "EPS_next_5Y", "Sales_past_5Y", "Price", "Change", "Volume")
financial = ("No.", "Ticker", "Market_Cap", "Dividend", "ROA", "ROE", "ROI", "Curr_R", "Quick_R", "LTDebt/Eq", "Debt/Eq", "Gross_M",
             "Oper_M", "Profit_M", "Earnings", "Price", "Change", "Volume")

# Go to finviz
URL = 'https://finviz.com/'

start = time.time()

# Principal loop
for section in sections:
    counter = 1
    main_url = URL+"screener.ashx?v="+str(section)+"&r="
    if section == sections[0]:
        while counter <= limit:
            driver.get(main_url+str(counter))
            rows_loop(driver, companies_overview)
            counter += jump
    elif section == sections[1]:
        while counter <= limit:
            driver.get(main_url+str(counter))
            rows_loop(driver, companies_valuation)
            counter += jump
    else:
        while counter <= limit:
            driver.get(main_url+str(counter))
            rows_loop(driver, companies_financial)
            counter += jump
    
# Quit Browser
quit_browser(driver)

end = time.time()
print(end - start)

# Create DataFrames
df_ov = pd.DataFrame(companies_overview, columns=overview).to_excel("/home/juanjm/Documents/ov.xlsx")
df_va = pd.DataFrame(companies_valuation, columns=valuation).to_excel("/home/juanjm/Documents/va.xlsx")
df_fi = pd.DataFrame(companies_financial, columns=financial).to_excel("/home/juanjm/Documents/fi.xlsx")