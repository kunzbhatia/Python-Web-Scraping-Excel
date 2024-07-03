import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
import os

drvPath = "/Users/kunalbhatia/Downloads/chromedriver-mac-arm64/chromedriver"

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--incognito")

# Initialize WebDriver with options
browser = webdriver.Chrome(options=chrome_options)
browser.get("https://www.morningstar.ca/ca/report/etf/performance.aspx?t=0P0001HWR9&lang=en-CA")

wait = WebDriverWait(browser, 30)

try:
    # Wait for prices element
    prices = wait.until(EC.visibility_of_element_located((By.ID, "message-box-price")))
    price = prices.text.strip()

    # Determine price trend
    try:
        PriceUp = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//i[contains(@ng-if, 'chartData.isPriceUp')]")))
        priceTrend = "Up" if len(PriceUp) > 0 else "Down"
    except TimeoutException:
        priceTrend = "Down"  # Set default value if elements are not found within timeout

    # Wait for change price and percent elements
    changePriceNPercent = wait.until(EC.visibility_of_element_located((By.ID, "message-box-percentage"))).text
    a = changePriceNPercent.replace("%", "").split("|")
    changePrice = a[0].strip()
    changePercent = a[1].strip()

    # Try to find ssDateTime element
    try:
        ssDateTime = wait.until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@ng-if, 'vm.footerShowLastPrice')]"))).text
        dts = ssDateTime.replace("| Last Price updated as of ", "").replace(",", "")
    except TimeoutException:
        dts = "N/A"  # Set default value if elements are not found within timeout

    # Create a DataFrame to store data
    data = {
        "Date-Time": [dts],
        "Price": [price],
        "Trend": [priceTrend],
        "Change $": [changePrice],
        "Change %": [changePercent]
    }
    df = pd.DataFrame(data)

    # Specify directory path for Excel file
    directory = "/Users/kunalbhatia/Documents"  # Example writable directory path
    fname = os.path.join(directory, "ETF.xlsx")

    # Check if directory exists; if not, create it
    if not os.path.exists(directory):
        os.makedirs(directory)

    # Check if Excel file exists; if not, create it
    if not os.path.exists(fname):
        df.to_excel(fname, index=False)
    else:
        with pd.ExcelWriter(fname, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name='ETF', index=False, header=not writer.sheets['ETF'].exists())

    print("Data written to ETF.xlsx successfully.")

except TimeoutException as e:
    print(f"Timeout occurred: {str(e)}")

except Exception as e:
    print(f"An error occurred: {str(e)}")

finally:
    browser.quit()