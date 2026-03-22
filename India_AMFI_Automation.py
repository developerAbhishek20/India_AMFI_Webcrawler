import os
import time
import glob
import pandas as pd
import win32com.client as win32

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# -----------------------------
# USER SETTINGS
# -----------------------------

FROM_DATE = "03/11/2026"
TO_DATE = "03/12/2026"

DOWNLOAD_FOLDER = r"C:\Users\asalvi\OneDrive - MORNINGSTAR INC\Documents\India Fund SIF Automation"

EMAIL_RECEIVER = "abhishek.salvi@morningstar.com"

SUBJECT_PREFIX = "India SIF NAV Data"

# -----------------------------
# WAIT FOR DOWNLOAD
# -----------------------------

def wait_for_download(folder, timeout=60):

    start = time.time()

    while True:

        files = os.listdir(folder)

        downloading = [f for f in files if f.endswith(".crdownload")]

        if not downloading:
            return True

        if time.time() - start > timeout:
            return False

        time.sleep(1)

# -----------------------------
# CLEAN OLD RAW FILES
# -----------------------------

for f in glob.glob(os.path.join(DOWNLOAD_FOLDER, "NAV_*.xlsx")):
    os.remove(f)

# -----------------------------
# CHROME SETTINGS
# -----------------------------

options = webdriver.ChromeOptions()

prefs = {
    "download.default_directory": DOWNLOAD_FOLDER,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True
}

options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=options)

driver.maximize_window()

wait = WebDriverWait(driver, 20)
quick_wait = WebDriverWait(driver, 5)

# -----------------------------
# OPEN WEBSITE
# -----------------------------

driver.get("https://www.amfiindia.com/sif/latest-nav/nav-history")

historical_nav = wait.until(
    EC.element_to_be_clickable((By.XPATH,"//span[contains(text(),'Historical NAV for a period')]"))
)

driver.execute_script("arguments[0].click();", historical_nav)

print("Historical NAV selected")

# -----------------------------
# GET ALL FUNDS
# -----------------------------

fund_dropdown = wait.until(
    EC.element_to_be_clickable((By.XPATH,"//input[@placeholder='Select']"))
)

fund_dropdown.click()

fund_elements = wait.until(
    EC.presence_of_all_elements_located((By.XPATH,"//li[contains(@class,'MuiAutocomplete-option')]"))
)

fund_list = [f.text for f in fund_elements]

driver.find_element(By.TAG_NAME,"body").click()

print("Funds Found:", fund_list)

# -----------------------------
# LOOP FUNDS
# -----------------------------

for fund in fund_list:

    print("\nProcessing Fund:", fund)

    fund_dropdown = wait.until(
        EC.element_to_be_clickable((By.XPATH,"//input[@placeholder='Select']"))
    )

    fund_dropdown.click()

    fund_option = wait.until(
        EC.element_to_be_clickable((By.XPATH,f"//li[normalize-space()='{fund}']"))
    )

    driver.execute_script("arguments[0].click();", fund_option)

    strategy_dropdown = wait.until(
        EC.element_to_be_clickable((By.XPATH,"(//input[@placeholder='Select'])[2]"))
    )

    strategy_dropdown.click()

    strategy_elements = wait.until(
        EC.presence_of_all_elements_located((By.XPATH,"//li[contains(@class,'MuiAutocomplete-option')]"))
    )

    strategy_list = [s.text for s in strategy_elements]

    driver.find_element(By.TAG_NAME,"body").click()

    print("Strategies:", strategy_list)

    # -----------------------------
    # LOOP STRATEGIES
    # -----------------------------

    for strategy in strategy_list:

        print("Strategy:", strategy)

        strategy_dropdown = wait.until(
            EC.element_to_be_clickable((By.XPATH,"(//input[@placeholder='Select'])[2]"))
        )

        strategy_dropdown.click()

        strategy_option = wait.until(
            EC.element_to_be_clickable((By.XPATH,f"//li[normalize-space()='{strategy}']"))
        )

        driver.execute_script("arguments[0].click();", strategy_option)

        # -----------------------------
        # ENTER DATES
        # -----------------------------

        from_date = wait.until(
            EC.element_to_be_clickable((By.XPATH,"(//input[contains(@class,'MuiInputBase-input')])[3]"))
        )

        to_date = wait.until(
            EC.element_to_be_clickable((By.XPATH,"(//input[contains(@class,'MuiInputBase-input')])[4]"))
        )

        from_date.clear()
        from_date.send_keys(FROM_DATE)

        to_date.clear()
        to_date.send_keys(TO_DATE)

        # -----------------------------
        # CLICK GO
        # -----------------------------

        go_button = wait.until(
            EC.presence_of_element_located((By.XPATH,"/html/body/div/div[2]/div/div[2]/div[2]/button"))
        )

        driver.execute_script("arguments[0].click();", go_button)

        print("GO clicked")

        try:

            download_btn = quick_wait.until(
                EC.element_to_be_clickable((By.XPATH,"//button[@aria-label='Download Excel']"))
            )

            download_btn.click()

            print("Downloading Excel...")

            wait_for_download(DOWNLOAD_FOLDER)

            time.sleep(2)

        except TimeoutException:

            print("No data available")

driver.quit()

print("\nAll downloads completed")

# ======================================================
# CONVERT FILES + SEND EMAIL
# ======================================================

print("\nProcessing downloaded files...")

excel_files = glob.glob(os.path.join(DOWNLOAD_FOLDER,"NAV_*.xlsx"))

outlook = win32.Dispatch("Outlook.Application")

mail_count = 0

for file in excel_files:

    try:

        raw = pd.read_excel(file,header=None)

        if len(raw) < 6:
            print("Skipping incomplete:",file)
            continue

        header_rows = raw.iloc[0:5,0].astype(str).tolist()

        investment_name = ""
        strategy_name = ""

        for text in header_rows:

            if "SIF" in text:
                investment_name = text.strip()

            if "Fund" in text:
                strategy_name = text.strip()

        nav = raw.iloc[5:,0].reset_index(drop=True)
        date = raw.iloc[5:,3].reset_index(drop=True)

        df = pd.DataFrame()

        df["Investment Name"] = [investment_name]*len(nav)
        df["Investment Strategy Name"] = [strategy_name]*len(nav)
        df["Price Date"] = pd.to_datetime(date).dt.strftime("%Y-%m-%d")
        df["NAV"] = nav

        if "Growth" in strategy_name:
            strategy_short = "Growth"
        elif "IDCW" in strategy_name:
            strategy_short = "IDCW"
        elif "Direct" in strategy_name:
            strategy_short = "Direct"
        elif "Regular" in strategy_name:
            strategy_short = "Regular"
        else:
            strategy_short = "Strategy"

        new_name = f"India SIF_{investment_name}_{strategy_short}.xlsx"

        new_path = os.path.join(DOWNLOAD_FOLDER,new_name)

        df.to_excel(new_path,index=False)

        os.remove(file)

        print("Converted:",new_name)

        # -----------------------------
        # SEND EMAIL
        # -----------------------------

        mail = outlook.CreateItem(0)

        mail.To = EMAIL_RECEIVER
        mail.Subject = f"{SUBJECT_PREFIX} - {investment_name}"

        mail.Body = f"""
Hello,

Please find attached NAV data file for:

{investment_name}

Regards
Morningstar India Pvt. Ltd.
"""

        mail.Attachments.Add(new_path)

        mail.Send()

        mail_count += 1

        print("Mail sent:", investment_name,"| Total:",mail_count)

        # prevent Outlook throttling
        time.sleep(5)

    except Exception as e:

        print("Error processing:",file)
        print(e)

print("\nAutomation Completed Successfully")
