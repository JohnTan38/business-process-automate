from selenium import webdriver # 1 login 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import time
import os
import subprocess
import shutil
import argparse
import sys

def create_driver():
    options = webdriver.ChromeOptions()
    service = Service(log_path=os.devnull)
    if hasattr(subprocess, "CREATE_NO_WINDOW"):
        service.creationflags = subprocess.CREATE_NO_WINDOW
    return webdriver.Chrome(service=service, options=options)

def login_esker(driver):
    #driver = webdriver.Chrome()
    driver.get("https://az3.ondemand.e***r.com/ondemand/webaccess/asf/home.aspx")
    driver.maximize_window()
    time.sleep(1)

    driver.find_element(By.XPATH, '//*[@id="ctl03_tbUser"]').send_keys("john.tan@abc.com")
    driver.find_element(By.XPATH, '//*[@id="ctl03_tbPassword"]').send_keys("YOUR_PASSWORD")
    driver.find_element(By.XPATH, '//*[@id="ctl03_btnSubmitLogin"]').click()
    time.sleep(2)
    #return driver
def hover(driver, x_path):
    elem_to_hover = driver.find_element(By.XPATH, x_path)
    hover = ActionChains(driver).move_to_element(elem_to_hover)
    hover.perform()

#driver = login_esker()
def hover_arrow(driver, x_path):
    elem_to_hover = driver.find_element(By.XPATH, x_path)
    hover = ActionChains(driver).move_to_element(elem_to_hover)
    hover.perform()

"""
x_path_hover = '//*[@id="mainMenuBar"]/td/table/tbody/tr/td[36]/a/div' #arrow
hover_arrow(driver, x_path_hover)

try:
    #drop_down=driver.find_element(By.XPATH, '//*[@id="DOCUMENT_TAB_100872215"]/a/div[2]').click()
    tables=driver.find_element(By.XPATH, '//*[@id="CUSTOMTABLE_TAB_100872176"]')
    tables.click()
    time.sleep(1)
except Exception as e:
    print(e) #VENDOR INVOICES (SUMMARY) #TABLES
time.sleep(1)
"""

import pyautogui  #20251002 great! #2
from pathlib import Path
import win32com.client  #esker vendor update Great 20241129!
import time #create inbox subfolder, download attachments, move email to subfolder.
import re
import datetime as dt
import pandas as pd
import numpy as np
import openpyxl
from datetime import datetime
import json
"""
email = 'john.tan@sh-cogent.com.sg'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.Folders(email).Folders("Inbox")

date_time = dt.datetime.now()
lastTwoDaysDateTime = dt.datetime.now() - dt.timedelta(days = 2)
newDay = dt.datetime.now().strftime('%Y%m%d')
path_vendor_update = r"C:/Users/john.tan/Documents/power_apps_esker_vendor/esker_vendor_update/"

sub_folder1 = inbox.Folders['esker_vendor']
try:
    myfolder = sub_folder1.Folders[newDay] #check if fldr exists, else create
    #print('folder exists')
except:
    sub_folder1.Folders.Add(newDay)
    #print('subfolder created')
dest = sub_folder1.Folders[newDay]

i=0
#messages = inbox.Items
messages = sub_folder1.Items
lastTwoDaysMessages = messages.Restrict("[ReceivedTime] >= '" +lastTwoDaysDateTime.strftime('%d/%m/%Y %H:%M %p')+"'") #AND "urn:schemas:httpmail:subject"Demurrage" & "'Bill of Lading'")
for message in lastTwoDaysMessages:
        if (("ESKER VENDOR EMAIL") or ("esker vendor email")) in message.Subject:
            
            for attachment in message.Attachments:
                      #print(attachment.FileName)
                      try:
                            attachment.SaveAsFile(os.path.join(path_vendor_update, str(attachment).split('.xlsx')[0]+'.xlsx'))#'_'+str(i)+
                            i +=1
                      except Exception as e: 
                            print(e)

path_vendor_update = r"C:/Users/john.tan/Documents/power_apps_esker_vendor/esker_vendor_update/"
paths = [(p.stat().st_mtime, p) for p in Path(path_vendor_update).iterdir() if p.suffix == ".xlsx"] #Save .xlsx files paths and modification time into paths
paths = sorted(paths, key=lambda x: x[0], reverse=True) # Sort by modification time
last_path = paths[0][1] ## Get the last modified file
#excel_vendor_update = 'vendor_update.xlsx'
try:
    vendor_update = pd.read_excel(last_path, sheet_name='vendors', engine='openpyxl')
except FileNotFoundError:
    print(f"Error: File '{last_path}' not found in path '{path_vendor_update}'.")
    time.sleep(10)
    exit()
"""
def format_vendor_data(df_vendor_update):
    """
    Formats the vendor data in the DataFrame, ensuring:
    1. 'vendor_number' and 'postal_code' are numeric
    2. 'postal_code' has 6 digits
    3. Empty or missing values in 'street', 'city', 'postal_code', and 'country_code' are replaced with empty strings
    Args:
        df_vendor_update (pd.DataFrame): The DataFrame containing vendor data.
    Returns:
        pd.DataFrame: The formatted DataFrame.
    """
    # Convert 'vendor_number' and 'postal_code' to numeric
    df_vendor_update['vendor_number'] = pd.to_numeric(df_vendor_update['vendor_number'], errors='coerce')
    #df_vendor_update['postal_code'] = pd.to_numeric(df_vendor_update['postal_code'], errors='coerce')
    
    # Ensure 'postal_code' has 6 digits
    #df_vendor_update['postal_code'] = df_vendor_update['postal_code'].astype(str).str.zfill(6)
    df_vendor_update.fillna('', inplace=True) # Replace empty or missing values with empty strings
    return df_vendor_update

def load_vendor_updates_from_json(json_path: Path) -> pd.DataFrame:
    """Read a vendor summary JSON file and return a DataFrame with one row."""
    with open(json_path, encoding="utf-8") as source:
        payload = json.load(source)

    pattern = re.compile(r"([A-Za-z0-9]+)\s+(\d+)\s+(.+)")

    def build_frame(company_code: str, vendor_number: str, vendor_name: str) -> pd.DataFrame:
        return pd.DataFrame(
            {
                "company_code": [company_code.strip()],
                "vendor_number": [vendor_number.strip()],
                "vendor_name": [vendor_name.strip()],
            }
        )

    triplet = payload.get("triplet")
    if triplet:
        if isinstance(triplet, dict):
            code = triplet.get("company_code") or triplet.get("companyCode")
            number = triplet.get("vendor_number") or triplet.get("vendorNumber")
            name = triplet.get("vendor_name") or triplet.get("vendorName")
            if code and number and name:
                return build_frame(code, str(number), name)
        elif isinstance(triplet, (list, tuple)) and len(triplet) == 3:
            code, number, name = triplet
            if code and number and name:
                return build_frame(str(code), str(number), str(name))
        elif isinstance(triplet, str):
            match = pattern.search(triplet)
            if match:
                return build_frame(match.group(1), match.group(2), match.group(3))

    text_candidates = []
    for key in ("body", "body_text", "bodyText", "bodyPreview", "subject"):
        value = payload.get(key)
        if value:
            text_candidates.extend(
                seg.strip() for seg in re.split(r"[\r\n]+", str(value)) if seg.strip()
            )

    for line in text_candidates:
        match = pattern.search(line)
        if match:
            return build_frame(match.group(1), match.group(2), match.group(3))

    raise ValueError(
        f"Unable to parse vendor details in {json_path.name}; inspected keys "
        f"{[key for key in ('triplet', 'body', 'body_text', 'bodyText', 'bodyPreview', 'subject')]}"
    )


def load_latest_vendor_json(json_dir: Path) -> pd.DataFrame:
    """
    Return the newest parsable vendor JSON payload inside the provided directory.

    The environment variable `ESKER_VENDOR_JSON_PATTERN` can be used to narrow the
    filenames (e.g. ``vendor_*.json``). Files that do not match the expected schema
    are skipped rather than causing the workflow to abort.
    """
    pattern = os.getenv("ESKER_VENDOR_JSON_PATTERN")
    candidates = json_dir.glob(pattern) if pattern else json_dir.glob("*.json")
    json_files = sorted(
        (path for path in candidates if path.is_file()),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    if not json_files:
        raise FileNotFoundError(f"No JSON files found in {json_dir}")

    errors: list[str] = []
    for json_path in json_files:
        try:
            return load_vendor_updates_from_json(json_path)
        except ValueError as exc:
            errors.append(f"{json_path.name}: {exc}")

    raise ValueError(
        "Unable to parse any JSON payloads in "
        f"{json_dir}. Encountered: {', '.join(errors)}"
    )

def create_log_file(path):
    """
    Checks if a log file exists at the specified path.
    If not, creates a new one with the current date and time.
    """
    os.makedirs(path, exist_ok=True)
    filename = f"log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
    full_path = os.path.join(path, filename)

    if not os.path.exists(full_path):
        with open(full_path, 'w') as f:
            f.write("")  # Create an empty file

    return full_path
"""
path_vendor_update = r"C:/Users/john.tan/Documents/power_apps_esker_vendor/esker_vendor_update/"
log_file = create_log_file(path_vendor_update+'Log/') # Create the log file if it doesn't exist

list_company_code =[]
list_vendor_number =[]
"""
def start_time():
    start_time=datetime.now()
    return start_time

def vendor_update_process(driver, df_vendor_update):
    """
    Processes vendor updates by iterating through the provided DataFrame.
    For each vendor, it performs a series of actions to update vendor information.
    Args:
        df_vendor_update (pd.DataFrame): The DataFrame containing vendor data.
    """
    for i in range(len(df_vendor_update)):
        print(f"company_code {df_vendor_update.loc[i, 'company_code']}")

        pyautogui.moveTo(35,350, duration=2) #move cursor to extreme left side
        time.sleep(1)
        pyautogui.click()
        #pyautogui.typewrite('S2P - Vendors')
        
        try:
            #s2p_vendors=driver.find_element(By.XPATH, '//*[@id="ViewSelector"]/div/div/div/div[1]/div[1]/span')
            #s2p_vendors.click()
            time.sleep(1)
        except Exception as e:
            time.sleep(0.5)
        
        try:                
                pyautogui.moveTo(70,805, duration=1.5)
                time.sleep(1.5)
                pyautogui.click(button='left')                             
        except Exception as e:
                btn_new=driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_CommonActionList"]/tbody/tr/td[1]/a')
                btn_new.click()
                
        time.sleep(2)
        try:
            pyautogui.moveTo(890,715, duration=1.5)
            time.sleep(3.5)
            pyautogui.click()
        except Exception as e:
            btn_continue=driver.find_element(By.XPATH, '//*[@id="form-container"]/div[5]/div[3]/div[2]/div[3]/a[1]')
            btn_continue.click()
        time.sleep(1)

        actions = ActionChains(driver)
        try:
            #input_company_code=driver.find_element(By.XPATH, '//*[@id="DataPanel_eskCtrlBorder_content"]/div/div/table/tbody/tr[1]/td[2]/div/input')
            #input_company_code.send_keys(df_vendor_update.loc[i, 'company_code'])
            pyautogui.typewrite(df_vendor_update.loc[i, 'company_code'])
            time.sleep(0.5)
            actions.send_keys(Keys.TAB).perform()
        except Exception as e:
            print(f"company_code not input {e}")      
                
        try:
            #input_vendor_number=driver.find_element(By.XPATH, '//*[@id="DataPanel_eskCtrlBorder_content"]/div/div/table/tbody/tr[2]/td[2]/div/input')
            #input_vendor_number.send_keys(str(df_vendor_update.loc[i, 'vendor_number']))
            vendor_number = df_vendor_update.loc[i, 'vendor_number']
            pyautogui.typewrite(str(vendor_number))
            time.sleep(0.5)
            actions.send_keys(Keys.TAB).perform()
        except Exception as e:
            print(f"vendor_number not input {e}")
        
        try:
            #input_vendor_name=driver.find_element(By.XPATH, '//*[@id="DataPanel_eskCtrlBorder_content"]/div/div/table/tbody/tr[3]/td[2]/div/input')
            #input_vendor_name.send_keys(df_vendor_update.loc[i, 'vendor_name'])
            pyautogui.typewrite(str(df_vendor_update.loc[i, 'vendor_name']))
            time.sleep(0.5)
            actions.send_keys(Keys.TAB).perform()
        except Exception as e:
            print(f"vendor_name not input {e}")
                
        """
        try:
            input_street=driver.find_element(By.XPATH, '//*[@id="VendorAddress_eskCtrlBorder_content"]/div/div/table/tbody/tr[2]/td[2]/div/input')
            input_street.send_keys(df_vendor_update.loc[i, 'street'])
            time.sleep(0.5)
        except Exception as e:
            print(e)
        actions.send_keys(Keys.TAB*2).perform()
        try:
            input_city=driver.find_element(By.XPATH, '//*[@id="VendorAddress_eskCtrlBorder_content"]/div/div/table/tbody/tr[4]/td[2]/div/input')
            input_city.send_keys(df_vendor_update.loc[i, 'city'])
            time.sleep(0.5)
        except Exception as e:
            print(e)
        actions.send_keys(Keys.TAB).perform()
        try:
            input_postal_code=driver.find_element(By.XPATH, '//*[@id="VendorAddress_eskCtrlBorder_content"]/div/div/table/tbody/tr[5]/td[2]/div/input')
            input_postal_code.send_keys(df_vendor_update.loc[i, 'postal_code'])
            time.sleep(0.5)
        except Exception as e:
            print(e)
        actions.send_keys(Keys.TAB*2).perform()
        try:
            input_country_code=driver.find_element(By.XPATH, '//*[@id="VendorAddress_eskCtrlBorder_content"]/div/div/table/tbody/tr[7]/td[2]/div/input')
            input_country_code.send_keys(df_vendor_update.loc[i, 'country_code'])
            time.sleep(0.5)
        except Exception as e:
            print(e)
        actions.send_keys(Keys.TAB).perform()
        """

        try:
            pyautogui.moveTo(55,1100, duration=2) #Save
            time.sleep(1)
            pyautogui.click()
            time.sleep(1)
        except Exception as e:
            btn_save=driver.find_element(By.XPATH, '//*[@id="form-footer"]/div[1]/a[1]/span')
            btn_save.click()
            print(f"failed to save: {e}")

        list_company_code.append(df_vendor_update.loc[i, 'company_code'])
        list_vendor_number.append(df_vendor_update.loc[i, 'vendor_number'])
        return list_vendor_number

    
def log_entry(log_file,started_time):
    with open(log_file, 'a') as f:  # Open in append mode
        f.write(f"Process started: {started_time}\n")
        f.write(f"Process completed: {datetime.now()}\n")
        f.write(f"Updated entities: {', '.join(list_company_code)}\n")
        f.write(f"Updated vendor: {str(list_vendor_number)}\n")


def _bool_from_env(var_name: str, default: bool = False) -> bool:
    """Return True if the environment variable is truthy, otherwise False."""
    raw_value = os.getenv(var_name)
    if raw_value is None:
        return default
    return raw_value.strip().lower() not in {"0", "false", "no", ""}


def main(
    mode: str = "worker",
    json_dir: Path | str | None = None,
    log_dir: Path | str | None = None,
    dry_run: bool | None = None,
) -> None:
    global list_company_code, list_vendor_number

    if mode != "worker":
        raise ValueError(f"Unsupported mode '{mode}'. Only 'worker' is implemented.")

    if json_dir is None:
        json_directory = Path(os.getenv("ESKER_VENDOR_JSON_DIR", r"C:/Users/john.tan/Downloads")) #AppData/Local/Temp
    else:
        json_directory = Path(json_dir)

    if log_dir is None:
        log_directory = Path(os.getenv("ESKER_LOG_DIR", r"C:/Users/john.tan/Documents/power_apps_esker_vendor/esker_vendor_update/Log"))
    else:
        log_directory = Path(log_dir)

    use_dry_run = dry_run if dry_run is not None else _bool_from_env("ESKER_DRYRUN", default=False)
    list_company_code = []
    list_vendor_number = []
    started_time = start_time()
    print(started_time)

    if use_dry_run:
        log_file = create_log_file(str(log_directory))
        print("Dry-run enabled; skipping Selenium automation.")
        log_entry(log_file, started_time)
        print("Process completed (dry-run).")
        time.sleep(1)
        return

    vendor_update = load_latest_vendor_json(json_directory)
    df_vendor_update = format_vendor_data(vendor_update)
    log_file = create_log_file(str(log_directory))

    driver = create_driver()
    try:
        login_esker(driver)

        x_path_hover = '//*[@id="mainMenuBar"]/td/table/tbody/tr/td[36]/a/div'
        hover_arrow(driver, x_path_hover)

        try:
            tables = driver.find_element(By.XPATH, '//*[@id="CUSTOMTABLE_TAB_100872176"]')
            tables.click()
            time.sleep(1)
        except Exception as e:
            print(e)
        time.sleep(1)

        vendor_update_process(driver, df_vendor_update)
        time.sleep(2)
    finally:
        driver.quit()

    log_entry(log_file, started_time)
    print("Process completed.")
    time.sleep(1)


def _main_cli(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Esker vendor update automation entry point.")
    parser.add_argument(
        "--mode",
        choices=["worker"],
        default="worker",
        help="Execution mode. Only 'worker' mode is currently supported.",
    )
    parser.add_argument(
        "--json-dir",
        type=Path,
        help="Directory containing vendor JSON payloads. Defaults to ESKER_VENDOR_JSON_DIR or the temp directory.",
    )
    parser.add_argument(
        "--log-dir",
        type=Path,
        help="Directory to store execution logs. Defaults to ESKER_LOG_DIR or the vendor update log folder.",
    )
    parser.add_argument(
        "--dry-run",
        dest="dry_run",
        action="store_true",
        help="Skip Selenium automation and only perform validation/logging.",
    )
    parser.add_argument(
        "--no-dry-run",
        dest="dry_run",
        action="store_false",
        help="Force disable dry-run mode even if the environment variable is set.",
    )
    parser.set_defaults(dry_run=None)

    args = parser.parse_args(argv)

    try:
        main(
            mode=args.mode,
            json_dir=args.json_dir,
            log_dir=args.log_dir,
            dry_run=args.dry_run,
        )
    except Exception as exc:  # pragma: no cover - CLI entry point
        raise SystemExit(f"Error: {exc}") from exc


if __name__ == "__main__":
    _main_cli()
