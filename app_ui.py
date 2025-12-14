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
from pathlib import Path

from functools import lru_cache

def _load_env_file_best_effort(path: Path) -> None:
    """Best-effort .env loader (no override)."""
    try:
        if not path.exists():
            return
        for raw_line in path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()
            if (
                len(value) >= 2
                and (value[0] == value[-1])
                and value[0] in {"'", '"'}
            ):
                value = value[1:-1]
            if key and key not in os.environ:
                os.environ[key] = value
    except Exception:
        return


def _load_repo_dotenv_best_effort() -> None:
    """Load repo root .env so worker flags apply even outside Flask."""
    repo_env = Path(__file__).resolve().parents[1] / ".env"
    try:
        from dotenv import load_dotenv  # type: ignore

        load_dotenv(dotenv_path=repo_env, override=False)
    except Exception:
        _load_env_file_best_effort(repo_env)


_load_repo_dotenv_best_effort()

def create_driver():
    options = webdriver.ChromeOptions()
    service = Service(log_path=os.devnull)
    if hasattr(subprocess, "CREATE_NO_WINDOW"):
        service.creationflags = subprocess.CREATE_NO_WINDOW
    return webdriver.Chrome(service=service, options=options)

def login_esker(driver):
    #driver = webdriver.Chrome()
    driver.get("https://az3.ondemand.e@@@@.com/ondemand/webaccess/asf/home.aspx")
    driver.maximize_window()
    time.sleep(1)

    driver.find_element(By.XPATH, '//*[@id="ctl03_tbUser"]').send_keys("john.tan@email.com")
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


def format_gl_data(df_gl_update: pd.DataFrame) -> pd.DataFrame:
    """
    Normalise GL payload data so downstream automation receives clean strings.
    """
    df_gl_update = df_gl_update.copy()
    for column in ("account", "coding_block", "company_code", "description"):
        if column in df_gl_update.columns:
            df_gl_update[column] = df_gl_update[column].astype(str).str.strip()
    df_gl_update.fillna("", inplace=True)
    return df_gl_update

_CODE_SPLIT_PATTERN = re.compile(r"\s*;\s*")


def _normalize_company_codes(value) -> list[str]:
    """Return a list of company codes from a string or iterable."""
    if value is None:
        return []
    if isinstance(value, (list, tuple, set)):
        codes: list[str] = []
        for item in value:
            codes.extend(_normalize_company_codes(item))
        return [code for code in codes if code]
    text = str(value).strip()
    if not text:
        return []
    return [code.strip() for code in _CODE_SPLIT_PATTERN.split(text) if code.strip()]

def build_vendor_frame(company_code: str, vendor_number: str, vendor_name: str) -> pd.DataFrame:
    codes = _normalize_company_codes(company_code)
    if not codes:
        raise ValueError("Vendor payload missing company_code.")
    vendor_number_str = "" if vendor_number is None else str(vendor_number).strip()
    vendor_name_str = "" if vendor_name is None else str(vendor_name).strip()
    return pd.DataFrame(
        {
            "company_code": codes,
            "vendor_number": [vendor_number_str] * len(codes),
            "vendor_name": [vendor_name_str] * len(codes),
        }
    )


def build_frame(account: str, coding_block: str, company_code: str, description: str) -> pd.DataFrame:
    codes = _normalize_company_codes(company_code)
    if not codes:
        raise ValueError("GL payload missing company_code.")
    account_str = "" if account is None else str(account).strip()
    coding_block_str = "" if coding_block is None else str(coding_block).strip()
    description_str = "" if description is None else str(description).strip()
    return pd.DataFrame(
        {
            "account": [account_str] * len(codes),
            "coding_block": [coding_block_str] * len(codes),
            "company_code": codes,
            "description": [description_str] * len(codes),
        }
    )


def parse_vendor_payload(payload: dict, json_path: Path) -> pd.DataFrame:
    """Return a DataFrame with vendor details extracted from the payload."""
    pattern = re.compile(
        r"^\s*([A-Za-z0-9]+(?:\s*;\s*[A-Za-z0-9]+)*)\s+([A-Za-z0-9]+)\s+(.+?)\s*$"
    )
    frames: list[pd.DataFrame] = []

    triplet = payload.get("triplet")
    if triplet:
        if isinstance(triplet, dict):
            code = triplet.get("company_code") or triplet.get("companyCode")
            number = triplet.get("vendor_number") or triplet.get("vendorNumber")
            name = triplet.get("vendor_name") or triplet.get("vendorName")
            if code and number and name:
                frames.append(build_vendor_frame(code, str(number), name))
        elif isinstance(triplet, (list, tuple)):
            if len(triplet) == 3 and not any(isinstance(item, (list, tuple, dict)) for item in triplet):
                code, number, name = triplet
                if code and number and name:
                    frames.append(build_vendor_frame(str(code), str(number), str(name)))
            else:
                for item in triplet:
                    try:
                        frames.append(parse_vendor_payload({"triplet": item}, json_path))
                    except ValueError:
                        continue
        elif isinstance(triplet, str):
            match = pattern.match(triplet)
            if match:
                frames.append(build_vendor_frame(match.group(1), match.group(2), match.group(3)))
        if frames:
            return pd.concat(frames, ignore_index=True)

    for key in ("body", "body_text", "bodyText", "bodyPreview"):
        value = payload.get(key)
        if not value:
            continue
        for line in (seg.strip() for seg in re.split(r"[\r\n]+", str(value)) if seg.strip()):
            match = pattern.match(line)
            if match:
                frames.append(build_vendor_frame(match.group(1), match.group(2), match.group(3)))
    if frames:
        return pd.concat(frames, ignore_index=True)

    subject_value = payload.get("subject")
    if subject_value:
        for line in (seg.strip() for seg in re.split(r"[\r\n]+", str(subject_value)) if seg.strip()):
            match = pattern.match(line)
            if match:
                frames.append(build_vendor_frame(match.group(1), match.group(2), match.group(3)))
    if frames:
        return pd.concat(frames, ignore_index=True)

    raise ValueError(
        f"Unable to parse vendor details in {json_path.name}; inspected keys "
        f"{[key for key in ('triplet', 'body', 'body_text', 'bodyText', 'bodyPreview', 'subject')]}"
    )


def parse_gl_payload(payload: dict, json_path: Path) -> pd.DataFrame:
    """Return a DataFrame with GL details extracted from the payload."""
    pattern_full = re.compile(
        r"^\s*([A-Za-z0-9]*\d[A-Za-z0-9]*)\s+([A-Za-z0-9]+)\s+([A-Za-z0-9]+(?:\s*;\s*[A-Za-z0-9]+)*)\s+(.+?)\s*$"
    )
    pattern_partial = re.compile(
        r"^\s*([A-Za-z0-9]*\d[A-Za-z0-9]*)\s+([A-Za-z0-9]+(?:\s*;\s*[A-Za-z0-9]+)*)\s+(.+?)\s*$"
    )

    def frame_from_string(value: str) -> pd.DataFrame | None:
        text = str(value or "").strip()
        if not text:
            return None
        match_full = pattern_full.match(text)
        if match_full:
            return build_frame(
                match_full.group(1),
                match_full.group(2),
                match_full.group(3),
                match_full.group(4),
            )
        match_partial = pattern_partial.match(text)
        if match_partial:
            return build_frame(
                match_partial.group(1),
                "",
                match_partial.group(2),
                match_partial.group(3),
            )
        return None

    quadruplet = payload.get("quadruplet")
    if quadruplet:
        if isinstance(quadruplet, dict):
            account = quadruplet.get("account") or quadruplet.get("gl_account") or quadruplet.get("glAccount")
            coding_block = (
                quadruplet.get("coding_block")
                or quadruplet.get("codingBlock")
                or ""
            )
            company_code = quadruplet.get("company_code") or quadruplet.get("companyCode")
            description = (
                quadruplet.get("description")
                or quadruplet.get("gl_description")
                or quadruplet.get("glDescription")
            )
            if account and company_code and description:
                return build_frame(
                    str(account),
                    str(coding_block),
                    str(company_code),
                    str(description),
                )
        elif isinstance(quadruplet, (list, tuple)):
            if len(quadruplet) == 4:
                account, coding_block, company_code, description = quadruplet
                if account and company_code and description:
                    return build_frame(
                        str(account),
                        str(coding_block or ""),
                        str(company_code),
                        str(description),
                    )
            elif len(quadruplet) == 3:
                account, company_code, description = quadruplet
                if account and company_code and description:
                    return build_frame(
                        str(account),
                        "",
                        str(company_code),
                        str(description),
                    )
        elif isinstance(quadruplet, str):
            frame = frame_from_string(quadruplet)
            if frame is not None:
                return frame

    frames: list[pd.DataFrame] = []
    for key in ("body", "body_text", "bodyText", "bodyPreview"):
        value = payload.get(key)
        if not value:
            continue
        for line in (seg.strip() for seg in re.split(r"[\r\n]+", str(value)) if seg.strip()):
            frame = frame_from_string(line)
            if frame is not None:
                frames.append(frame)
    if frames:
        return pd.concat(frames, ignore_index=True)

    subject_value = payload.get("subject")
    if subject_value:
        for line in (seg.strip() for seg in re.split(r"[\r\n]+", str(subject_value)) if seg.strip()):
            frame = frame_from_string(line)
            if frame is not None:
                frames.append(frame)
    if frames:
        return pd.concat(frames, ignore_index=True)

    raise ValueError(
        f"Unable to parse GL details in {json_path.name}; inspected keys "
        f"{[key for key in ('quadruplet', 'body', 'body_text', 'bodyText', 'bodyPreview', 'subject')]}"
    )


def dataframe_from_payload(payload: dict, json_path: Path) -> tuple[pd.DataFrame, str]:
    """Return a DataFrame and payload type based on the email subject or payload hints."""
    subject = str(payload.get("subject", "")).lower()
    if "esker gl email" in subject:
        return parse_gl_payload(payload, json_path), "gl"
    if "esker vendor email" in subject:
        return parse_vendor_payload(payload, json_path), "vendor"

    if "quadruplet" in payload:
        return parse_gl_payload(payload, json_path), "gl"
    if "triplet" in payload:
        return parse_vendor_payload(payload, json_path), "vendor"

    raise ValueError(f"Unable to determine payload type for {json_path.name}")


def load_latest_payload_dataframe(json_dir: Path) -> tuple[pd.DataFrame, str, Path]:
    """
    Return the newest parsable payload inside the provided directory.

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
            with open(json_path, encoding="utf-8") as source:
                payload = json.load(source)
            dataframe, payload_type = dataframe_from_payload(payload, json_path)
            return dataframe, payload_type, json_path
        except (ValueError, json.JSONDecodeError) as exc:
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
list_company_code: list[str] = []
list_vendor_number: list[str] = []
list_vendor_name: list[str] = []
list_gl_account: list[str] = []
list_gl_description: list[str] = []


def _format_unique(values: list[str]) -> str:
    """Return space-separated unique values preserving the original order."""
    seen: set[str] = set()
    ordered: list[str] = []
    for value in values:
        item = str(value).strip()
        if not item or item in seen:
            continue
        seen.add(item)
        ordered.append(item)
    return " ".join(ordered)


MASTER_SOURCE_DIR = Path(r"C:/Users/john.tan/Downloads")
MASTER_FILES = {
    "vendor": MASTER_SOURCE_DIR / "master_vendor.txt",
    "gl": MASTER_SOURCE_DIR / "master_gl.txt",
}


def _normalise_token(value: str) -> str:
    return " ".join(str(value or "").strip().split()).upper()


@lru_cache(maxsize=None)
def _load_master_entries(kind: str) -> set[str]:
    path = MASTER_FILES.get(kind)
    if not path:
        return set()
    try:
        with path.open(encoding="utf-8") as handle:
            return {
                " ".join(line.strip().upper().split())
                for line in handle
                if line.strip()
            }
    except FileNotFoundError:
        print(f"[warning] Master file for {kind} not found at {path}; treating as empty.")
        return set()


def _vendor_signature(row) -> str:
    return " ".join(
        filter(
            None,
            [
                _normalise_token(row.get("company_code", "")),
                _normalise_token(row.get("vendor_number", "")),
                _normalise_token(row.get("vendor_name", "")),
            ],
        )
    )


def _gl_signature(row) -> str:
    tokens = [
        _normalise_token(row.get("account", "")),
    ]
    coding = _normalise_token(row.get("coding_block", ""))
    if coding:
        tokens.append(coding)
    tokens.append(_normalise_token(row.get("company_code", "")))
    tokens.append(_normalise_token(row.get("description", "")))
    return " ".join(filter(None, tokens))


def _filter_against_master(df: pd.DataFrame, payload_type: str) -> pd.DataFrame:
    master_entries = _load_master_entries(payload_type)
    if not master_entries:
        return df

    signatures: list[str] = []
    for _, row in df.iterrows():
        if payload_type == "vendor":
            signatures.append(_vendor_signature(row))
        else:
            signatures.append(_gl_signature(row))

    keep_indices = [sig not in master_entries for sig in signatures]
    if not any(keep_indices):
        return df.iloc[0:0].copy()

    filtered = df.loc[keep_indices].copy()
    filtered.reset_index(drop=True, inplace=True)
    dropped = len(df) - len(filtered)
    if dropped:
        print(f"[info] Skipped {dropped} existing {payload_type} entries found in master list.")
    return filtered


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
        actions = ActionChains(driver)

        try:
            pyautogui.moveTo(65,475, duration=2) #move to tables_input_search_box
            time.sleep(0.5)
            pyautogui.click()
            time.sleep(1)
            pyautogui.typewrite("S2P - Vendors")
            time.sleep(2.5)
            actions.send_keys(Keys.ENTER).perform()
        except Exception as e:
            time.sleep(0.5)
            #print(f"search box not found: {e}")
            
            time.sleep(1)
            #s2p_vendors=driver.find_element(By.XPATH, '//*[@id="ViewSelector"]/div/div/div/div[1]/div[1]/span')
            #s2p_vendors.click()
                
        try:                
                pyautogui.moveTo(70,805, duration=1.5)
                time.sleep(1.0)
                pyautogui.click(button='left')
                time.sleep(1)                          
        except Exception as e:
                btn_new=driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_CommonActionList"]/tbody/tr/td[1]/a')
                btn_new.click()
                
        time.sleep(2)
        try:
            pyautogui.moveTo(890,715, duration=1.5)
            time.sleep(2)
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
        list_vendor_name.append(df_vendor_update.loc[i, 'vendor_name'])

    return list_vendor_number


def gl_update_process(driver, df_gl_update: pd.DataFrame):
    """
    Placeholder GL update automation. Captures payload data for future implementation.
    Returns a list of identifiers mirroring vendor_update_process for compatibility.
    """
    for i in range(len(df_gl_update)):
        account = str(df_gl_update.loc[i, 'account'])
        coding_block = str(df_gl_update.loc[i, 'coding_block'])
        company_code = str(df_gl_update.loc[i, 'company_code'])
        description = str(df_gl_update.loc[i, 'description'])

        print(
            f"GL update received for account {account}, coding block {coding_block}, "
            f"company {company_code}: {description}"
        )


        def hover_arrow(driver, x_path):
            elem_to_hover = driver.find_element(By.XPATH, x_path)
            hover = ActionChains(driver).move_to_element(elem_to_hover)
            hover.perform()

        pyautogui.moveTo(35,350, duration=2) #move cursor to extreme left side
        time.sleep(1)
        pyautogui.click()

        x_path='//*[@id="mainMenuBar"]/td/table/tbody/tr/td[31]/a/div'
        try:
            hover_arrow(driver, x_path)
        except Exception as e:
            time.sleep(0.5) 
        pyautogui.moveTo(1690,370, duration=2.5) #arrow icon
        time.sleep(0.5)
        pyautogui.click()
        time.sleep(1)

        actions = ActionChains(driver)
        actions.send_keys(Keys.TAB).perform()
        actions.send_keys(Keys.TAB).perform()
        actions.send_keys(Keys.TAB).perform()
        actions.send_keys(Keys.ENTER).perform()

        pyautogui.moveTo(65,475, duration=2) #move to tables_input_search_box
        time.sleep(0.5)
        pyautogui.click()
        time.sleep(2)
        pyautogui.typewrite("[Manual Import] S2P - G/L accounts")
        actions.send_keys(Keys.ENTER).perform()

        try:                
            pyautogui.moveTo(70,805, duration=1.5)
            time.sleep(1)
            pyautogui.click(button='left')                             
        except Exception as e:
            btn_new=driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_CommonActionList"]/tbody/tr/td[1]/a')
            btn_new.click()
        time.sleep(1) 
        
        try:
            pyautogui.moveTo(890,715, duration=1.5)
            time.sleep(2.5)
            pyautogui.click()
        except Exception as e:
            btn_continue=driver.find_element(By.XPATH, '//*[@id="form-container"]/div[5]/div[3]/div[2]/div[3]/a[1]')
            btn_continue.click()
        time.sleep(1)

        pyautogui.moveTo(225,310,duration=2) #first input box
        pyautogui.click()

        pyautogui.typewrite(company_code)
        actions.send_keys(Keys.TAB).perform()
        pyautogui.typewrite(account)
        actions.send_keys(Keys.TAB).perform()
        pyautogui.typewrite(coding_block)
        actions.send_keys(Keys.TAB).perform()
        pyautogui.typewrite(description)
        actions.send_keys(Keys.TAB).perform()
        time.sleep(0.5)

        pyautogui.moveTo(45,1085, duration=2) #move to Save button
        time.sleep(1.5)
        pyautogui.click()

        list_company_code.append(company_code)
        list_gl_account.append(account)
        list_gl_description.append(description)
        list_vendor_number.append(account)

    return list_vendor_number


def log_entry(log_file: str, started_time: datetime, payload_type: str) -> None:
    """Append a summary of the automation run to the log file."""
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f"Process started: {started_time}\n")
        f.write(f"Process type: {payload_type}\n")
        f.write(f"Process completed: {datetime.now()}\n")

        if list_company_code:
            f.write(f"Company codes: {_format_unique(list_company_code)}\n")

        if payload_type == "vendor" and list_vendor_number:
            f.write(f"Vendors: {', '.join(map(str, list_vendor_number))}\n")
            f.write(f"Vendor names: {', '.join(map(str, list_vendor_name))}\n")
        elif payload_type == "gl":
            if list_gl_account:
                f.write(f"GL accounts: {', '.join(map(str, list_gl_account))}\n")
            if list_gl_description:
                f.write(f"Descriptions: {', '.join(map(str, list_gl_description))}\n")



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
    skip_master: bool | None = None,
    force_run: bool | None = None,
) -> None:
    global list_company_code, list_vendor_number, list_gl_account, list_gl_description

    if mode != "worker":
        raise ValueError(f"Unsupported mode '{mode}'. Only 'worker' is implemented.")

    if json_dir is None:
        json_directory = Path(os.getenv("ESKER_VENDOR_JSON_DIR", r"C:/Users/john.tan/Downloads"))  # AppData/Local/Temp
    else:
        json_directory = Path(json_dir)

    if log_dir is None:
        log_directory = Path(os.getenv("ESKER_LOG_DIR", r"C:/Users/john.tan/Documents/power_apps_esker_vendor/esker_vendor_update/Log"))
    else:
        log_directory = Path(log_dir)

    use_dry_run = dry_run if dry_run is not None else _bool_from_env("ESKER_DRYRUN", default=False)
    use_skip_master = skip_master if skip_master is not None else _bool_from_env("ESKER_SKIP_MASTER", default=False)
    use_force_run = force_run if force_run is not None else _bool_from_env("ESKER_FORCE_RUN", default=False)
    list_company_code = []
    list_vendor_number = []
    list_gl_account = []
    list_gl_description = []

    started_time = start_time()
    print(started_time)

    log_file = create_log_file(str(log_directory))

    if use_dry_run:
        print("Dry-run enabled; skipping Selenium automation.")
        log_entry(log_file, started_time, "dry-run")
        print("Process completed (dry-run).")
        time.sleep(1)
        return

    payload_df, payload_type, payload_path = load_latest_payload_dataframe(json_directory)
    if payload_type == "vendor":
        df_payload = format_vendor_data(payload_df)
    else:
        df_payload = format_gl_data(payload_df)

    df_to_process = df_payload
    if use_skip_master:
        print("[info] ESKER_SKIP_MASTER enabled; running without master list filtering.")
    else:
        df_to_process = _filter_against_master(df_payload, payload_type)

    if df_to_process.empty:
        if use_force_run and not df_payload.empty:
            print(
                f"[info] ESKER_FORCE_RUN enabled; running automation even though no new {payload_type} entries "
                "were detected after master list filtering."
            )
            df_to_process = df_payload
        else:
            print(f"No new {payload_type} entries detected; skipping automation run.")
            log_entry(log_file, started_time, payload_type)
            print("Process completed (no automation required).")
            time.sleep(1)
            return

    print(f"Processing {payload_type} payload from {payload_path.name}")

    driver = create_driver()
    try:
        login_esker(driver)

        if payload_type == "vendor":
            x_path_hover = '//*[@id="mainMenuBar"]/td/table/tbody/tr/td[36]/a/div'
            hover_arrow(driver, x_path_hover)

            try:
                tables = driver.find_element(By.XPATH, '//*[@id="CUSTOMTABLE_TAB_100872176"]')
                tables.click()
                time.sleep(1)
            except Exception as e:
                print(e)
            time.sleep(1)

            vendor_update_process(driver, df_to_process)
            time.sleep(2)
        else:
            gl_update_process(driver, df_to_process)
    finally:
        driver.quit()

    log_entry(log_file, started_time, payload_type)
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
    parser.add_argument(
        "--skip-master",
        dest="skip_master",
        action="store_true",
        help="Skip filtering against the master list (also via ESKER_SKIP_MASTER=1).",
    )
    parser.add_argument(
        "--no-skip-master",
        dest="skip_master",
        action="store_false",
        help="Force enable master list filtering even if the environment variable is set.",
    )
    parser.add_argument(
        "--force-run",
        dest="force_run",
        action="store_true",
        help="Run Selenium even if the master filter removes all rows (also via ESKER_FORCE_RUN=1).",
    )
    parser.add_argument(
        "--no-force-run",
        dest="force_run",
        action="store_false",
        help="Do not force Selenium runs even if the environment variable is set.",
    )
    parser.set_defaults(dry_run=None)
    parser.set_defaults(skip_master=None)
    parser.set_defaults(force_run=None)

    args = parser.parse_args(argv)

    try:
        main(
            mode=args.mode,
            json_dir=args.json_dir,
            log_dir=args.log_dir,
            dry_run=args.dry_run,
            skip_master=args.skip_master,
            force_run=args.force_run,
        )
    except Exception as exc:  # pragma: no cover - CLI entry point
        raise SystemExit(f"Error: {exc}") from exc


if __name__ == "__main__":
    _main_cli()
