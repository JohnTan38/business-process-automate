from selenium import webdriver # (1) login CDAS
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import time, os

driver = webdriver.Chrome()
#driver.get("https://az3.ondemand.e@@@@.com/ondemand/webaccess/asf/home.aspx")
driver.get("https://invoice.eservices.c@@@.link/login")
driver.maximize_window()
time.sleep(3)

import pyautogui
pyautogui.moveTo(720, 710, duration=1.5)
pyautogui.click(button='left')
pyautogui.typewrite("user.name@email.com.sg")
pyautogui.press('tab')
pyautogui.typewrite("PASSWORD")
pyautogui.press('enter')
time.sleep(1)

try:
    all_invoices = driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[1]/aside/div/div/a[2]/div[3]')
    all_invoices.click()
    time.sleep(0.5)
except Exception as e:
    print(e)

def hover(driver, x_path):
    elem_to_hover = driver.find_element(By.XPATH, x_path)
    hover = ActionChains(driver).move_to_element(elem_to_hover)
    hover.perform()

def hover_click(driver, x_path):
    elem_to_hover = driver.find_element(By.XPATH, x_path)
    #hover = ActionChains(driver).move_to_element(elem_to_hover)
    hover = ActionChains(driver).click(elem_to_hover)
    hover.perform()
time.sleep(1)

#(2) download all bills
import pandas as pd
import pathlib
from pathlib import Path
global username

# CODEX added code to read username from environment variable or default to system username
username = os.environ.get("CDAS_USERNAME") or Path.home().name
path_bill = Path(f"C:/Users/{username}/Downloads")
source_path = os.environ.get("ESKER_INVOICE_WORKBOOK", path_bill / "cdas_n.xlsx")
source_sheet = os.environ.get("ESKER_INVOICE_SHEET", "bill")
source_column = os.environ.get("ESKER_INVOICE_COLUMN", "bill")

df_bill = pd.read_excel(source_path, sheet_name=source_sheet, engine="openpyxl")
if source_column not in df_bill.columns:
    raise KeyError(...)
list_bill = (
    df_bill[source_column]
    .dropna()
    .astype(str)
    .str.strip()
    .tolist()
)

#username = 'john.tan' #get_username_date_to_download()[0]

# CODEX TO READ BILL LIST FROM EXCEL added 20251105
source_path = os.environ.get("ESKER_INVOICE_WORKBOOK", Path(path_bill) / "cdas_n.xlsx")
source_sheet = os.environ.get("ESKER_INVOICE_SHEET", "bill")
source_column = os.environ.get("ESKER_INVOICE_COLUMN", "bill")

df_bill = pd.read_excel(source_path, sheet_name=source_sheet, engine="openpyxl")
if source_column not in df_bill:
    raise KeyError(f"{source_column} column missing in {source_sheet} from {source_path}")

list_bill = (
    df_bill[source_column]
    .dropna()
    .astype(str)
    .str.strip()
    .tolist()
)

"""
path_bill = "C:/Users/"+username+"/Downloads/"
path_move_merged = "C:/Users/"+username+"/Documentss/AP/"
df_bill = pd.read_excel(path_bill+ r'cdas_n.xlsx', sheet_name='bill', engine='openpyxl')
list_bill = df_bill['bill'].tolist()
list_bill = list(filter(None, list_bill)) #remove empty values
#list_bill = list(set(list_bill)) #unique bill ref
"""

date_to_download_0 = str(df_bill['date_to_download'][0])
date_to_download = (date_to_download_0[-2:])
date_to_download = date_to_download.lstrip("0")
print(f"date to download: {date_to_download}")

def advanced_filter_calendar(driver):
        x_path_hover = '//*[@id="q-app"]/div/div[2]/div/div[2]/div[2]/div/div' #Advanced filter
        hover(driver, x_path_hover)
        time.sleep(0.5)
        x_path_hover_click = '//*[@id="q-app"]/div/div[2]/div/div[2]/div[2]/div/div'
        hover_click(driver, x_path_hover_click)
        
        pyautogui.press('pageup')        
        pyautogui.moveTo(530,760, duration=2) ##calendar icon
        time.sleep(1)
        pyautogui.click(button='left')
        time.sleep(0.5)
        

def click_date_to_download(date_to_download, date_to_download_0):     
        #date_to_download = str(yesterday(frmt='%Y%m%d', string=True))
        
        date_to_download = str(date_to_download)
        def process_date_to_download(date_to_download):       
            if len(date_to_download) >=2 and date_to_download[:-2] == '0':
                return date_to_download[:-2] + date_to_download[-1] # Remove '0' from the string if it exists
            else:
                return date_to_download # Otherwise, return the string as is
        
        from datetime import datetime
        def month_diff(date_to_download_0):            
                target_date = datetime.strptime(date_to_download_0, '%Y%m%d') # Convert date string to datetime object  
                current_date = datetime.now() # Get current date    
                # Calculate month difference
                month_diff = (target_date.year - current_date.year) * 12 + target_date.month - current_date.month
                return abs(month_diff)
        
        date_to_download = (process_date_to_download(date_to_download))
        month_to_download_back_click=month_diff(date_to_download_0)
        if month_to_download_back_click > 0:
                pyautogui.press("pageup")
                actions=ActionChains(driver)
                actions.send_keys("PAGEUP").perform()
                #pyautogui.moveTo(360,520, duration=1.5) #back click to month
                chevron_left = driver.find_element(By.XPATH, "/html/body/div[3]/div/div[2]/div[1]/div/div[1]/div[1]/button/span[2]/span/i")
                for _ in range(month_to_download_back_click):
                        chevron_left.click()
                        time.sleep(1)

        idx_to_click = str(int(date_to_download) + 1) ##3
        print(f"idx to click: {idx_to_click}")
        x_path_date = '/html/body/div[3]/div/div[2]/div[1]/div/div[3]/div/div['+idx_to_click+']/button/span[2]/span/span'
        
        
        time.sleep(0.5)
        date = driver.find_element(By.XPATH, x_path_date)
        if date.get_attribute('textContent') == date_to_download:
                date.click()
                date.click()        
                time.sleep(3.5)

        btn_close=driver.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/div[2]/div[3]/button[2]/span[2]/span')
        btn_close.click()
        time.sleep(0.5)
        pyautogui.press('pagedown')
        time.sleep(0.5)
    

def all_invoices():
    try:
        #all_invoices = driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[1]/aside/div/div/a[2]/div[3]')
        pyautogui.moveTo(180, 345, duration=1.5)
        pyautogui.click(button='left')
        time.sleep(0.5)
    except Exception as e:
        print(e)        

def bill_transaction_number(bill_ref):
        pyautogui.moveTo(400, 490, duration=1) #move to white space
        pyautogui.click(button='left')
        pyautogui.press('pgup')       
        
        pyautogui.moveTo(880, 610, duration=1) #Bill Transaction Number. (620,450) Advance Filter
        pyautogui.click(button='left')
        for _ in range(10):
                pyautogui.press('backspace')
        bill_ref=str(bill_ref)
        pyautogui.typewrite(bill_ref)
        time.sleep(1)
        actions = ActionChains(driver)
        for _ in range(11):
            pyautogui.press('tab')
            
        time.sleep(0.5)
        pyautogui.press('enter')
        pyautogui.press('pagedown')
        time.sleep(0.5)
                
        """
        for _ in range(11):
                #tab_press=actions.send_keys(Keys.TAB)
                #tab_press.perform()
        pyautogui.press('enter')
        """
        for _ in range(3):
                pyautogui.press('pagedown')
        time.sleep(1)
        

from pathlib import Path
import pathlib

import mss
import mss.tools

def get_screenshot(path_output):
        with mss.mss() as sct:
            # The screen part to capture
            monitor = {"top": 190, "left": 1360, "width": 320, "height": 90}
    
            sct_img = sct.grab(monitor) # Grab the data
            # Save to the picture file
            mss.tools.to_png(sct_img.rgb, sct_img.size, output=path_output)

path_output = r"C:/Users/"+username+r"/Downloads/cdas_merged/screenshot.png"
    
from PIL import Image
import pytesseract
import re

def extract_text(image_path):
        # Set the path to the Tesseract executable (only required on Windows)
        pytesseract.pytesseract.tesseract_cmd = r'C:/Users/john.tan/AppData/Local/Programs/Tesseract-OCR/tesseract.exe'
        image = Image.open(image_path) # Open the image using PIL

        # Preprocess the image (convert to grayscale and apply thresholding)
        image = image.convert('L')  # Convert to grayscale
        image = image.point(lambda x: 0 if x < 128 else 255, '1')  # Apply thresholding

        text = pytesseract.image_to_string(image) # Extract text from the image
        # Use regex to remove special characters like '+' and newline characters
        cleaned_text = re.sub(r'[+\n]', '', text).strip()
        return cleaned_text

def check_screenshot_exists(path_output):
        if Path(path_output).exists():
            #print("screenshot file exists")
            #screenshot += 1
            time.sleep(0.5)
        else:
            time.sleep(1.5)
            get_screenshot(path_output)
            extract_text(path_output)
            list_save_as_pdf = ["Save as PDF", "Microsoft Print to PDF"]
            if extract_text(path_output) not in list_save_as_pdf:            
                pyautogui.moveTo(1500,225, duration=1.5) #move to 'save as pdf'
                pyautogui.click(button='left')
                pyautogui.press('down')
                pyautogui.press('return')


def remove_existing_screenshot(dir_screenshot):
    directory = Path(dir_screenshot)
    if not directory.exists():
        return
    for filename in ("screenshot.png", "screenshot.PNG"):
        screenshot_path = directory / filename
        if screenshot_path.exists():
            try:
                screenshot_path.unlink()
            except OSError as exc:
                print(f"Warning: unable to delete {screenshot_path}: {exc}")


def navigate_to_view(driver, bill_ref):
    """
    try:        
        document=driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[2]/div/div[2]/div[3]/div[2]/table/tbody/tr/td[5]/div/div/div/div/a')
        document_ref=document.get_attribute('textContent')
        document_ref = document_ref.replace(".pdf", "")
        print(document_ref)
        if document_ref.strip() == str(bill_ref).strip():
            document.click()
            time.sleep(1)
    except Exception as e:
        time.sleep(0.5)    
        #break
    """        

    try:        
            view = driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[2]/div/div[2]/div[3]/div[2]/table/tbody/tr/td[1]/button/span[2]/span')
            view.click()
            time.sleep(0.5)
    except Exception as e:
            pyautogui.moveTo(550,910, duration=2) #VIEW button
            pyautogui.click(button='left')
            time.sleep(0.5)

def process_epay(driver):    
    try:
            pyautogui.moveTo(525,685, duration=2) #download arrow EPAY
            time.sleep(0.5)
            pyautogui.click(button='left')
            time.sleep(3.5) #wait prompt to disappear
    except Exception as e:
            time.sleep(0.5)
    try:
            pyautogui.moveTo(550,820, duration=1.5) #PRINT button (EPAY)
            pyautogui.click(button='left')
    except Exception as e:
            time.sleep(0.5)

def move_to_save_as_pdf():
        pyautogui.moveTo(1500, 960, duration=2.5) #move to Print btn (save) as pdf
        time.sleep(2)
        pyautogui.click(button='left')
        time.sleep(0.5)

def print_save_pdf(bill_ref):
        try:
            """
            try:
                btn_print=driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[2]/div/div[2]/div[3]/div/div/button/span[2]/span/span')
                time.sleep(1.5)
                btn_print.click()
                time.sleep(0.5)
            except Exception as e:
                pyautogui.moveTo(550, 820, duration=1.5) ##PRINT button
                time.sleep(1.5)
                pyautogui.click(button='left')
                time.sleep(3)
            check_screenshot_exists(path_output)
            
            time.sleep(0.5)
            move_to_save_as_pdf()
            
            for _ in range(3):
                pyautogui.press('tab')
            time.sleep(0.5)
            pyautogui.press('return')
            """
            time.sleep(0.5)
            pyautogui.moveTo(600,580, duration=2.5) # Move to 'File name' input box

            pyautogui.press('delete')
            time.sleep(0.5)
            pyautogui.typewrite(bill_ref+ '_bill')
            pyautogui.press('return')
            time.sleep(0.5)
            list_bill_ref_saved.append(bill_ref)
        except Exception as e:
            print(f'pdf not saved {e}')
            #break


#(3) merge all bills
from PyPDF2 import PdfReader, PdfWriter
import fitz
import os
import time
import hashlib
from io import BytesIO
def get_pdf_files_with_invoice_number(path_pdf, bill_ref):
    """
    Gets a list of PDF files in the given directory that contain the specified invoice number in their filename.
    Args:
        path_pdf: The path to the directory containing the PDF files.
        bill_ref: The invoice number to search for.
    Returns:
        A list of PDF filenames that contain the invoice number or were saved within a time period.
    """
    # Get the current time in seconds since the epoch
    current_time = time.time()
    # Calculate the time 2 minutes ago
    two_minutes_ago = current_time - (0.95 * 60)

    pdf_files = []
    # Iterate through all items in the specified directory
    for f in os.listdir(path_pdf):
        full_path = os.path.join(path_pdf, f)
        
        # Check if the item is a file and ends with '.pdf'
        if os.path.isfile(full_path) and (f.endswith(".pdf") or f.endswith(".PDF")):
                        
            # Check if the filename contains the invoice number
            if bill_ref in f:                
                pdf_files.append(f)
            
            # Check if the file was modified in the last 0.95 minutes. This is a separate condition.
            modification_time = os.path.getmtime(full_path)
            if modification_time >= two_minutes_ago:
                if f not in pdf_files: # Avoid adding duplicate files
                    pdf_files.append(f)
    return pdf_files


def merge_pdfs(pdf_files, path_pdf):
        """
        Merges multiple PDF files into a single PDF.
        Args:
            pdf_files: A list of PDF filenames.
            path_pdf: The path to the directory containing the PDF files.
        Returns:
            The merged PDF writer object.
        """
        pdf_writer = PdfWriter()

        for pdf in pdf_files:
            try:
                with open(path_pdf + pdf, 'rb') as pdf_file:
                    pdf_reader = PdfReader(pdf_file)
                    for page_num in range(len(pdf_reader.pages)):
                        page = pdf_reader.pages[page_num]
                        pdf_writer.add_page(page)
            except EOFError as e:
                #print(f"Error reading PDF file '{pdf}': {e}")
                continue  # Skip the file that caused the error
        return pdf_writer

def _hash_pdf_page(page):
        """Return a stable hash for a PDF page."""
        buffer = BytesIO()
        page_writer = PdfWriter()
        page_writer.add_page(page)
        page_writer.write(buffer)
        return hashlib.sha1(buffer.getvalue()).hexdigest()

def remove_duplicate_pages(pdf_path):
        """
        Remove duplicate pages (regardless of order) from a PDF while keeping the first occurrence.
        Args:
            pdf_path (str | Path): Full path to the merged PDF.
        Returns:
            int: Number of pages removed.
        """
        pdf_path = Path(pdf_path)
        if not pdf_path.exists():
                return 0

        reader = PdfReader(str(pdf_path))
        if len(reader.pages) <= 1:
                return 0

        writer = PdfWriter()
        seen_hashes = set()
        duplicates_removed = 0

        for page in reader.pages:
                current_hash = _hash_pdf_page(page)
                if current_hash in seen_hashes:
                        duplicates_removed += 1
                        continue
                writer.add_page(page)
                seen_hashes.add(current_hash)

        if duplicates_removed:
                with pdf_path.open("wb") as merged_pdf:
                        writer.write(merged_pdf)

        return duplicates_removed
    
import glob, os.path
#import pathlib
#from pathlib import Path
import shutil
    
# define function move pdf
def move_pdf(src_dir, dest_dir):
        """
        Move each *_merged.pdf from the source directory into the destination directory.
        Creates the target directory if it does not exist and appends a timestamp to avoid overwrites.
        """
        from datetime import datetime

        src_path = Path(src_dir)
        dest_path = Path(dest_dir)

        if not src_path.exists():
                return f"Error: Source directory '{src_path}' does not exist."

        try:
                dest_path.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
                return f"Error: Unable to create destination directory '{dest_path}': {exc}"

        moved_files = []
        for pdf_file in src_path.iterdir():
                if not pdf_file.is_file():
                        continue
                if not pdf_file.name.lower().endswith("_merged.pdf"):
                        continue

                target_path = dest_path / pdf_file.name
                if target_path.exists():
                        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                        target_path = dest_path / f"{pdf_file.stem}_{timestamp}{pdf_file.suffix}"

                shutil.move(str(pdf_file), str(target_path))
                moved_files.append(target_path.name)

        if not moved_files:
                return f"Info: No merged PDF files found in '{src_path}'."

        return f"Success: Moved {len(moved_files)} merged PDF file(s) to '{dest_path}'."

def remove_recent_pdf_files(file_path):
    """
    Deletes .pdf and .PDF files created in the last 2 minutes within a specified directory.
    Args:
        file_path (str): The path to the directory to search and remove files from.
    """
    # Get the current time in seconds since the epoch
    current_time = time.time()

    # Calculate the timestamp for 2 minutes ago
    two_minutes_ago = current_time - (2 * 60)

    # Iterate over all items in the directory
    for filename in os.listdir(file_path):
        full_file_path = os.path.join(file_path, filename)
        # Ensure we are dealing with a file and not a directory
        if os.path.isfile(full_file_path):
            try:
                # Use os.path.getctime() to get the creation time on most systems.
                # Note: On some Unix systems, this might be the time of last metadata change.
                file_creation_time = os.path.getctime(full_file_path)

                # Check if the file is a PDF and was created in the last 2 minutes
                if (
                    filename.lower().endswith('.pdf')
                    and not filename.lower().endswith('_merged.pdf')
                    and file_creation_time > two_minutes_ago
                ):
                    os.remove(full_file_path)
            except Exception as e:
                print(f"Error processing file '{filename}': {e}")


from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
def first_second_download():
    try:
        wait = WebDriverWait(driver, 5) # Wait up to 5 seconds
        #first_download=driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[2]/div/div[2]/div[2]/div/div[4]/ul/div/li[1]/button[1]/span[2]/span/i')
        first_download = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="q-app"]/div/div[2]/div/div[2]/div[2]/div/div[4]/ul/div/li[1]/button[1]/span[2]/span/i')))
        time.sleep(1)
        first_download.click() #download attachment files
        time.sleep(2) 
    except Exception as e:
        pyautogui.moveTo(535,680, duration=2) #first 'down' arrow
        pyautogui.click(button='left')
        time.sleep(2)

    time.sleep(2.5)
    try:
        wait = WebDriverWait(driver, 5) # Wait up to 5 seconds
        #second_download= driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[2]/div/div[2]/div[2]/div/div[4]/ul/div/li[2]/button[1]/span[2]/span/i')
        second_download = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="q-app"]/div/div[2]/div/div[2]/div[2]/div/div[4]/ul/div/li[2]/button[1]/span[2]/span/i')))
        
        time.sleep(1)
        second_download.click() #download attachment files
        time.sleep(2)  
    except Exception as e:
        pyautogui.moveTo(535,735, duration=2) #second 'down' arrow
        pyautogui.click(button='left')
        time.sleep(2)
    pyautogui.moveTo(560,875, duration=1.5) #PRINT GOCL GI
    pyautogui.click(button='left')

import os.path
from datetime import datetime

date_today = datetime.now().strftime("%Y%m%d")
file_log = 'cdas_log_'+date_today+'.txt'
path_cdas_log = "C:/Users/"+ username+ "/Downloads/cdas_merged/" #
path_log = path_cdas_log + file_log
if os.path.exists(path_log) == False:
        open(path_log, "w").close #check log file exists, if not create

import logging
# Configure logging (do this once at the beginning of your script)
logging.basicConfig(filename=path_log, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def write_log(log_message):
        if isinstance(log_message, list):
            for item in log_message:
                logging.info(item) # Log each item individually
        else:
            logging.info(log_message) # Log the message as is


list_bill_ref_saved = []
username='john.tan'
path_pdf = r"C:/Users/"+username+"/Downloads/"
write_log(f"processing started {datetime.now().strftime('%Y%m%d %H:%M:%S')} ...")
def main(list_bill):
    username='john.tan'
    dir_screenshot = f"C:/Users/"+username+"/Downloads/cdas_merged/"
    remove_existing_screenshot(dir_screenshot)
    list_bill_ref_saved = []
    for bill_ref in list_bill:
            write_log(f"processing {bill_ref} for day {date_to_download} ...")
            all_invoices()
            time.sleep(1)
            pyautogui.press("pageup")
            advanced_filter_calendar(driver)
            time.sleep(1)
            click_date_to_download(date_to_download, date_to_download_0)
            bill_ref = str(bill_ref)
            #print(bill_ref)
            bill_transaction_number(bill_ref)
            time.sleep(1)
            navigate_to_view(driver, bill_ref)
            time.sleep(1)
            if bill_ref.startswith(('GOCL', 'MGEU')):
                    first_second_download()
                    time.sleep(1)
                    check_screenshot_exists(path_output)
                    move_to_save_as_pdf()
                    print_save_pdf(bill_ref)
            else:
                process_epay(driver)
                check_screenshot_exists(path_output)
                move_to_save_as_pdf() #EPAY GI R
                print_save_pdf(bill_ref)
            time.sleep(1)
            list_bill_ref_saved.append(bill_ref)
        
            try:
                username = 'john.tan'
                path_pdf = r"C:/Users/"+username+"/Downloads/"
                #get_pdf_files_with_invoice_number(path_pdf, bill_ref)
                merged_writer_bill=merge_pdfs(get_pdf_files_with_invoice_number(path_pdf, bill_ref), path_pdf)
                merged_pdf_path = os.path.join(path_pdf, f"{bill_ref}_merged.pdf")
                with open(merged_pdf_path, "wb") as output_file:
                        merged_writer_bill.write(output_file) # Save the merged PDF
            except Exception as e:
                write_log(f"Error merging PDF files for bill '{bill_ref}': {e}")
                continue

            try:
                removed_pages = remove_duplicate_pages(merged_pdf_path)
                if removed_pages:
                        write_log(f"Removed {removed_pages} duplicate page(s) from {bill_ref}_merged.pdf")
            except Exception as e:
                write_log(f"Error removing duplicate pages for bill '{bill_ref}': {e}")
                continue
            
            date_folder = str(df_bill['date_to_download'][0]).strip()
            dest_folder = Path(f"C:/Users/{username}/AP/cdas_merged_") / date_folder
            try:
                    move_pdf(path_pdf, dest_folder)
            except Exception as e:
                    print(f"Error moving PDF files for bill '{bill_ref}': {e}")
                    continue

            try:
                remove_recent_pdf_files(path_pdf)
            except Exception:
                  time.sleep(0.5)

            list_pdf_files_remove = get_pdf_files_with_invoice_number(path_pdf, bill_ref)
            for pdf_file in list_pdf_files_remove:
                    if pdf_file.lower().endswith('_merged.pdf'):
                        continue
                    try:
                            os.remove(os.path.join(path_pdf, pdf_file))
                    except FileNotFoundError:
                            pass
            #list_bill_ref_saved= [elem + "_bill" for elem in list_bill_ref_saved]
            all_invoices()
            time.sleep(0.5)
                   
    return list_bill_ref_saved            

if __name__ == '__main__':
    main(list_bill)    
   
    write_log(f"number of bills saved {len(list_bill_ref_saved)}")
    write_log(f"bill ref saved: {(list_bill_ref_saved)}")
    write_log(f"completed {datetime.now().strftime('%Y%m%d %H:%M:%S')} ...")
    #driver.quit()


