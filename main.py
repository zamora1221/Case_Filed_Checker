import os
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import csv
from selenium.webdriver.chrome.options import Options
import tkinter as tk
from tkinter import filedialog
import threading
import tkinter.ttk as ttk
import sys
import openpyxl
import requests
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException
import re



class AnyOfTheseElementsLocated:
    def __init__(self, *locators):
        self.locators = locators

    def __call__(self, driver):
        for locator in self.locators:
            try:
                element = driver.find_element(*locator)
                print(f"Found element: {locator}")
                return element
            except NoSuchElementException:
                pass
        return False

URL = "https://portal-txguadalupe.tylertech.cloud/PublicAccess/JailingSearch.aspx?ID=500"


def read_names_from_xlsx(file_path):
    df = pd.read_excel(file_path)
    df = df.drop_duplicates()
    names = []
    suffixes = ["Jr.", "Sr.", "I", "II", "III"]

    for index, row in df.iterrows():
        if pd.notnull(row['People::Name Full']):
            full_name = row['People::Name Full'].strip().split()
            first_name = full_name[0]
            last_name = full_name[-1]
            # Check if the last name is in the suffixes list
            if last_name in suffixes and len(full_name) > 2:
                last_name = full_name[-2]
        else:
            first_name = ''
            last_name = ''

        if pd.isnull(row['People::D.O.B.']):
            dob = ''  # Assign an empty string if the D.O.B value is NaT
        else:
            dob = row['People::D.O.B.'].strftime('%m/%d/%Y')

        name = {
            'first_name': first_name,
            'last_name': last_name,
            'dob': dob
        }
        names.append(name)
    return names


def search_form(driver, last_name, first_name, dob=''):
    last_name_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "LastName")))
    first_name_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "FirstName")))
    search_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "SearchSubmit")))

    last_name_input.clear()
    first_name_input.clear()
    last_name_input.send_keys(last_name)
    first_name_input.send_keys(first_name)

    if dob:  # Only fill in dob_input if dob is not an empty string
        dob_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "DateOfBirth")))
        dob_input.clear()
        dob_input.send_keys(dob)

    search_button.click()


def get_criminal_case_records(driver, county, last_name, first_name, dob='', names=None):
    filed_cases = []
    search_url = {
        "Guadalupe": "https://portal-txguadalupe.tylertech.cloud/PublicAccess/default.aspx",
        "Comal": "http://public.co.comal.tx.us/default.aspx",  # Replace with the actual URL
        "Hays": "https://public.co.hays.tx.us/default.aspx"  # Replace with the actual URL
    }[county]

    driver.get(search_url)

    if county in ["Guadalupe", "Comal", "Hays"]:
        print("Looking for the Criminal Case Records link...")
        criminal_case_records_link = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.LINK_TEXT, "Criminal Case Records")))
        criminal_case_records_link.click()
        # Select "Defendant" from the drop-down menu
        if county == "Guadalupe":
            search_type_dropdown = Select(driver.find_element(By.ID, "SearchBy"))
            search_type_dropdown.select_by_visible_text("Defendant")

    search_form(driver, last_name, first_name, dob)

    print("Waiting for search results...")

    filed_div_locator = (By.XPATH, "//div[contains(text(), 'Filed')]")
    no_cases_matched_locator = (By.XPATH, "//span[contains(text(), 'No cases matched your search criteria.')]")

    html_content = None  # Define html_content variable with an initial value

    try:
        WebDriverWait(driver, 10).until(AnyOfTheseElementsLocated(filed_div_locator, no_cases_matched_locator))

        if has_filed_status(driver.page_source):
            case_number_links = driver.find_elements(By.XPATH,
                                                     '//tr[./td/div[text()="Filed"]]//a[contains(@href, "CaseDetail.aspx?CaseID=")]')

            if case_number_links:
                case_number_link = case_number_links[0]  # Get the first case link

                case_id = case_number_link.get_attribute('href').split('=')[-1]

                case_number_link.click()
                # Extract the court dates using regular expressions
                html_content = driver.page_source
                court_dates = []

                # Extract all court dates using regular expressions
                date_pattern = r"\d{2}/\d{2}/\d{4}"
                matches = re.findall(date_pattern, html_content)
                if matches:
                    court_dates = matches

                # Set the court date as the last date found
                court_date = court_dates[-1] if court_dates else None

                # Add the case details to the case list
                print(court_date)
                filed_cases.append({
                    'last_name': last_name,
                    'first_name': first_name,
                    'dob': dob,
                    'case_id': case_id,
                    'court_date': court_date
                })

            return filed_cases, True

    except StaleElementReferenceException:
        print("Stale element reference exception occurred. Retrying...")
        return get_criminal_case_records(driver, county, last_name, first_name, dob, names)

    except TimeoutException:
        print(f"No results found for {last_name}, {first_name}.")
        return None, False


def has_filed_status(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    no_cases_matched = soup.find("span", string="No cases matched your search criteria.")
    if no_cases_matched:
        return False

    filed_div = soup.find('div', string='Filed')
    return filed_div is not None

class TextRedirector(object):
    def __init__(self, widget):
        self.widget = widget

    def write(self, string):
        self.widget.insert(tk.END, string)
        self.widget.see(tk.END)

    def flush(self):
        pass

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Web Scraper")

        self.file_path_var = tk.StringVar()
        self.county_var = tk.StringVar()
        self.filed_cases_path_var = tk.StringVar()
        self.no_case_filed_path_var = tk.StringVar()

        self.county_label = tk.Label(root, text="County:")
        self.county_label.grid(row=1, column=0, padx=5, pady=5)

        self.county_combobox = ttk.Combobox(root, textvariable=self.county_var, values=["Guadalupe", "Comal", "Hays"])
        self.county_combobox.grid(row=1, column=1, padx=5, pady=5)

        self.file_path_label = tk.Label(root, text="File path:")
        self.file_path_label.grid(row=0, column=0, padx=5, pady=5)

        self.file_path_entry = tk.Entry(root, textvariable=self.file_path_var, width=50)
        self.file_path_entry.grid(row=0, column=1, padx=5, pady=5)

        self.browse_button = tk.Button(root, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=0, column=2, padx=5, pady=5)

        self.start_button = tk.Button(root, text="Start", command=self.start_scraper)
        self.start_button.grid(row=2, column=1, padx=5, pady=5)

        self.progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=3, column=0, columnspan=3, padx=5, pady=5)

        self.filed_cases_button = tk.Button(root, text="Save Filed Cases To...", command=self.browse_filed_cases_file)
        self.filed_cases_button.grid(row=2, column=2, padx=5, pady=5)

        self.no_case_filed_button = tk.Button(root, text="Save No Case Filed To...",
                                              command=self.browse_no_case_filed_file)
        self.no_case_filed_button.grid(row=3, column=2, padx=5, pady=5)

        self.console_output = tk.Text(root, height=10, width=50)
        self.console_output.grid(row=4, column=1, padx=5, pady=5)

        # Redirect stdout
        sys.stdout = TextRedirector(self.console_output)

    def write_filed_cases_to_csv(self, filed_cases, file_path):
        fieldnames = ["Last Name", "First Name", "D.O.B.", "Court Date"]

        with open(file_path, mode="w", newline="", encoding="utf-8") as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
            writer.writeheader()

            for case in filed_cases:
                row = {
                    "Last Name": case.get("last_name", ""),
                    "First Name": case.get("first_name", ""),
                    "D.O.B.": case.get("dob", ""),
                    "Court Date": case.get("court_date", "")
                }
                writer.writerow(row)

    def write_no_case_filed_to_csv(self, no_case_filed, file_path):
        with open(file_path, mode="w", newline="", encoding="utf-8") as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(["Last Name", "First Name", "D.O.B."])

            for case in no_case_filed:
                writer.writerow([case["last_name"], case["first_name"], case["dob"]])

    def browse_file(self):
        self.file_path_var.set(filedialog.askopenfilename())

    def start_scraper(self):
        # Run the scraper in a separate thread to prevent blocking the GUI
        threading.Thread(target=self.run_scraper, daemon=True).start()

    def browse_filed_cases_file(self):
        self.filed_cases_path_var.set(filedialog.asksaveasfilename(defaultextension=".csv"))

    def browse_no_case_filed_file(self):
        self.no_case_filed_path_var.set(filedialog.asksaveasfilename(defaultextension=".csv"))

    def run_scraper(self):
        county = self.county_var.get()
        file_path = self.file_path_var.get()
        chrome_options = Options()

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

        filed_cases = []
        no_case_filed = []

        # Read names from .xlsx file
        names = read_names_from_xlsx(file_path)

        for i, name in enumerate(names):
            html_content, has_filed = get_criminal_case_records(driver, county, name['last_name'], name['first_name'], name['dob'])



            if has_filed:
                filed_cases.append(name)
            else:
                no_case_filed.append(name)

            # Update the progress bar after each name
            self.progress['value'] = (i + 1) / len(names) * 100
            self.root.update_idletasks()

        # Writing the results to csv files
        self.write_filed_cases_to_csv(filed_cases, self.filed_cases_path_var.get())
        self.write_no_case_filed_to_csv(no_case_filed, self.no_case_filed_path_var.get())

        print("Scraping completed. Check the csv files for the results.")
        driver.quit()


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
