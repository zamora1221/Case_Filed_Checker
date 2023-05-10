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
    names = []
    for index, row in df.iterrows():
        if pd.notnull(row['Name']):
            full_name = row['Name'].strip().split()
            first_name = full_name[0]
            middle_name = full_name[1] if len(full_name) > 2 else ''
            last_name = full_name[-1]
        else:
            first_name = ''
            middle_name = ''
            last_name = ''

        if pd.isnull(row['D.O.B']):
            dob = ''  # Assign an empty string if the D.O.B value is NaT
        else:
            dob = row['D.O.B'].strftime('%m/%d/%Y')

        name = {
            'first_name': first_name,
            'middle_name': middle_name,
            'last_name': last_name,
            'dob': dob
        }
        names.append(name)
    return names


def write_filed_cases_to_csv(filed_cases):
    with open("filed_cases.csv", mode="w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(["Last Name", "First Name", "Middle Name"])

        for case in filed_cases:
            writer.writerow([case["last_name"], case["first_name"], case["middle_name"]])

def write_no_case_filed_to_csv(no_case_filed):
    with open("no_case_filed.csv", mode="w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(["Last Name", "First Name", "Middle Name"])

        for case in no_case_filed:
            writer.writerow([case["last_name"], case["first_name"], case["middle_name"]])

def search_form(driver, last_name, first_name, middle_name=None, dob=''):
    last_name_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "LastName")))
    first_name_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "FirstName")))
    middle_name_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MiddleName")))
    dob_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "DateOfBirth")))
    search_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "SearchSubmit")))

    last_name_input.clear()
    first_name_input.clear()
    middle_name_input.clear()
    dob_input.clear()
    last_name_input.send_keys(last_name)
    first_name_input.send_keys(first_name)
    dob_input.send_keys(dob)

    if middle_name:
        middle_name_input.send_keys(middle_name)

    search_button.click()
def parse_jail_records(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    div_elements = soup.find_all("div")
    records = []

    for div_element in div_elements:
        text = div_element.text.strip()
        if ',' in text:
            record = {
                'name': text,
            }
            records.append(record)

    return records

def get_criminal_case_records(driver, last_name, first_name, middle_name='', dob=''):
    search_url = "https://portal-txguadalupe.tylertech.cloud/PublicAccess/default.aspx"
    driver.get(search_url)
    print("Looking for the Criminal Case Records link...")
    criminal_case_records_link = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.LINK_TEXT, "Criminal Case Records")))
    criminal_case_records_link.click()
    print("Looking for the Defendant radio button...")
    # Select "Defendant" from the drop-down menu
    search_type_dropdown = Select(driver.find_element(By.ID, "SearchBy"))
    search_type_dropdown.select_by_visible_text("Defendant")
    search_form(driver, last_name, first_name, middle_name, dob)


    print("Waiting for search results...")

    filed_div_locator = (By.XPATH, "//div[contains(text(), 'Filed')]")
    no_cases_matched_locator = (By.XPATH, "//span[contains(text(), 'No cases matched your search criteria.')]")

    try:
        WebDriverWait(driver, 5).until(AnyOfTheseElementsLocated(filed_div_locator, no_cases_matched_locator))
        html_content = driver.page_source
        return html_content
    except TimeoutException:
        print(f"{last_name}, {first_name} has no case file.")
        return None

def has_filed_status(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    no_cases_matched = soup.find("span", string="No cases matched your search criteria.")
    if no_cases_matched:
        return False

    filed_div = soup.find('div', string='Filed')
    return filed_div is not None

def main():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    filed_cases = []
    no_case_filed = []

    # Read names from .xlsx file
    names = read_names_from_xlsx("guadrecords2 (1).xlsx")

    for name in names:
        first_name = name['first_name']
        middle_name = name.get('middle_name', '')
        last_name = name['last_name']
        dob = name['dob']

        print(f"\nProcessing {first_name} {middle_name} {last_name}...")
        html_content = get_criminal_case_records(driver, last_name, first_name, middle_name, dob)
        if html_content is None:
            no_case_filed.append({"last_name": last_name, "first_name": first_name, "middle_name": middle_name})
            print(f"Unable to retrieve criminal case records for {last_name}, {first_name}. Skipping...")
            continue

        if has_filed_status(html_content):
            print(f"Filed case found for {last_name}, {first_name}.")
            filed_cases.append({"last_name": last_name, "first_name": first_name, "middle_name": middle_name})
        else:
            print(f"No case filed for {last_name}, {first_name}.")
            no_case_filed.append({"last_name": last_name, "first_name": first_name, "middle_name": middle_name})

    print("\nWriting results to CSV files...")
    write_filed_cases_to_csv(filed_cases)
    write_no_case_filed_to_csv(no_case_filed)

    print("\nDone!")

    driver.quit()
if __name__ == "__main__":
    main()
