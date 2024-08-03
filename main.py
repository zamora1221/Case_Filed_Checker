import os
import base64
import streamlit as st
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
import pandas as pd
import time
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import csv
from dateutil.parser import parse

st.title("Criminal Case Record Search")

uploaded_file = st.file_uploader("Choose a file")
county = st.selectbox("Select County", ["Guadalupe", "Comal", "Hays", "Williamson"])
start_button = st.button("Start")

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

def read_names_from_xlsx(file_path):
    df = pd.read_excel(file_path)
    df = df.drop_duplicates()
    names = []
    suffixes = ["Jr.", "Sr.", "I", "II", "III"]

    df['People::D.O.B.'] = pd.to_datetime(df['People::D.O.B.'], errors='coerce')

    for index, row in df.iterrows():
        if pd.notnull(row['People::Name Full']):
            full_name = row['People::Name Full'].strip().split()
            first_name = full_name[0]

            if len(full_name) == 2:
                last_name = full_name[-1]
            elif len(full_name) == 3 and len(full_name[1]) == 1:
                last_name = full_name[-1]
            else:
                if len(full_name) > 1 and full_name[-2] in suffixes:
                    last_name = " ".join(full_name[-3:-1])
                else:
                    last_name = " ".join(full_name[-3:])
        else:
            first_name = ''
            last_name = ''

        if pd.isnull(row['People::D.O.B.']):
            dob = ''
        else:
            dob = row['People::D.O.B.'].strftime('%m/%d/%Y')

        name = {'first_name': first_name, 'last_name': last_name, 'dob': dob}
        names.append(name)
    return names

def write_filed_cases_to_csv(filed_cases, file_path):
    with open(file_path, mode="w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(["People::Name Full", "People::D.O.B.", "Case Number", "Court Dates"])
        for case in filed_cases:
            full_name = "{} {}".format(case["first_name"], case["last_name"])
            court_dates_str = ', '.join(date for date in case['court_dates'] if date is not None)
            writer.writerow([full_name, case["dob"], case["case_number"], court_dates_str])

def write_no_case_filed_to_csv(no_case_filed, file_path):
    with open(file_path, mode="w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(["People::Name Full", "People::D.O.B."])
        for case in no_case_filed:
            full_name = "{} {}".format(case["first_name"], case["last_name"])
            writer.writerow([full_name, case["dob"]])

def search_form(driver, last_name, first_name, dob=''):
    last_name_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "LastName")))
    first_name_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "FirstName")))
    search_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "SearchSubmit")))

    last_name_input.clear()
    first_name_input.clear()
    last_name_input.send_keys(last_name)
    first_name_input.send_keys(first_name)

    if dob:
        dob_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "DateOfBirth")))
        dob_input.clear()
        dob_input.send_keys(dob)

    search_button.click()

def get_criminal_case_records(driver, county, last_name, first_name, filed_cases, no_case_filed, dob=''):
    search_url = {
        "Guadalupe": "https://portal-txguadalupe.tylertech.cloud/PublicAccess/default.aspx",
        "Comal": "http://public.co.comal.tx.us/default.aspx",
        "Hays": "https://public.co.hays.tx.us/default.aspx",
        "Williamson": "https://judicialrecords.wilco.org/PublicAccess/default.aspx"  # Added Williamson County
    }[county]

    driver.get(search_url)
    time.sleep(2)

    # Use the same logic for Williamson as Hays
    if county in ["Guadalupe", "Comal", "Hays", "Williamson"]:  # Include Williamson here
        print("Looking for the Criminal Case Records link...")
        for _ in range(5):
            try:
                criminal_case_records_link = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.LINK_TEXT, "Criminal Case Records")))
                criminal_case_records_link.click()
                break
            except TimeoutException:
                print("Timed out waiting for 'Criminal Case Records' link, retrying...")
                driver.refresh()

        # Specific code for Guadalupe
        if county == "Guadalupe":
            for _ in range(5):  # Try up to 3 times
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "SearchBy")))
                    break  # Break the loop if we succeed
                except TimeoutException:
                    print("Timed out waiting for element with ID 'SearchBy', refreshing...")
                    driver.refresh()
            search_type_dropdown = Select(driver.find_element(By.ID, "SearchBy"))
            search_type_dropdown.select_by_visible_text("Defendant")

    search_form(driver, last_name, first_name, dob)
    case_record = {'first_name': first_name, 'last_name': last_name, 'dob': dob, 'court_dates': [], 'case_number': ''}
    print("Waiting for search results...")

    filed_div_locator = (By.XPATH, "//div[contains(text(), 'Filed')]")
    no_cases_matched_locator = (By.XPATH, "//span[contains(text(), 'No cases matched your search criteria.')]")

    try:
        WebDriverWait(driver, 10).until(AnyOfTheseElementsLocated(filed_div_locator, no_cases_matched_locator))
        html_content = driver.page_source

        if has_filed_status(html_content):
            soup = BeautifulSoup(html_content, 'html.parser')
            table_rows = soup.find_all('tr')
            latest_court_dates = []

            for row in table_rows:
                if row.find('div', string='Filed'):
                    case_number_link = row.find('a', href=True, style="color: blue")
                    if case_number_link and "CaseDetail.aspx?" in case_number_link['href']:
                        case_number_url = case_number_link['href']
                        case_number = case_number_link.text
                        case_record['case_number'] = case_number
                        print(f"Clicking on case number: {case_number}")

                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, f"//a[@href='{case_number_url}']"))).click()
                        # Wait for page load and parse it
                        # Locate the table with case details and get last date
                        latest_court_date = get_latest_court_date(driver.page_source)
                        print(f"Latest Court Date: {latest_court_date}")
                        case_record['court_dates'].append(latest_court_date)
                        # Go back to the search results page to find the next case
                        driver.back()
                        time.sleep(2)

            if case_record['court_dates']:
                # If we did, return the record and True
                return case_record, True, None
            else:
                # If we didn't, print a message and return None and False
                print(f"{last_name}, {first_name} is not filed.")
                time.sleep(2)
                return None, False, None
    except TimeoutException:
        print(f"{last_name}, {first_name} is not filed.")
        return None, False, None
    return None, False, None

def get_latest_court_date(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    court_dates = soup.find_all("th", {"class": "ssTableHeaderLabel", "valign": "top"})
    if court_dates:
        # Parse dates and ignore any invalid ones.
        parsed_dates = []
        for date in court_dates:
            try:
                parsed_date = parse(date.text.strip())
                parsed_dates.append(parsed_date)
            except ValueError:
                continue

        # If any valid dates were found, return the latest one.
        if parsed_dates:
            return max(parsed_dates).strftime('%m/%d/%Y')
    return None

def get_case_number(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    case_number_tag = soup.find("a", {"style": "color: blue"})
    if case_number_tag:
        case_number = case_number_tag.text.strip()
        return case_number
    return None

def has_filed_status(html_content):
    return "Filed" in html_content

def download_csv(file_path):
    with open(file_path, 'rb') as f:
        data = f.read()
    b64 = base64.b64encode(data).decode('UTF-8')
    href = f'<a href="data:file/csv;base64,{b64}" download="{file_path}">Download CSV file</a>'
    st.markdown(href, unsafe_allow_html=True)

if uploaded_file is not None and county and start_button:
    file_path = os.path.join(os.getcwd(), uploaded_file.name)
    with open(file_path, 'wb') as f:
        f.write(uploaded_file.getbuffer())

    st.write(f"File uploaded successfully: {uploaded_file.name}")
    st.write("Starting the process...")

    firefox_options = webdriver.FirefoxOptions()
    firefox_options.add_argument("--headless")
    driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()), options=firefox_options)

    filed_cases = []
    no_case_filed = []
    names = read_names_from_xlsx(file_path)

    total_names = len(names)
    progress_increment = 100 / total_names if total_names else 1
    current_progress = 0

    progress_bar = st.progress(0)

    for name in names:
        first_name = name['first_name']
        last_name = name['last_name']
        dob = name['dob']

        case_record, is_filed, _ = get_criminal_case_records(driver, county, last_name, first_name, filed_cases, no_case_filed, dob)
        if is_filed:
            filed_cases.append(case_record)
        else:
            no_case_filed.append(name)

        current_progress += progress_increment
        progress_bar.progress(int(current_progress))

    driver.quit()
    write_filed_cases_to_csv(filed_cases, 'filed_cases.csv')
    write_no_case_filed_to_csv(no_case_filed, 'no_case_filed.csv')

    st.write('Process complete.')
    st.write(f'Filed cases: {len(filed_cases)}')
    st.write(f'No case filed: {len(no_case_filed)}')

    st.markdown("### Download filed cases")
    download_csv('filed_cases.csv')

    st.markdown("### Download no case filed")
    download_csv('no_case_filed.csv')

else:
    st.write('Upload a file, select a county, and click Start.')

