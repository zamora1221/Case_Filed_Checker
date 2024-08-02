import streamlit as st
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import io


def read_names_from_xlsx(file):
    df = pd.read_excel(file)
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
        dob = row['People::D.O.B.'].strftime('%m/%d/%Y') if pd.notnull(row['People::D.O.B.']) else ''
        names.append({'first_name': first_name, 'last_name': last_name, 'dob': dob})
    return names


def get_criminal_case_records(driver, county, last_name, first_name, filed_cases, no_case_filed, dob=''):
    search_url = {
        "Guadalupe": "https://portal-txguadalupe.tylertech.cloud/PublicAccess/default.aspx",
        "Comal": "http://public.co.comal.tx.us/default.aspx",
        "Hays": "https://public.co.hays.tx.us/default.aspx",
        "Williamson": "https://judicialrecords.wilco.org/PublicAccess/default.aspx"
    }[county]
    driver.get(search_url)
    time.sleep(2)
    if county in ["Guadalupe", "Comal", "Hays", "Williamson"]:
        try:
            criminal_case_records_link = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.LINK_TEXT, "Criminal Case Records")))
            criminal_case_records_link.click()
        except TimeoutException:
            st.error("Timed out waiting for 'Criminal Case Records' link.")
            return None, False, None
    search_form(driver, last_name, first_name, dob)
    case_record = {'first_name': first_name, 'last_name': last_name, 'dob': dob, 'court_dates': [], 'case_number': ''}
    filed_div_locator = (By.XPATH, "//div[contains(text(), 'Filed')]")
    no_cases_matched_locator = (By.XPATH, "//span[contains(text(), 'No cases matched your search criteria.')]")
    try:
        WebDriverWait(driver, 10).until(AnyOfTheseElementsLocated(filed_div_locator, no_cases_matched_locator))
        html_content = driver.page_source
        if has_filed_status(html_content):
            soup = BeautifulSoup(html_content, 'html.parser')
            table_rows = soup.find_all('tr')
            for row in table_rows:
                if row.find('div', string='Filed'):
                    case_number_link = row.find('a', href=True, style="color: blue")
                    if case_number_link and "CaseDetail.aspx?" in case_number_link['href']:
                        case_number_url = case_number_link['href']
                        case_number = case_number_link.text
                        case_record['case_number'] = case_number
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, f"//a[@href='{case_number_url}']"))).click()
                        latest_court_date = get_latest_court_date(driver.page_source)
                        case_record['court_dates'].append(latest_court_date)
                        driver.back()
                        time.sleep(2)
            if case_record['court_dates']:
                return case_record, True, None
            else:
                return None, False, None
    except TimeoutException:
        return None, False, None
    return None, False, None


def get_latest_court_date(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    court_dates = soup.find_all("th", {"class": "ssTableHeaderLabel", "valign": "top"})
    parsed_dates = []
    for date in court_dates:
        try:
            parsed_date = parse(date.text.strip())
            parsed_dates.append(parsed_date)
        except ValueError:
            continue
    if parsed_dates:
        return max(parsed_dates).strftime('%m/%d/%Y')
    return None


def has_filed_status(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    no_cases_matched = soup.find("span", string="No cases matched your search criteria.")
    return not no_cases_matched and soup.find('div', string='Filed') is not None


def main():
    st.title('Criminal Case Record Scraper')
    uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")
    county = st.selectbox("Select County", ["Guadalupe", "Comal", "Hays", "Williamson"])

    if uploaded_file and county:
        st.write("Processing...")
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        filed_cases = []
        no_case_filed = []
        names = read_names_from_xlsx(uploaded_file)
        for name in names:
            case_record, has_filed, _ = get_criminal_case_records(driver, county, name['last_name'], name['first_name'],
                                                                  filed_cases, no_case_filed, name['dob'])
            if has_filed:
                filed_cases.append(case_record)
            else:
                no_case_filed.append(name)
        driver.quit()
        filed_cases_df = pd.DataFrame(filed_cases)
        no_case_filed_df = pd.DataFrame(no_case_filed)
        st.write("Scraping completed.")
        st.download_button("Download Filed Cases", filed_cases_df.to_csv(index=False), file_name="filed_cases.csv")
        st.download_button("Download No Cases Filed", no_case_filed_df.to_csv(index=False),
                           file_name="no_cases_filed.csv")


if __name__ == "__main__":
    main()
