import streamlit as st
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
from datetime import datetime
import pandas as pd
from pyairtable import Table
import io
from datetime import datetime, timedelta
# Airtable Configurations
AIRTABLE_API_KEY = st.secrets["AIRTABLE_API_KEY"]
BASE_ID = st.secrets["BASE_ID"]
TABLE_NAME = st.secrets["TABLE_NAME"]
def excel_date_to_iso(excel_date):
    try:
        excel_epoch = datetime(1899, 12, 30)
        delta_days = timedelta(days=float(excel_date))
        return excel_epoch + delta_days
    except:
        return None
def execute_script():
    try:
        st.info("Initializing Chrome WebDriver...")
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--window-size=1920,1080")
        driver = webdriver.Chrome(options=chrome_options)
        st.success("Chrome WebDriver initialized successfully.")

        st.info("Accessing TalentLMS...")
        driver.get('https://espaceformations-monparcoursenligne.talentlms.com/')
        wait = WebDriverWait(driver, 10)
        st.success("TalentLMS accessed successfully.")

        st.info("Logging in...")
        username = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="login"]')))
        password = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="password"]')))
        username.send_keys(st.secrets["TALENTLMS_USERNAME"])
        password.send_keys(st.secrets["TALENTLMS_PASSWORD"])
        login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="submit"]')))
        login_button.click()
        time.sleep(5)
        st.success("Logged in successfully.")

        st.info("Navigating to specific page...")
        driver.get('https://espaceformations-monparcoursenligne.talentlms.com/reports/courseinfo/id:239')
        time.sleep(5)
        st.success("Navigation complete.")

        st.info("Getting download URL...")
        download_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tl-export-course"]')))
        download_button.click()
        time.sleep(5)
        download_url = driver.find_element(By.XPATH, '//*[@id="tl-export-course"]').get_attribute('href')
        st.success(download_url)
        cookies = driver.get_cookies()
        cookie_dict = {cookie['name']: cookie['value'] for cookie in cookies}
        st.success("Download URL obtained successfully.")

        st.info("Downloading and processing file...")
        response = requests.get(download_url, cookies=cookie_dict)
        
        if response.status_code == 200:
            df = pd.read_excel(io.BytesIO(response.content), sheet_name='Utilisateurs')
            st.success("File downloaded and processed successfully.")
            
            st.subheader("Downloaded Data:")
            st.dataframe(df)

            st.info("Updating Airtable records...")
            table = Table(AIRTABLE_API_KEY, BASE_ID, TABLE_NAME)

            updated_records = 0
            for index, row in df.iterrows():
                email = row['Email']
                statut = str(row['Statut']).strip()
                update_date = datetime.now().strftime("%Y-%m-%d")
                
                # New fields
                excel_datetime = excel_date_to_iso(row['Date de fin du cours'])
                if excel_datetime:
                    formatted_end_date = excel_datetime.strftime('%Y-%m-%d %H:%M:%S')
                else:
                    formatted_end_date = None
                temps = row['Temps']
                note_moyenne = row['Note moyenne']
                st.write(formatted_end_date,temps,note_moyenne)
                # records = table.all(formula=f"{{Email}} = '{email}'")
                # if records:
                #     record_id = records[0]['id']
                #     table.update(record_id, {
                #         'Progress': statut,
                #         'Date de mise à jour elearning': update_date,
                #         'Date de fin du cours': formatted_end_date,
                #         'Temps': temps,
                #         'Note moyenne': note_moyenne
                #     })
                #     updated_records += 1

            st.success(f"Script executed successfully. {updated_records} records updated in Airtable.")

        else:
            st.error("Unable to download the file. Please check your credentials and download URL.")

    except Exception as e:
        st.error(f"An error occurred: {str(e)}")

    finally:
        st.info("Closing WebDriver....")
        driver.quit()
        st.success("WebDriver closed successfully.")

# Streamlit interface
st.title("Execute TalentLMS Script")

if st.button("Execute Script"):
    execute_script()