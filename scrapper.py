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
from pyairtable import Table, Api
from pyairtable.formulas import match
import io

# Airtable Configurations
AIRTABLE_API_KEY = st.secrets["AIRTABLE_API_KEY"]
BASE_ID = st.secrets["BASE_ID"]
TABLE_NAME = st.secrets["TABLE_NAME"]

# Initialize Airtable API
api = Api(AIRTABLE_API_KEY)
table = Table(AIRTABLE_API_KEY, BASE_ID, TABLE_NAME)

def ensure_fields_exist(base_id, table_id, new_fields):
    url = f"https://api.airtable.com/v0/meta/bases/{base_id}/tables/{table_id}/fields"
    headers = {
        "Authorization": f"Bearer {AIRTABLE_API_KEY}",
        "Content-Type": "application/json"
    }

    existing_fields = table.fields

    for field in new_fields:
        if field not in existing_fields:
            payload = {
                "name": field,
                "type": "singleLineText",
                "description": f"Automatically created field for {field}"
            }
            
            response = requests.post(url, headers=headers, json=payload)
            
            if response.status_code == 200:
                st.success(f"Created new field: {field}")
            else:
                st.warning(f"Could not create field {field}. Status code: {response.status_code}, Response: {response.text}")

# New fields to add
new_fields = ['Date de fin du cours', 'Temps', 'Note moyenne']

def execute_script():
    try:
        # Ensure new fields exist
        ensure_fields_exist(BASE_ID, TABLE_NAME, new_fields)

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

            updated_records = 0
            for index, row in df.iterrows():
                email = row['Email']
                statut = str(row['Statut']).strip()
                update_date = datetime.now().strftime("%Y-%m-%d")
                
                # New fields
                date_fin_cours = row['Date de fin du cours']
                temps = row['Temps']
                note_moyenne = row['Note moyenne']
                
                records = table.all(formula=match({"Email": email}))
                if records:
                    record_id = records[0]['id']
                    table.update(record_id, {
                        'Progress': statut,
                        'Date de mise à jour elearning': update_date,
                        'Date de fin du cours': date_fin_cours,
                        'Temps': temps,
                        'Note moyenne': note_moyenne
                    })
                    updated_records += 1

            st.success(f"Script executed successfully. {updated_records} records updated in Airtable.")

        else:
            st.error("Unable to download the file. Please check your credentials and download URL.")

    except Exception as e:
        st.error(f"An error occurred: {str(e)}")

    finally:
        st.info("Closing WebDriver....")
        if 'driver' in locals():
            driver.quit()
        st.success("WebDriver closed successfully.")

# Streamlit interface
st.title("Execute TalentLMS Script")

if st.button("Execute Script"):
    execute_script()