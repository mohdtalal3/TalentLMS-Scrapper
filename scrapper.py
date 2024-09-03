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

# Airtable Configurations
AIRTABLE_API_KEY = st.secrets["AIRTABLE_API_KEY"]
BASE_ID = st.secrets["BASE_ID"]
TABLE_NAME = st.secrets["TABLE_NAME"]

def process_course(driver, wait, course_id):
    st.info(f"Navigating to course page {course_id}...")
    driver.get(f'https://espaceformations-monparcoursenligne.talentlms.com/reports/courseinfo/id:{course_id}')
    time.sleep(5)
    st.success("Navigation complete.")

    st.info("Getting download URL...")
    download_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tl-export-course"]')))
    download_button.click()
    time.sleep(5)
    download_url = driver.find_element(By.XPATH, '//*[@id="tl-export-course"]').get_attribute('href')
    st.success(f"Download URL for course {course_id} obtained successfully.")

    cookies = driver.get_cookies()
    cookie_dict = {cookie['name']: cookie['value'] for cookie in cookies}

    st.info(f"Downloading and processing file for course {course_id}...")
    response = requests.get(download_url, cookies=cookie_dict)
    
    if response.status_code == 200:
        df = pd.read_excel(io.BytesIO(response.content), sheet_name='Utilisateurs')
        st.success(f"File for course {course_id} downloaded and processed successfully.")
        
        st.subheader(f"Downloaded Data for Course {course_id}:")
        st.dataframe(df)

        return df
    else:
        st.error(f"Unable to download the file for course {course_id}. Please check your credentials and download URL.")
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

        course_ids = [239, 281]
        all_data = pd.DataFrame()

        for course_id in course_ids:
            df = process_course(driver, wait, course_id)
            if df is not None:
                all_data = pd.concat([all_data, df], ignore_index=True)

        if not all_data.empty:
            st.info("Updating Airtable records...")
            table = Table(AIRTABLE_API_KEY, BASE_ID, TABLE_NAME)

            updated_records = 0
            for index, row in all_data.iterrows():
                email = row['Email']
                statut = str(row['Statut']).strip()
                update_date = datetime.now().strftime("%Y-%m-%d")
                
                date_fin_cours = row['Date de fin du cours']
                if date_fin_cours and date_fin_cours != '-':
                    if isinstance(date_fin_cours, datetime):
                        date_fin_cours = date_fin_cours.strftime("%Y-%m-%d")
                    elif isinstance(date_fin_cours, str):
                        try:
                            parsed_date = datetime.strptime(date_fin_cours, "%Y-%m-%d")
                            date_fin_cours = parsed_date.strftime("%Y-%m-%d")
                        except ValueError:
                            date_fin_cours = None
                    else:
                        date_fin_cours = None
                else:
                    date_fin_cours = None
                temps = row['Temps']
                note_moyenne = row['Note moyenne']
                
                records = table.all(formula=f"{{Email}} = '{email}'")
                if records:
                    record_id = records[0]['id']
                    table.update(record_id, {
                        'Progress': statut,
                        'Date de mise à jour elearning': update_date,
                        'Temps elearning': temps,
                        'Note moy. elearning':note_moyenne,
                        'Date de fin elearning':date_fin_cours
                    })
                    updated_records += 1

            st.success(f"Script executed successfully. {updated_records} records updated in Airtable.")

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