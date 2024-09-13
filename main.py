import time
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException, \
    ElementClickInterceptedException, NoSuchWindowException
from getpass import getpass
import pandas as pd
import sys
import re

HOSPITAL_NUMBER_COLUMN = 0
DRUG_NAME_COLUMN = 1
PATIENTS_COMPLETED = 0

def setup_browser():
    """Setup the bot, take username and password input from user"""
    username = input("EPMA username: ").upper()
    password = getpass('EPMA Password: ')

    headless_mode = input("Would you like to run in headless mode? (Y/N) ").upper()

    print("Opening Browser")

    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.browser_version = 'stable'

    if headless_mode == 'Y':
        options.add_argument('--headless=new')      # Runs in the background

    options.add_argument("--log-level=3")
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()

    driver.get('https://epma.nnuh.nhs.uk:50000/Account/Login')
    print("Opening EPMA\n")

    return driver, username, password

def read_data():
    """Reads in the Hospital Numbers and corresponding Drug Types.
       Groups drug types by hospital number."""
    print("Reading data from spreadsheet")

    try:
        file_loc = "Order_Drug_Suppression.xls"
        data_frame = pd.read_excel(file_loc, header=None, dtype=str)

        # Drop rows where both columns are NaN
        data_frame = data_frame.dropna(how="all", subset=[HOSPITAL_NUMBER_COLUMN, DRUG_NAME_COLUMN])

        # Filter to keep only rows where column 0 has a 7-digit integer
        #data_frame = data_frame[data_frame[HOSPITAL_NUMBER_COLUMN].str.match(r'^\d{7}$')]

        # Ensure drug types are uppercase
        data_frame[DRUG_NAME_COLUMN] = data_frame[DRUG_NAME_COLUMN].str.upper()
        data_frame = data_frame.replace(r'\s+', '', regex=True)

        # Group the data by column 0 (Hospital Numbers) and create a dictionary
        grouped_data = data_frame.groupby(HOSPITAL_NUMBER_COLUMN)[DRUG_NAME_COLUMN].apply(list).to_dict()

    except FileNotFoundError:
        print('\x1b[1;31;40m' + "ERROR: Spreadsheet could not be found. Spreadsheet name should be: 'Order_Drug_Suppression.xls'" + '\x1b[0m')
        sys.exit()

    except (pd.errors.EmptyDataError, AttributeError):
        print("The spreadsheet is empty or corrupted.")
        sys.exit()

    except IOError as e:
        print(f"An error occurred while reading the file: {e}")
        return

    print("Data read")
    print(grouped_data)

    number_of_patients = len(data_frame.drop_duplicates([HOSPITAL_NUMBER_COLUMN]))

    return grouped_data, number_of_patients

def nav_to_inpatient_finder(driver, username, password):
    """Inputs user credentials and navigates to the inpatient finder"""
    username_field = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.ID, "usr")))
    username_field.send_keys(username)

    password_field = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.ID, "pwd")))
    password_field.click()
    password_field.send_keys(password + Keys.ENTER)

    driver.get('https://epma.nnuh.nhs.uk:50000/PatientSearch/Inpatient')

def enter_patient_hospital_number(driver, hospital_number):
    """Searches for the given Hosp. Num. and then clicks on Patient Notes"""
    try:
        print(f"Searching hospital number: {hospital_number}")
        hospital_number_field = WebDriverWait(driver, 10).until(
            ec.visibility_of_element_located((By.XPATH, '//*[@id="hospitalNumber"]')))
        hospital_number_field.click()
        hospital_number_field.send_keys(hospital_number)
        hospital_number_field.send_keys(Keys.ENTER)

        #Detects if no results are found (i.e. patient has been discharged)
        no_results_found_warning_message = no_results_found(driver)

        #If there are no results, skip this patient
        if no_results_found_warning_message:
            print(f"No results for {hospital_number} found. Trying next patient.")
            return

        patient = WebDriverWait(driver, 20).until(ec.visibility_of_element_located(
            (By.XPATH, '//*[@id="view-search-patient"]/div[3]/div[2]/div/div[1]/div/div[1]/div')))
        patient.click()

        # Wait for the URL to possibly contain 'Inpatient?patientId'
        try:
            if WebDriverWait(driver, 3).until(ec.url_contains('Inpatient?patientId')):
                # URL contains 'Inpatient?patientId', proceed with the next steps
                print("Patient page loaded, opening patient notes.")
            else:
                # URL does not contain 'Inpatient?patientId' => there are similarly named patients
                similarly_named_patients(driver)

        except TimeoutException:
            # URL did not change in time, assume similarly named patients need handling
            similarly_named_patients(driver)

        # Otherwise, click on 'Patient Notes'
        WebDriverWait(driver, 10).until(
            ec.visibility_of_element_located((By.LINK_TEXT, 'PATIENT NOTES'))
        ).click()

    except ElementClickInterceptedException:
        if WebDriverWait(driver, 10).until(ec.url_contains('Inpatient?patientId')):
            WebDriverWait(driver, 10).until(
                ec.visibility_of_element_located((By.LINK_TEXT, 'PATIENT NOTES'))
            ).click()
        else:
            return

    except TimeoutException:
        return

def no_results_found(driver):
    """Detects 'No search results found.' warning"""
    try:
        element = WebDriverWait(driver, 2).until(
            ec.presence_of_element_located((By.XPATH, '//*[@id="x-msg-zone-default"]/div/p/span'))
        )
        if "No search results found." in element.text:
            return True

        return False

    except (NoSuchElementException, TimeoutException):
        return False

def similarly_named_patients(driver):
    """Checks if the similarly named patients warning appears and then selects the active patient."""
    try:
        # Wait for the warning message or similar indicator that similar patients exist
        warning_message = WebDriverWait(driver, 5).until(
            ec.visibility_of_element_located(
                (By.XPATH, "//span[contains(text(), 'similarly named patients')]")
            )
        )

        if warning_message:
            print("Similar patients warning detected.")
            # Wait for the active patient line element
            active_patient = WebDriverWait(driver, 5).until(
                ec.element_to_be_clickable(
                    (By.CSS_SELECTOR, '.x-patient-line.active')
                )
            )

            hosp_num = WebDriverWait(driver, 5).until(
                ec.visibility_of_element_located(
                    (By.CSS_SELECTOR, '#view-search-patient > div.main > div.x-content-results > div > div:nth-child(1) > div.x-patient-line.active > div.x-patient-details > div:nth-child(2) > div.col-xs-2 > span')
                )
            )
            print(f"Selecting hospital number {hosp_num.get_attribute("title")} from available options.")

            active_patient.click()
            return True

        return False

    except (NoSuchElementException, TimeoutException, StaleElementReferenceException):
        return False

def suppress_note(driver, drug_name):
    """Clicks Edit note, change title to SUPPRESSED, and suppresses the note."""
    try:
        # Suppress the note
        note_title = WebDriverWait(driver, 10).until(ec.element_to_be_clickable(
            (By.CSS_SELECTOR, 'li.note.active span.title')))

        # If the note has a drug link, print that, otherwise just print the note title
        if drug_name:
            print(f"Suppressing note with title: {note_title.text} and drug link: {drug_name}")
        else:
            print(f"Suppressing note with title: {note_title.text}")

        #Edit Button
        WebDriverWait(driver, 5).until(ec.visibility_of_element_located(
            (By.XPATH, '//*[@id="x-content-noteView-01"]/div[2]/div/button'))).click()

        # Hardcoded delay to get around the jiggery pokery that is the EPMA website
        sleep(0.5)

        # Click on the text field
        WebDriverWait(driver, 5).until(ec.visibility_of_element_located(
            (By.XPATH, '//*[@id="notePEtitle"]'))).click()

        sleep(0.5)

        # Clear on the text field
        WebDriverWait(driver, 5).until(ec.visibility_of_element_located(
            (By.XPATH, '//*[@id="notePEtitle"]'))).clear()

        # Click on the text field, to make sure it's still selected (jiggery pokery)
        WebDriverWait(driver, 5).until(ec.visibility_of_element_located(
            (By.XPATH, '//*[@id="notePEtitle"]'))).click()

        # Type 'SUPPRESSED' in text field
        WebDriverWait(driver, 5).until(ec.visibility_of_element_located(
            (By.XPATH, '//*[@id="notePEtitle"]'))).send_keys("SUPPRESSED")

        #Suppression Date Picker
        WebDriverWait(driver, 5).until(ec.visibility_of_element_located(
            (By.XPATH, '//*[@id="notePEsuppressionDate-img"]'))).click()

        #Select Date
        WebDriverWait(driver, 5).until(ec.visibility_of_element_located(
            (By.XPATH, '//*[@id="e-notePEsuppressionDate"]/div[2]'))).click()

        #Save Button
        WebDriverWait(driver, 5).until(ec.visibility_of_element_located(
            (By.XPATH, '//*[@id="x-form-PatientNotesEditor"]/div[5]/div[1]/button[2]'))).click()

        print("Note suppressed!\n")

        # Update notes list due to DOM update
        notes = find_order_drug_notes(driver)

        # Jiggery Pokery
        sleep(1)

        return notes

    except TimeoutException:
        suppress_note(driver, drug_name)

    except ElementClickInterceptedException:
        WebDriverWait(driver, 10).until(ec.element_to_be_clickable(
            (By.XPATH, '/*[@id="x-form-PatientNotesEditor"]/div[5]/div[1]/button[1]')))
        suppress_note(driver, drug_name)
        return

    except Exception as e:
        print(f"An error occurred while suppressing the note: {e}")

def order_link_exists(driver):
    """Checks if the note has an order link"""
    try:
        # Loo for an order link
        WebDriverWait(driver, 10).until(
            ec.visibility_of_element_located((By.XPATH, '//*[@id="x-content-noteView-01"]/div[1]/div/div[1]/div[1]/div/div[3]'))
        )
        return True
    except (NoSuchElementException, TimeoutException):
        # If a NoSuchElementException or TimeoutException => there is no order link
        return False

def find_order_drug_notes(driver):
    """Finds and returns visible **Order Drug** notes that are not suppressed."""
    try:
        # Locate all matching notes (including suppressed ones)
        notes = WebDriverWait(driver, 3).until(
            ec.presence_of_all_elements_located(
                (By.XPATH, '//li[contains(@class, "note") or contains(@class, "note active")]'      
                           '[.//span[@class="tag" and contains(text(), "**Order Drug**")]]'     
                           '[not(.//span[@class="title" and text()="SUPPRESSED"])]'     # Title of 'Suppressed'
                           '[not(contains(@style, "display: none"))]'       # Suppressed notes
                 )
            ))

        return notes

    except TimeoutException:
        return []  # Return an empty list if no notes are found

def find_notes_to_suppress(driver, drug_list):
    """Finds and suppresses any **Order Drug** notes in the Patient Notes page."""
    notes = find_order_drug_notes(driver)
    i = 0

    # If no notes are found, skip this patient
    if len(notes) == 0:
        print("No order drug notes to be suppressed.\n")
        return

    # Loop over all the notes until they have all been evaluated
    while i < len(notes):
        try:
            # Click on the note
            WebDriverWait(driver, 10).until(
                ec.element_to_be_clickable(notes[i])
            ).click()

            # Check if the order link exists and matches the drug list
            if not order_link_exists(driver):
                notes = suppress_note(driver, '')  # No order link => no drug, so we send '', which is nothing
                i = 0       # Note is removed from the list, list is updated, so we must update the iterator
                continue    # Continue to next iteration
            else:
                # Finds the drug name of the current note
                drug_name = driver.find_element(By.XPATH, '//*[@id="noteOrderLink"]').get_attribute('value').upper()
                drug_name = re.sub(r"\s+", "", drug_name, flags=re.UNICODE)

                # If the drug name is in the list of drugs associated with the current hospital number, suppress
                if drug_name in drug_list:
                    notes = suppress_note(driver, drug_name)
                    i = 0
                else:
                    print(f"Drug name {drug_name} not in drug list.")
                    i += 1      # No suppression has occurred, so move on to the next note in the list

        except TimeoutException:
            print(f"TimeoutException: Trying next patient.\n")
            return

        except StaleElementReferenceException:
            # In my testing, StaleElement = > the WebDriver is trying to act
            # before the page has loaded. So we simply reload the notes list and continue to the next iteration
            notes = find_order_drug_notes(driver)
            continue

        except ElementClickInterceptedException:
            # Similar to Stale exception. In this case we just wait, to try and get the page to load.
            # In the even these two exceptions reoccur, the timeout exception will handle that
            sleep(0.5)
            continue

def to_inpatient_finder(driver):
    """Clicks the Inpatient Finder button"""
    print("\nReturning to Inpatient Finder")
    driver.get('https://epma.nnuh.nhs.uk:50000/PatientSearch/Inpatient')

def populate_notes(driver):
    """ DEBUG Automatically create notes, to save the faff"""
    for i in range(5):
        try:
            WebDriverWait(driver, 10).until(
                ec.visibility_of_element_located(
                    (By.XPATH, '//*[@id="overlay-PatientNotes"]/div[3]/ul/li[1]/a'))).click()

            WebDriverWait(driver, 10).until(ec.element_to_be_clickable(
                (By.XPATH, '//*[@id="notePEtitle"]'))).click()

            WebDriverWait(driver, 10).until(ec.element_to_be_clickable(
                (By.XPATH, '//*[@id="notePEtitle"]'))).clear()

            WebDriverWait(driver, 10).until(ec.element_to_be_clickable(
                (By.XPATH, '//*[@id="notePEtitle"]'))).click()

            WebDriverWait(driver, 10).until(ec.element_to_be_clickable(
                (By.XPATH, '//*[@id="notePEtitle"]'))).send_keys("Suppress this!")

            WebDriverWait(driver, 10).until(ec.element_to_be_clickable(
                (By.XPATH, '//*[@id="notePEtitle"]'))).click()

            WebDriverWait(driver, 10).until(ec.element_to_be_clickable(
                (By.XPATH, '//*[@id="notePEtype"]'))).click()

            WebDriverWait(driver, 10).until(ec.element_to_be_clickable(
                (By.XPATH, '//*[@id="notePEtype"]'))).send_keys("*" + Keys.ENTER)

            WebDriverWait(driver, 10).until(ec.element_to_be_clickable(
                (By.XPATH, '//*[@id="x-form-PatientNotesEditor"]/div[3]/div[1]/div[1]/a/div/button[3]'))
            ).click()

            WebDriverWait(driver, 10).until(ec.element_to_be_clickable(
                (By.XPATH, '//*[@id="x-form-PatientNotesEditor"]/div[3]/div[1]/div[1]/a/div/button[3]'))
            ).send_keys('Suppress This!')

            save_button = WebDriverWait(driver, 10).until(ec.element_to_be_clickable(
                (By.XPATH, '//*[@id="x-form-PatientNotesEditor"]/div[5]/div[1]/button[2]')))
            save_button.click()

        except ElementClickInterceptedException:
            continue

        except TimeoutException:
            continue

def execute_suppressions(driver, hospital_data, number_of_patients):
    """Iterates over the hospital numbers and their associated drug types, searches for any **Order Drugs**"""
    global PATIENTS_COMPLETED
    for hospital_num, drug_list in hospital_data.items():
        enter_patient_hospital_number(driver, hospital_num)

        #populate_notes(driver)

        find_notes_to_suppress(driver, drug_list)

        PATIENTS_COMPLETED += 1
        print(f"{PATIENTS_COMPLETED} out of {number_of_patients} completed.")

        to_inpatient_finder(driver)

def main():
    """Main function"""
    hospital_data, number_of_patients = read_data()     # Returns a dict grouped by hospital number, with associated drugs

    driver, username, password = setup_browser()

    nav_to_inpatient_finder(driver, username, password)

    start = time.time()
    execute_suppressions(driver, hospital_data, number_of_patients)
    end = time.time()

    print(f"Time taken to execute suppressions: {int(end - start)} seconds")

    driver.close()

    print("All patients processed. Closing browser.")

if __name__ == "__main__":
    try:
        main()

    except (KeyboardInterrupt, NoSuchWindowException):
        print("\nStopping...")
        sys.exit()