"""Creates an outlay ticket in OPUS from queue element."""
import json
import os
import time
from pynput.keyboard import Key, Controller
from mbu_dev_shared_components.utils.fernet_encryptor import Encryptor
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException
from OpenOrchestrator.database.queues import QueueStatus


def initialize_browser():
    """Initialize the Selenium Chrome WebDriver."""
    chrome_options = Options()
    prefs = {
        "safebrowsing.enabled": False
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("test-type")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-search-engine-choice")

    return webdriver.Chrome(options=chrome_options)


def click_element_with_retries(browser, by, value, retries=4):
    """Click an element with retries and handle common exceptions."""
    for attempt in range(retries):
        try:
            element = WebDriverWait(browser, 2).until(
                EC.element_to_be_clickable((by, value))
            )
            element.click()
            print(f"Successfully clicked element '{value}' on attempt {attempt + 1}")
            return True
        except ElementClickInterceptedException as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            time.sleep(1)

    # If all retries are exhausted an exception
    error_message = f"Failed to click element '{value}' after {retries} attempts"
    raise RuntimeError(error_message)


def decrypt_cpr(element_data):
    """Decrypt the CPR number from the element data."""
    encryptor = Encryptor()
    encrypted_cpr = element_data['cpr_encrypted']
    return encryptor.decrypt(encrypted_cpr.encode('utf-8'))


def handle_opus(browser, queue_element, path, orchestrator_connection):
    """Handle the OPUS ticket creation process."""
    element_data = json.loads(queue_element.data)
    attachment_path = os.path.join(path, f'receipt_{element_data["uuid"]}.pdf')

    try:
        navigate_to_opus(browser)
        fill_form(browser, element_data)
        upload_attachment(browser, attachment_path, orchestrator_connection)
        complete_form_and_submit(browser, element_data, orchestrator_connection)

        orchestrator_connection.log_trace("Successfully created outlay ticket.")

    except (RuntimeError) as e:
        orchestrator_connection.log_error(f"Error handling OPUS: {e}")
        orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.FAILED, f"Error handling OPUS: {e}")

    finally:
        browser.quit()
        os.remove(attachment_path)


def navigate_to_opus(browser):
    """Navigate to OPUS page and open required tabs."""
    browser.get("https://ssolaunchpad.kmd.dk/?kommune=1574&start=portal")
    wait_and_click(browser, By.XPATH, "//div[@class='TabText_SmallTabs' and text()='Min Økonomi']")
    wait_and_click(browser, By.XPATH, "//div[text()='Bilag og fakturaer']")
    wait_and_click(browser, By.XPATH, "/html/body/div[1]/table/tbody/tr[1]/td/div/div[1]/div[9]/div[2]/span[2]")


def fill_form(browser, element_data):
    """Fill out the form with data from element_data."""
    browser.switch_to.default_content()
    switch_to_frame(browser, 'contentAreaFrame')
    switch_to_frame(browser, 'ivuFrm_page0ivu0')
    enter_text(browser, 'WD9A', decrypt_cpr(element_data))  # Kreditor
    wait_and_click(browser, By.ID, 'WD9D')  # Hent button
    time.sleep(3)

    enter_text(browser, 'WDF6', element_data['posteringstekst'])  # Udbetalingstekst
    enter_text(browser, 'WD0112', element_data['posteringstekst'])  # Posteringstekst
    enter_text(browser, 'WD0119', element_data['reference'])  # Reference
    enter_text(browser, 'WD0123', element_data['beloeb'])  # Beløb
    enter_text(browser, 'WD0156', element_data['naeste_agent'])  # Næste agent


def upload_attachment(browser, attachment_path, orchestrator_connection):
    """Upload the attachment file to the browser form."""
    wait_and_click(browser, By.ID, 'WD0189')  # Click 'Vedhæft nyt' button
    WebDriverWait(browser, 10).until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )
    browser.switch_to.default_content()
    switch_to_frame(browser, 'URLSPW-0')
    wait_and_click(browser, By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/div/div[3]/table/tbody/tr/td/div/div/span/span[2]/form')  # Click 'Vælg fil' button

    # Upload the file
    if not os.path.isfile(attachment_path):
        orchestrator_connection.log_error("File not found")
        raise FileNotFoundError(f"File not found: {attachment_path}")

    time.sleep(3)
    keyboard = Controller()
    keyboard.type(attachment_path)
    time.sleep(1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)

    # Confirm attachment upload
    wait_and_click(browser, By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/div/div[4]/div/table/tbody/tr/td[3]/table/tbody/tr/td[1]/div')  # Click 'OK' button


def complete_form_and_submit(browser, element_data, orchestrator_connection):
    """Complete the form and submit the ticket."""
    browser.switch_to.default_content()
    switch_to_frame(browser, 'contentAreaFrame')
    switch_to_frame(browser, 'ivuFrm_page0ivu0')

    keyboard = Controller()

    wait_and_click(browser, By.ID, 'WD0222')
    keyboard.type(element_data['arts_konto'])  # Artskonto

    wait_and_click(browser, By.ID, 'WD0228')
    keyboard.type(element_data['beloeb'])  # Beløb

    wait_and_click(browser, By.ID, 'WD0239-r')
    keyboard.type(element_data['psp'])  # PSP

    wait_and_click(browser, By.ID, 'WD023F')
    keyboard.type(element_data['posteringstekst'])  # Posteringstekst

    wait_and_click(browser, By.ID, 'WD1E')  # Click 'Kontroller' button
    time.sleep(2)

    if not browser.find_elements(By.XPATH, "//*[contains(text(), 'Udgiftsbilag er kontrolleret og OK')]"):
        orchestrator_connection.log_error("Control check failed")
        raise RuntimeError("Control check failed.")

    print("Clicking the Opret button...")
    wait_and_click(browser, By.ID, 'WD1B')  # Click 'Opret' button
    orchestrator_connection.log_trace("Successfully clicked the created ticket")


def switch_to_frame(browser, frame):
    """Switch to the required frames to access the form."""
    WebDriverWait(browser, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, frame)))


def enter_text(browser, element_id, text):
    """Helper to enter text into a form element."""
    input_element = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.ID, element_id))
    )
    input_element.send_keys(text)


def wait_and_click(browser, by, value):
    """Wait for an element to be clickable, then click it."""
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((by, value)))
    click_element_with_retries(browser, by, value)
