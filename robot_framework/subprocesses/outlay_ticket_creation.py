"""This module contains the logic for creating an outlay ticket in OPUS."""
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


def initialize_browser():
    """Initialize the Selenium Chrome WebDriver."""
    chrome_options = Options()
    prefs = {
        "safebrowsing.enabled": False
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("test-type")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-web-security")
    chrome_options.add_argument("--allow-running-insecure-content")
    chrome_options.add_argument("--disable-search-engine-choice-screen")

    return webdriver.Chrome(options=chrome_options)


def click_element_with_retries(browser, by, value, retries=4):
    """Click an element with retries and handle common exceptions."""
    for attempt in range(retries):
        try:
            element = WebDriverWait(browser, 2).until(
                EC.element_to_be_clickable((by, value))
            )
            element.click()
            return True
        except Exception as e:  # pylint: disable=broad-except
            print(f"Attempt {attempt + 1} failed: {e}")
            time.sleep(1)
    return False


def decrypt_cpr(element_data):
    """Decrypt the CPR number from the element data."""
    encryptor = Encryptor()
    encrypted_cpr = element_data['cpr_encrypted']
    return encryptor.decrypt(encrypted_cpr.encode('utf-8'))


def handle_opus(queue_element, path, browser, orchestrator_connection):
    """Handle the OPUS ticket creation process."""

    element_data = json.loads(queue_element.data)
    attachment_path = os.path.join(path, f'receipt_{element_data["uuid"]}.pdf')

    navigate_to_opus(browser)
    fill_form(browser, element_data)
    upload_attachment(browser, attachment_path)

    complete_form_and_submit(browser, element_data)

    orchestrator_connection.log_trace("Successfully created outlay ticket.")
    print("Successfully created outlay ticket.")


def navigate_to_opus(browser):
    """Navigate to OPUS page and open required tabs."""
    browser.get("https://ssolaunchpad.kmd.dk/?kommune=1574&start=portal")
    WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, "//div[@class='TabText_SmallTabs' and text()='Min Økonomi']")))
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


def upload_attachment(browser, attachment_path):
    """Upload the attachment file to the browser form."""
    wait_and_click(browser, By.ID, 'WD0189')  # Click 'Vedhæft nyt' button
    WebDriverWait(browser, 20).until(
        lambda driver: driver.execute_script("return document.readyState") == "complete"
    )
    browser.switch_to.default_content()
    switch_to_frame(browser, 'URLSPW-0')
    wait_and_click(browser, By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/div/div[3]/table/tbody/tr/td/div/div/span/span[2]/form')  # Click 'Vælg fil' button

    time.sleep(4)
    keyboard = Controller()
    keyboard.type(attachment_path)
    time.sleep(2)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)

    wait_and_click(browser, By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/div/div[4]/div/table/tbody/tr/td[3]/table/tbody/tr/td[1]/div')  # Click 'OK' button
    time.sleep(2)


def press_key(keyboard, key):
    """Press and release a key on the keyboard."""
    keyboard.press(key)
    keyboard.release(key)


def complete_form_and_submit(browser, element_data):
    """Complete the form and submit the ticket."""

    from robot_framework.exceptions import BusinessError

    browser.switch_to.default_content()
    switch_to_frame(browser, 'contentAreaFrame')
    switch_to_frame(browser, 'ivuFrm_page0ivu0')

    keyboard = Controller()

    wait_and_click(browser, By.ID, 'WD0222')
    keyboard.type(element_data['arts_konto'])  # Artskonto

    press_key(keyboard, Key.tab)
    keyboard.type(element_data['beloeb'])  # Beløb

    press_key(keyboard, Key.tab)
    press_key(keyboard, Key.tab)
    press_key(keyboard, Key.tab)
    keyboard.type(element_data['psp'])  # PSP

    press_key(keyboard, Key.tab)
    keyboard.type(element_data['posteringstekst'])  # Posteringstekst

    time.sleep(1)

    wait_and_click(browser, By.ID, 'WD1E')  # Click 'Kontroller' button
    time.sleep(4)

    # Check for business error here
    if not browser.find_elements(By.XPATH, "//*[contains(text(), 'Udgiftsbilag er kontrolleret og OK')]"):
        raise BusinessError("Fejl ved kontrol af udgiftsbilag.")

    wait_and_click(browser, By.ID, 'WD1B')  # Click 'Opret' button
    time.sleep(4)
    if not browser.find_elements(By.XPATH, "//*[contains(text(), 'er oprettet')]"):
        time.sleep(1)
        raise BusinessError("Fejl ved oprettelse af udgiftsbilag, kontrol OK.")


def switch_to_frame(browser, frame):
    """Switch to the required frames to access the form."""
    WebDriverWait(browser, 30).until(EC.frame_to_be_available_and_switch_to_it((By.ID, frame)))


def enter_text(browser, element_id, text):
    """Helper to enter text into a form element."""
    input_element = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((By.ID, element_id))
    )
    input_element.send_keys(text)


def wait_and_click(browser, by, value):
    """Wait for an element to be clickable, then click it."""

    WebDriverWait(browser, 30).until(EC.presence_of_element_located((by, value)))
    click_element_with_retries(browser, by, value)
