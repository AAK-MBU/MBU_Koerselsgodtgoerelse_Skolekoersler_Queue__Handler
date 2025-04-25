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
from selenium.webdriver.common.action_chains import ActionChains


def initialize_browser(opus_username, opus_password):
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
    chrome_options.add_argument("--incognito")

    browser = webdriver.Chrome(options=chrome_options)

    login_to_opus(browser, opus_username, opus_password)

    return browser


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


def login_to_opus(browser, username, password):
    """Login to OPUS."""
    browser.get("https://portal.kmd.dk/irj/portal")
    wait_and_click(browser, By.ID, 'logonuidfield')
    enter_text(browser, By.ID, 'logonuidfield', {username})
    enter_text(browser, By.ID, 'logonpassfield', {password})
    wait_and_click(browser, By.ID, 'buttonLogon')


def navigate_to_opus(browser):
    """Navigate to OPUS page and open required tabs."""
    browser.get("https://portal.kmd.dk/irj/portal")
    wait_and_click(browser, By.XPATH, "//div[text()='Min Økonomi']")
    wait_and_click(browser, By.XPATH, "//div[text()='Bilag og fakturaer']")
    wait_and_click(browser, By.XPATH, "/html/body/div[1]/table/tbody/tr[1]/td/div/div[1]/div[9]/div[2]/span[2]")


def fill_form(browser, element_data):
    """Fill out the form with data from element_data."""
    browser.switch_to.default_content()
    switch_to_frame(browser, 'contentAreaFrame')
    switch_to_frame(browser, 'ivuFrm_page0ivu0')
    root_xpath = "/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/div/div/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[1]/div/div/table/tbody/tr/td/div/div/table/tbody/"
    enter_text(
        browser,
        By.XPATH,
        root_xpath + "tr[2]/td/div/div/table/tbody/tr/td[1]/div/div/table/tbody/tr[1]/td[2]/div/div/table/tbody/tr/td[1]/span/input",
        decrypt_cpr(element_data),
    )  # Kreditor
    wait_and_click(
        browser,
        By.XPATH,
        root_xpath + "tr[2]/td/div/div/table/tbody/tr/td[1]/div/div/table/tbody/tr[1]/td[2]/div/div/table/tbody/tr/td[2]/div",
    )  # Hent button
    time.sleep(3)

    enter_text(
        browser,
        By.XPATH,
        root_xpath + "tr[3]/td/div/div/table/tbody/tr[1]/td[1]/div/div/table/tbody/tr/td/div/div/table/tbody/tr[1]/td[2]/span/input",
        element_data["posteringstekst"],
    )  # Udbetalingstekst
    enter_text(
        browser,
        By.XPATH,
        root_xpath + "tr[3]/td/div/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td[2]/span/input",
        element_data["posteringstekst"],
    )  # Posteringstekst
    enter_text(
        browser,
        By.XPATH,
        root_xpath + "tr[3]/td/div/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[3]/td[2]/span/input",
        element_data["reference"],
    )  # Reference
    enter_text(
        browser,
        By.XPATH,
        root_xpath + "tr[3]/td/div/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[4]/td[2]/div/div/table/tbody/tr/td[1]/span/input",
        element_data["beloeb"],
    )  # Beløb
    enter_text(
        browser,
        By.XPATH,
        root_xpath + "tr[4]/td/div/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr[1]/td[1]/span/input",
        element_data["naeste_agent"],
    )  # Næste agent

    # Click item next to "udbeatlingstekst" to add column with child name
    wait_and_click(
        browser,
        By.XPATH,
        root_xpath + "tr[3]/td/div/div/table/tbody/tr[1]/td[1]/div/div/table/tbody/tr/td/div/div/table/tbody/tr[1]/td[3]/div",
    )
    browser.switch_to.default_content()  # Popup is not appearing on current frame
    switch_to_frame(browser, "URLSPW-0")  # Switch to popup
    # Type text at cursor (element id is dynamic but cursor always starts at next empty line)
    actions = ActionChains(browser)
    actions.send_keys(element_data["barnets_navn"])
    actions.perform()
    # Click "Gem"
    # Find all buttons in frame:
    buttons = browser.find_elements(By.CLASS_NAME, "lsButton")
    # Search and click on "gem"
    for button in buttons:
        if button.text.lower() == "gem":
            button.click()
    # Back to previous frame
    browser.switch_to.default_content()
    switch_to_frame(browser, "contentAreaFrame")
    switch_to_frame(browser, "ivuFrm_page0ivu0")


def upload_attachment(browser, attachment_path):
    """Upload the attachment file to the browser form."""
    wait_and_click(browser, By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/div/div/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/table/tbody/tr/td/div/table/tbody/tr[3]/td/div/span/span/div/span/span[1]/table/thead/tr[2]/th/div/div/div/span/div')  # Click 'Vedhæft nyt' button
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

    wait_and_click(browser, By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/div/div/div/table/tbody/tr[2]/td/div/span/span[1]/div/span/span[1]/div/div/div/span/span/table/tbody/tr[2]/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/table/tbody/tr[2]/td[3]/table/tbody/tr/td/span')
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

    wait_and_click(browser, By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/div/div[2]/div/div/div/span[4]/div')  # Click 'Kontroller' button
    time.sleep(4)

    # Check for business error here
    if not browser.find_elements(By.XPATH, "//*[contains(text(), 'Udgiftsbilag er kontrolleret og OK')]"):
        raise BusinessError("Fejl ved kontrol af udgiftsbilag.")

    wait_and_click(browser, By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/div/div[2]/div/div/div/span[1]/div')  # Click 'Opret' button
    time.sleep(4)
    if not browser.find_elements(By.XPATH, "//*[contains(text(), 'er oprettet')]"):
        time.sleep(1)
        raise BusinessError("Fejl ved oprettelse af udgiftsbilag, kontrol OK.")


def switch_to_frame(browser, frame):
    """Switch to the required frames to access the form."""
    WebDriverWait(browser, 30).until(EC.frame_to_be_available_and_switch_to_it((By.ID, frame)))


def enter_text(browser, by, value, text):
    """Helper to enter text into a form element."""
    input_element = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((by, value))
    )
    input_element.send_keys(text)


def wait_and_click(browser, by, value):
    """Wait for an element to be clickable, then click it."""

    WebDriverWait(browser, 50).until(EC.presence_of_element_located((by, value)))
    click_element_with_retries(browser, by, value)
