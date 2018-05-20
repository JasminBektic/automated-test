from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from faker import Faker
from shutil import copy2
import pyautogui
import os


# Copy test file on Desktop
DESKTOP = os.environ['HOME'] + '/Desktop/'
TEST_FILE = 'SqlfgH1Pr_upload.xlsx'
copy2(TEST_FILE, DESKTOP)

browser = webdriver.Chrome()
browser.get('http://www.agdia.com/testing-services/sample-submission-form-online.cfm')

# Elements
username = browser.find_element_by_name('txtName')
company = browser.find_element_by_name('scName')
address = browser.find_element_by_name('txtAddress')
city = browser.find_element_by_name('txtCity')
state = browser.find_element_by_name('txtState')
zip_code = browser.find_element_by_name('txtZip')
phone = browser.find_element_by_name('txtPhone')
fax = browser.find_element_by_name('txtFax')
email = browser.find_element_by_name('email')
send_invoice = browser.find_element_by_name('sendInvoiceSame')
payment_check = browser.find_element_by_name('check1')
upload_xlsx = browser.find_element_by_name('file')
sample_text = browser.find_element_by_name('txtComment')
report_checkbox = browser.find_element_by_name('sendHardReport')
attn = browser.find_element_by_name('txtName3')
attn_company = browser.find_element_by_name('scName3')
attn_address = browser.find_element_by_name('txtAddress3')
attn_city = browser.find_element_by_name('txtCity3')
attn_state = browser.find_element_by_name('txtState3')
attn_zip = browser.find_element_by_name('txtZip3')
attn_phone = browser.find_element_by_name('txtPhone3')
attn_fax = browser.find_element_by_name('txtFax3')
attn_email = browser.find_element_by_name('email3')
submit = browser.find_element_by_name('cmdSubmit')

faker = Faker()

# Actions
username.send_keys(faker.name())
company.send_keys(faker.company())
address.send_keys(faker.street_address())
city.send_keys(faker.city())
state.send_keys(faker.state())
zip_code.send_keys(faker.zipcode())
phone.send_keys(faker.phone_number())
fax.send_keys(faker.phone_number())
email.send_keys(faker.email())
send_invoice.click()
payment_check.click()
upload_xlsx.click()
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.typewrite(TEST_FILE, 0.05)
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('space')
pyautogui.press('enter')
sample_text.send_keys(faker.text(max_nb_chars=200, ext_word_list=None))
report_checkbox.click()
attn.send_keys(faker.name())
attn_company.send_keys(faker.company())
attn_address.send_keys(faker.street_address())
attn_city.send_keys(faker.city())
attn_state.send_keys(faker.state())
attn_zip.send_keys(faker.zipcode())
attn_phone.send_keys(faker.phone_number())
attn_fax.send_keys(faker.phone_number())
attn_email.send_keys(faker.email())

submit.click()

# Remove test file
os.remove(DESKTOP + TEST_FILE)