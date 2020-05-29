from selenium import webdriver
from selenium.webdriver.common.keys import Keys
# import pyautogui
import xlwings as xw
import time

wb = xw.Book('C:\\Users\\craig\\Documents\\CS3398-Vulcans-S2020\\example.xlsm')
sheet = wb.sheets['Sheet1']

url = sheet.range('AA2').value
username = sheet.range('AA3').value
password = sheet.range('AA4').value
# Legacy Code
# url = pg.prompt("Enter link to website.", "Website URL", "Link Here.")
# username = pg.prompt("Enter Username or Email.", "Username/EMail", "Username or EMail.")
# password = pg.prompt("Enter Password", "Password", "Password.")

driver = webdriver.Chrome(
    r"C:\Users\craig\AppData\Local\Packages\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\TempState\Downloads\chromedriver_win32\chromedriver.exe")


# r"C:\Users\craig\AppData\Local\Packages\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\TempState\Downloads\chromedriver_win32\chromedriver.exe"
def login(url, username, password):
    driver.get(url)
    driver.find_element_by_link_text('Sign in').click()
    time.sleep(1)
    driver.find_element_by_id('username').send_keys(username)
    driver.find_element_by_id("password").send_keys(password)
    driver.find_element_by_id("password").send_keys(Keys.ENTER)
    # driver.find_element_by_link_text('A I').click()

login(url, username, password)
