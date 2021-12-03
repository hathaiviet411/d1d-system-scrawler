# %%
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
from selenium.common.exceptions import NoSuchElementException
import datetime

# %%
# ----------------------- Start Scrapper -----------------------

# %%
# Initialize Chrome
driver = webdriver.Chrome()
driver.maximize_window()
driver.implicitly_wait(10)

# %%
# Navigate to D1D system
url = 'http://itpv2.transtron.fujitsu.com/'
driver.get(url)

# %%
# Fill login information
sleep(1)
input_field = driver.find_element(By.ID, 'userid')
input_field.send_keys('izmb-free')

password_field = driver.find_element(By.ID, 'password')
password_field.send_keys('izumi001')

# %%
# Login
sleep(1)
btn_login = driver.find_element(By.ID, 'login')
btn_login.click()

# %%
# Waiting the client loading
sleep(60)

# %%
def joinYMD(string):
    year = string.split('-')[0]
    month = string.split('-')[1]
    day = string.split('-')[2]
    return str(year) + '/' + str(month) + '/' + str(day)

# %%
# Get yesterday
today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)

# %%
# Transform the format of the scrap date
scrapDataDate = joinYMD(str(yesterday))

# %%
# Reset action
def resetActionKey():
    resetSpace =  driver.find_element(By.XPATH, '//*[@id="F04_mainview"]/div[2]/div[1]')
    resetSpace.click()

# %%
# Function to get the correct xpath of each calendar
def getCurrentCalendarXPATH(calendar_id, year, month):
    return "//table[@id='" + calendar_id + "_calendar_start_" + year + '_' + month + "']//a"

# %%
# Open Menu
sleep(1)
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu')
btn_menu.click()

# %%
# Hover Daily Report
sleep(1)
daily_report = driver.find_element(By.ID, 'F20')
hover = ActionChains(driver).move_to_element(daily_report)
hover.perform()

# %%
# Navigate to Alcohol Measurement Result
sleep(1)
alcohol_measurement_result = driver.find_element(By.XPATH, '//*[@id="F25"]')
alcohol_measurement_result.click()

# %%
# Cancel button to ensure the menu is turn off
sleep(5)
resetActionKey()

# %%
# Select department
sleep(1)
select_department = driver.find_element(By.ID, 'F25eigyousyo_list')
select_department.click()
sleep(1)
driver.find_element(By.XPATH, '//*[@id="F25eigyousyo_list"]/option[2]').click()

# Handle the scrapper date - should be yester day of today
# Start date
start_date = driver.find_element(By.ID, 'F25_StartDate')
start_date.click()

index = 0

sleep(1)
while index <= 2:
    start_date.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    start_date.send_keys(Keys.BACKSPACE)
    index += 1

start_date.send_keys(scrapDataDate)

sleep(1)
resetActionKey()

# End date
end_date = driver.find_element(By.ID, 'F25_EndDate')
end_date.click()

index = 0

while index <= 2:
    end_date.send_keys(Keys.ARROW_RIGHT)
    index += 1

while index <= 12:
    end_date.send_keys(Keys.BACKSPACE)
    index += 1

end_date.send_keys(scrapDataDate)

sleep(1)
resetActionKey()

sleep(1)
resetActionKey()

# %%
# Options in the select department
sleep(1)
options = [item for item in select_department.find_elements(By.TAG_NAME, 'option')]
del options[0]

count = 0
existMessageBox = True

for element in (options):
    print('Processing department: ' + str(element.text))
    element.click()

    # Check if the selected department have data or not
    try:
        noDataNotifyModal = driver.find_element(By.ID, 'function_window_F25_messagebox_0')
    except NoSuchElementException:
        existMessageBox = False

    # If the selected departsment have data then download the data
    # otherwise, click on the OK button to turn off the modal 
    if existMessageBox == True:
        driver.find_element(By.ID, 'msg_confirm').click()
    elif existMessageBox == False:
        driver.find_element(By.ID, 'P25_csvout_button').click()

    if existMessageBox == True:
        print('Department: ' + element.text + ' have no data, passed !')
    else:
        print('Department: ' + element.text + ' proceed successfully !')
    print('--------------------------------------------------------------')
    existMessageBox = True
    sleep(2)
    count += 1

print('----------------------------- END SCRAPPER -----------------------------')


