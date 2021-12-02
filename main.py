# %%
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import datetime
from time import sleep
import calendar
from selenium.webdriver.support.ui import Select

# %%
# ----------------------- Start Scrapper -----------------------

# %%
# Initialize Chrome
driver = webdriver.Chrome()
driver.maximize_window()

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
# Hover Function List
sleep(1)
function_list = driver.find_element(By.ID, 'F100')
hover = ActionChains(driver).move_to_element(function_list)
hover.perform()

# %%
# Open Results Dropdown
sleep(1)
results_menu = driver.find_element(By.ID, 'F40_dl')
results_menu.click()

# %%
# ----------------------- Start Vehicle Expense Report -----------------------

# %%
# Vehicel Expense Report
sleep(1)
vehicle_expense_report = driver.find_element(By.XPATH, '//*[@id="F45"]')
vehicle_expense_report.click()

# %%
# Cancel button to ensure the menu is turn off
sleep(10)
resetActionKey()

# %%
# Checkbox button
sleep(5)
checkbox = driver.find_element(By.XPATH, '//*[@id="F45_eigyousyoLevel_view"]/li/div/span/div/div').click()

# %%
# Handle the Start Date
sleep(1)
inputStartDate = driver.find_element(By.ID, "F45_StartDate")
inputStartDate.click()

index = 0

sleep(1)
while index <= 2:
    inputStartDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputStartDate.send_keys(Keys.BACKSPACE)
    index += 1

inputStartDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Handle the End Date
sleep(1)
inputEndDate = driver.find_element(By.ID, "F45_EndDate")
inputEndDate.click()

index = 0

sleep(1)
while index <= 2:
    inputEndDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputEndDate.send_keys(Keys.BACKSPACE)
    index += 1

inputEndDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Dowload button
sleep(2)
driver.find_element(By.ID, 'F45_CSV_button').click()

# %%
# Waiting for the handling and dowload process
sleep(60)

# %%
# ----------------------- End Vehicle Expense Report -----------------------

# %%
# ----------------------- Start Driver Expense Report -----------------------

# %%

# Open Menu
sleep(1)
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu')
btn_menu.click()

# %%
# Hover Function List
sleep(1)
function_list = driver.find_element(By.ID, 'F100')
hover = ActionChains(driver).move_to_element(function_list)
hover.perform()

# %%
# Driver Expense Report
sleep(1)
driver_expense_report = driver.find_element(By.XPATH, '//*[@id="F46"]')
driver_expense_report.click()

# %%
# Cancel button to ensure the menu is turn off
sleep(10)
resetActionKey()

# %%
# Checkbox button
sleep(5)
checkbox = driver.find_element(By.XPATH, '//*[@id="F46_eigyousyoLevel_view"]/li/div/span/div').click()

# %%
# Handle the Start Date
sleep(1)
inputStartDate = driver.find_element(By.ID, "F46_from")
inputStartDate.click()

index = 0

sleep(1)
while index <= 2:
    inputStartDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputStartDate.send_keys(Keys.BACKSPACE)
    index += 1

inputStartDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Handle the End Date
sleep(1)
inputEndDate = driver.find_element(By.ID, "F46_to")
inputEndDate.click()

index = 0

sleep(1)
while index <= 2:
    inputEndDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputEndDate.send_keys(Keys.BACKSPACE)
    index += 1

inputEndDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Dowload button
sleep(2)
driver.find_element(By.ID, 'F46_csvout_button').click()

# %%
# Waiting for the handling and dowload process
sleep(60)

# %%
# ----------------------- End Driver Expense Report -----------------------

# %%
# ----------------------- Start Driving Report by Driver -----------------------

# %%
# Open Menu
sleep(1)
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu')
btn_menu.click()

# %%
# Hover Function List
sleep(1)
function_list = driver.find_element(By.ID, 'F100')
hover = ActionChains(driver).move_to_element(function_list)
hover.perform()

# %%
# Driving Report by Driver
sleep(1)
driving_report_by_driver = driver.find_element(By.XPATH, '//*[@id="F42"]')
driving_report_by_driver.click()

# %%
# Cancel button to ensure the menu is turn off
sleep(10)
resetActionKey()

# %%
# Checkbox button
sleep(5)
checkbox = driver.find_element(By.XPATH, '//*[@id="F42_eigyousyoLevel_view"]/li/div/span/div').click()

# %%
# Handle the Start Date
sleep(1)
inputStartDate = driver.find_element(By.ID, "F42_from")
inputStartDate.click()

index = 0

sleep(1)
while index <= 2:
    inputStartDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputStartDate.send_keys(Keys.BACKSPACE)
    index += 1

inputStartDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Handle the End Date
sleep(1)
inputEndDate = driver.find_element(By.ID, "F42_to")
inputEndDate.click()

index = 0

sleep(1)
while index <= 2:
    inputEndDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputEndDate.send_keys(Keys.BACKSPACE)
    index += 1

inputEndDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Dowload button
sleep(2)
driver.find_element(By.ID, 'F42_csvout_button').click()

# %%
# Waiting for the handling and dowload process
sleep(60)

# %%
# ----------------------- End Driving Report by Driver -----------------------

# %%
# ----------------------- Start Driver Report by Vehicle -----------------------

# %%
# Open Menu
sleep(1)
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu')
btn_menu.click()

# %%
# Hover Function List
sleep(1)
function_list = driver.find_element(By.ID, 'F100')
hover = ActionChains(driver).move_to_element(function_list)
hover.perform()

# %%
# Driver Report by Vehicle
sleep(1)
driver_report_by_vehicle = driver.find_element(By.XPATH, '//*[@id="F43"]')
driver_report_by_vehicle.click()

# %%
# Cancel button to ensure the menu is turn off
sleep(10)
resetActionKey()

# %%
# Checkbox button
sleep(5)
checkbox = driver.find_element(By.XPATH, '//*[@id="F43_eigyousyoLevel_view"]/li/div/span/div').click()

# %%
# Handle the Start Date
sleep(1)
inputStartDate = driver.find_element(By.ID, "F43_StartDate")
inputStartDate.click()

index = 0

sleep(1)
while index <= 2:
    inputStartDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputStartDate.send_keys(Keys.BACKSPACE)
    index += 1

inputStartDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Handle the End Date
sleep(1)
inputEndDate = driver.find_element(By.ID, "F43_EndDate")
inputEndDate.click()

index = 0

sleep(1)
while index <= 2:
    inputEndDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputEndDate.send_keys(Keys.BACKSPACE)
    index += 1

inputEndDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Dowload button
sleep(2)
driver.find_element(By.ID, 'F43_CSV_button').click()

# %%
# Waiting for the handling and dowload process
sleep(60)

# %%
# ----------------------- End Driver Report by Vehicle -----------------------

# %%
# ----------------------- Start Yearly Mileage Report -----------------------

# %%
# Open Menu
sleep(1)
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu')
btn_menu.click()

# %%
# Hover Function List
sleep(1)
function_list = driver.find_element(By.ID, 'F100')
hover = ActionChains(driver).move_to_element(function_list)
hover.perform()

# %%
# Yearly Mileage Report
sleep(1)
yearly_mileage_report = driver.find_element(By.XPATH, '//*[@id="F44"]')
yearly_mileage_report.click()

# %%
# Cancel button to ensure the menu is turn off
sleep(10)
resetActionKey()

# %%
# Checkbox button
sleep(5)
checkbox = driver.find_element(By.XPATH, '//*[@id="F44_eigyousyoLevel_view"]/li/div/span/div').click()

# %%
sleep(1)

# Open dropdown option Trip Start Year
trip_start_year = driver.find_element(By.ID, 'ComboBox_StartYear')
trip_start_year.click()

# Set the year to current year
sleep(1)
select_trip_start_year = Select(trip_start_year)
select_trip_start_year.select_by_visible_text(str(today.year))

# Open dropdown option Trip Start Month
trip_start_month = driver.find_element(By.ID, 'ComboBox_StartMonth')
trip_start_month.click()

# Set the year to current year
sleep(1)
select_trip_start_month = Select(trip_start_month)
select_trip_start_month.select_by_visible_text(str(today.month))

# Open dropdown option Trip End Year
trip_end_year = driver.find_element(By.ID, 'ComboBox_EndYear')
trip_end_year.click()

# Set the year to current year
sleep(1)
select_trip_end_year = Select(trip_end_year)
select_trip_end_year.select_by_visible_text(str(today.year))

# Open dropdown option Trip End Month
trip_end_month = driver.find_element(By.ID, 'ComboBox_EndMonth')
trip_end_month.click()

# Set the year to current year
sleep(1)
select_trip_end_month = Select(trip_end_month)
select_trip_end_month.select_by_visible_text(str(today.month))

# %%
# Dowload button
sleep(2)
driver.find_element(By.ID, 'F33_CSV_button').click()

# %%
# Waiting for the handling and dowload process
sleep(60)

# %%
# ----------------------- End Yearly Mileage Report -----------------------

# %%
# ----------------------- Start Digital Tachograph Data -----------------------

# %%
# Open Menu
sleep(1)
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu')
btn_menu.click()

# %%
# Hover Function List
sleep(1)
function_list = driver.find_element(By.ID, 'F100')
hover = ActionChains(driver).move_to_element(function_list)
hover.perform()

# %%
# Digital Tachograph Data
sleep(1)
digital_tachograph_data = driver.find_element(By.XPATH, '//*[@id="F49"]')
digital_tachograph_data.click()

# %%
# Cancel button to ensure the menu is turn off
sleep(10)
resetActionKey()

# %%
# Checkbox button
sleep(5)
checkbox = driver.find_element(By.XPATH, '//*[@id="F49_eigyousyoLevel_view"]/li/div/span/div').click()

# %%
# Handle the Start Date
sleep(1)
inputStartDate = driver.find_element(By.ID, "F49_from")
inputStartDate.click()

index = 0

sleep(1)
while index <= 2:
    inputStartDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputStartDate.send_keys(Keys.BACKSPACE)
    index += 1

inputStartDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Dowload button
sleep(2)
driver.find_element(By.ID, 'F49_get_button').click()

# %%
# Waiting for the handling and dowload process
sleep(60)

# %%
# ----------------------- End Digital Tachograph Data -----------------------

# %%
# ----------------------- Start Safety Driving Ranking Report -----------------------

# %%
# Open Menu
sleep(1)
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu')
btn_menu.click()

# %%
# Hover Function List
sleep(1)
function_list = driver.find_element(By.ID, 'F100')
hover = ActionChains(driver).move_to_element(function_list)
hover.perform()

# %%
# Safery Driving Ranking Report
sleep(1)
safety_driving_ranking_report = driver.find_element(By.XPATH, '//*[@id="F41"]')
safety_driving_ranking_report.click()

# %%
# Cancel button to ensure the menu is turn off
sleep(10)
resetActionKey()

# %%
# Checkbox button
sleep(5)
checkbox = driver.find_element(By.XPATH, '//*[@id="F41_eigyousyoLevel_view"]/li/div/span/div').click()

# %%
# Report check button
sleep(5)
report_check = driver.find_element(By.XPATH, '//*[@id="F41_main_form"]/p[2]/div').click()

# %%
# Handle the Start Date
sleep(1)
inputStartDate = driver.find_element(By.ID, "F41_StartDate")
inputStartDate.click()

index = 0

sleep(1)
while index <= 2:
    inputStartDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputStartDate.send_keys(Keys.BACKSPACE)
    index += 1

inputStartDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Handle the End Date
sleep(1)
inputEndDate = driver.find_element(By.ID, "F41_EndDate")
inputEndDate.click()

index = 0

sleep(1)
while index <= 2:
    inputEndDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputEndDate.send_keys(Keys.BACKSPACE)
    index += 1

inputEndDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Dowload button
sleep(2)
driver.find_element(By.ID, 'F41_CSV_button').click()

# %%
# Waiting for the handling and dowload process
sleep(60)

# %%
# ----------------------- End Safety Driving Ranking Report -----------------------

# %%
# ----------------------- Start Driver Report -----------------------

# %%
# Open Menu
sleep(1)
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu')
btn_menu.click()

# %%
# Hover Function List
sleep(1)
function_list = driver.find_element(By.ID, 'F100')
hover = ActionChains(driver).move_to_element(function_list)
hover.perform()

# %%
# Driver Report
sleep(1)
driver_report = driver.find_element(By.XPATH, '//*[@id="F4B"]')
driver_report.click()

# %%
# Cancel button to ensure the menu is turn off
sleep(10)
resetActionKey()

# %%
# Checkbox button
sleep(5)
checkbox = driver.find_element(By.XPATH, '//*[@id="F41_eigyousyoLevel_view"]/li/div/span/div').click()

# %%
# Handle the Start Date
sleep(1)
inputStartDate = driver.find_element(By.ID, "F41_StartDate")
inputStartDate.click()

index = 0

sleep(1)
while index <= 2:
    inputStartDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputStartDate.send_keys(Keys.BACKSPACE)
    index += 1

inputStartDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Handle the End Date
sleep(1)
inputEndDate = driver.find_element(By.ID, "F41_EndDate")
inputEndDate.click()

index = 0

sleep(1)
while index <= 2:
    inputEndDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputEndDate.send_keys(Keys.BACKSPACE)
    index += 1

inputEndDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Dowload button
sleep(2)
driver.find_element(By.ID, 'F41_CSV_button').click()

# %%
# Waiting for the handling and dowload process
sleep(60)

# %%
# ----------------------- End Driver Report -----------------------

# %%
# ----------------------- Start Driving Report by Vehicle -----------------------

# %%
# Open Menu
sleep(1)
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu')
btn_menu.click()

# %%
# Hover Function List
sleep(1)
function_list = driver.find_element(By.ID, 'F100')
hover = ActionChains(driver).move_to_element(function_list)
hover.perform()

# %%
# Driving Report By Vehicle
sleep(1)
driving_report_by_vehicle = driver.find_element(By.XPATH, '//*[@id="F4C"]')
driving_report_by_vehicle.click()

# %%
# Cancel button to ensure the menu is turn off
sleep(10)
resetActionKey()

# %%
# Checkbox button
sleep(5)
checkbox = driver.find_element(By.XPATH, '//*[@id="F41_eigyousyoLevel_view"]/li/div/span/div').click()

# %%
# Handle the Start Date
sleep(1)
inputStartDate = driver.find_element(By.ID, "F41_StartDate")
inputStartDate.click()

index = 0

sleep(1)
while index <= 2:
    inputStartDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputStartDate.send_keys(Keys.BACKSPACE)
    index += 1

inputStartDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Handle the End Date
sleep(1)
inputEndDate = driver.find_element(By.ID, "F41_EndDate")
inputEndDate.click()

index = 0

sleep(1)
while index <= 2:
    inputEndDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputEndDate.send_keys(Keys.BACKSPACE)
    index += 1

inputEndDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Dowload button
sleep(2)
driver.find_element(By.ID, 'F41_CSV_button').click()

# %%
# Waiting for the handling and dowload process
sleep(60)

# %%
# ----------------------- End Driving Report by Vehicle -----------------------

# %%
# ----------------------- Start 拠点別集計表 -----------------------

# %%
# Open Menu
sleep(1)
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu')
btn_menu.click()

# %%
# Hover Function List
sleep(1)
function_list = driver.find_element(By.ID, 'F100')
hover = ActionChains(driver).move_to_element(function_list)
hover.perform()

# %%
# 拠点別集計表
sleep(1)
拠点別集計表 = driver.find_element(By.XPATH, '//*[@id="P21"]')
拠点別集計表.click()

# %%
# Cancel button to ensure the menu is turn off
sleep(10)
resetActionKey()

# %%
# Checkbox button
sleep(5)
checkbox = driver.find_element(By.XPATH, '//*[@id="P21_eigyousyoLevel_view"]/li/div/span/div').click()

# %%
# Handle the Start Date
sleep(1)
inputStartDate = driver.find_element(By.ID, "P21_StartDate")
inputStartDate.click()

index = 0

sleep(1)
while index <= 2:
    inputStartDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputStartDate.send_keys(Keys.BACKSPACE)
    index += 1

inputStartDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Handle the End Date
sleep(1)
inputEndDate = driver.find_element(By.ID, "P21_EndDate")
inputEndDate.click()

index = 0

sleep(1)
while index <= 2:
    inputEndDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputEndDate.send_keys(Keys.BACKSPACE)
    index += 1

inputEndDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Dowload button
sleep(2)
driver.find_element(By.ID, 'P21_CSV_button').click()

# %%
# Waiting for the handling and dowload process
sleep(60)

# %%
# ----------------------- End 拠点別集計表 -----------------------

# %%
# ----------------------- Start 車両別安全評価実績表 -----------------------

# %%
# Open Menu
sleep(1)
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu')
btn_menu.click()

# %%
# Hover Function List
sleep(1)
function_list = driver.find_element(By.ID, 'F100')
hover = ActionChains(driver).move_to_element(function_list)
hover.perform()

# %%
# 車両別安全評価実績表
sleep(1)
車両別安全評価実績表 = driver.find_element(By.XPATH, '//*[@id="F4G"]')
車両別安全評価実績表.click()

# %%
# Cancel button to ensure the menu is turn off
sleep(10)
resetActionKey()

# %%
# Checkbox button
sleep(5)
checkbox = driver.find_element(By.XPATH, '//*[@id="F4G_eigyousyoLevel_view"]/li/div/span/div').click()

# %%
# Remove 営業所毎にページを分けて印刷する checkbox
sleep(5)
last_checkbox = driver.find_element(By.XPATH, '//*[@id="F4G_main_form"]/p[3]/div').click()

# %%
# Check 車両別安全評価実績表 （積算）
sleep(5)
second_checkbox = driver.find_element(By.XPATH, '//*[@id="F4G_main_form"]/p[2]/div').click()

# %%
# Handle the Start Date
sleep(1)
inputStartDate = driver.find_element(By.ID, "F4G_StartDate")
inputStartDate.click()

index = 0

sleep(1)
while index <= 2:
    inputStartDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputStartDate.send_keys(Keys.BACKSPACE)
    index += 1

inputStartDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Handle the End Date
sleep(1)
inputEndDate = driver.find_element(By.ID, "F4G_EndDate")
inputEndDate.click()

index = 0

sleep(1)
while index <= 2:
    inputEndDate.send_keys(Keys.ARROW_RIGHT)
    index += 1

sleep(1)
while index <= 12:
    inputEndDate.send_keys(Keys.BACKSPACE)
    index += 1

inputEndDate.send_keys(scrapDataDate)

sleep(2)
resetActionKey()

# %%
# Dowload button
sleep(2)
driver.find_element(By.ID, 'F4G_CSV_button').click()

# %%
# Waiting for the handling and dowload process
sleep(60)

# %%
# ----------------------- End 車両別安全評価実績表 -----------------------

# %%
# ----------------------- End Scrapper -----------------------


