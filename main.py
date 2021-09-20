import os
import time
import requests
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from utils.my_ultis import MyUtils
from utils.my_locators import MyLocators


# make change dir to run test script
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# declare driver with options
chrome_options = Options()
driver = webdriver.Chrome(
    ChromeDriverManager().install(), options=chrome_options)

# change list urls to check page speed insights
file_name = os.path.join(os.getcwd(), "./data/list_urls.txt")

# declare some variables
now = datetime.now()
str_time = now.strftime('%b_%d_%Y_%H')
excel_path = f"./report/{str_time}_Report_PageSpeed_Insights.xlsx"

# declare class objects from 'utils' folder
my_utils = MyUtils(driver, chrome_options, excel_path, file_name)
my_locators = MyLocators


def check_url_with_page_speed_insight(url, result_row):
    my_utils.my_open_chrome_browser()
    driver.get(my_locators.pagespeed_url)
    input_txt = driver.find_element_by_css_selector(my_locators.url_input)
    input_txt.clear()
    input_txt.send_keys(url)
    time.sleep(2)
    click_analyze_btn = driver.find_element_by_xpath(my_locators.analyze_btn)
    click_analyze_btn.click()
    # wait until check page speed analyzing the page
    time.sleep(5)
    try:
        w_mobile = WebDriverWait(driver, 120).until(
            EC.visibility_of_element_located(
                (By.XPATH, my_locators.mobile_result))
        )
        if w_mobile:
            print('Appeared')
    except:
        print('There was a problem with the request. Please try again later.')
    else:
        # get mobile performance score
        total_score = driver.find_elements(
            By.XPATH, my_locators.scores_performance)
        print('Mobile Score: ', total_score[0].text)
        my_utils.write_excel_result_performance(
            result_row, 2, int(total_score[0].text))
        w_desktop_result = driver.find_element(
            By.XPATH, my_locators.desktop_result)
        w_desktop_result.click()
        # get desktop performance score
        total_score = driver.find_elements(
            By.XPATH, my_locators.scores_performance)
        print('Desktop Score: ', total_score[1].text)
        my_utils.write_excel_result_performance(
            result_row, 3, int(total_score[1].text))


def get_info_page_speed_insight():
    urls = my_utils.read_file_txt(file_name)
    print('[INFO] Total URLs:', len(urls))
    print('\n###################### CHECKING ######################')
    for result_row, url in enumerate(urls, start=2):
        my_utils.write_excel(result_row, url)
        with requests.Session() as session:
            session.headers = my_locators.my_header
            try:
                r = session.get(url)
                print(f'Checking... {r.url} => {r.status_code}')
            except:
                e_err = 'This site can not be reached.'
                my_utils.write_excel_result(result_row, str(e_err))
            else:
                if r.ok:
                    my_utils.write_excel_result(result_row, r.reason)
                    # open chrome to check performance website after access website successfully.
                    check_url_with_page_speed_insight(url, result_row)
                else:
                    sta_err = f'Error Code {r.status_code}'
                    my_utils.write_excel_result(result_row, sta_err)


if __name__ == '__main__':
    try:
        my_utils.create_excel()
        get_info_page_speed_insight()
    finally:
        my_utils.fill_color()
        print('###################### DONE ######################')
        driver.quit()
