import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from itertools import count
import openpyxl

url = "https://www.kioskmanager.co.kr/admin/ver2/login.php"

# ID, PW
biz = 4561237890    # 사업자번호
pw = 1234567        # 비밀번호

def set_chrome_driver():
    chrome_options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    return driver

def category_sorting(biz,pw):
    sequence = 'q310,q320,q330,q340,q350,q360'
    driver = set_chrome_driver()
    driver.get(url)
    driver.find_element("id","biz_imput").send_keys(biz)
    driver.find_element("id","pw_input").send_keys(pw)
    driver.find_element("id","login_btn").click()
    time.sleep(2)
    driver.get("https://www.kioskmanager.co.kr/admin/ver2/category_in.php")

    driver.implicitly_wait(5)
    time.sleep(2)
    driver.find_element('xpath','//*[@id="nav_sort_btn"]/img').click()
    driver.execute_script('$(".menu_sort").val("'+sequence+',");')
    driver.find_element('xpath','//*[@id="categoty_sort_confirm"]').click()

category_sorting(biz,pw)