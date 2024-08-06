from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

class Web_Contrller:
    def __init__(self, webpage, _sleep_time):
        self._sleep_time = _sleep_time
        self.webpage = webpage
        self.driver = webdriver.Firefox()
        self.driver.get(self.webpage.Domain)
        self.driver.implicitly_wait(self._sleep_time)
        
    def login(self, user_name_xpath, password_xpath, login_button_xpath):
        WebDriverWait(self.driver, self._sleep_time).until(EC.presence_of_element_located((By.NAME, user_name_xpath)))
        self.driver.find_element(By.NAME, user_name_xpath).send_keys(self.webpage.User_Name)
        self.driver.find_element(By.NAME, password_xpath).send_keys(self.webpage.Password)
        self.driver.find_element(By.ID, login_button_xpath).click()

    def wait_element_load_n_click(self, xpath):
        WebDriverWait(self.driver, self._sleep_time).until(EC.presence_of_element_located((By.XPATH, xpath)))
        self.driver.find_element(By.XPATH, xpath).click()

    def fillin_value(self, xpath, value):
        self.driver.find_element(By.XPATH, xpath).clear()
        self.driver.find_element(By.XPATH, xpath).send_keys(value)
