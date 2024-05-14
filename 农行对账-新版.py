'''农行对账辅助'''
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service

class Element:
    '''封装成类'''
    def __init__(self):
        options = webdriver.EdgeOptions()
        options.add_experimental_option("debuggerAddress", "127.0.0.1:5001")
        service = Service(r'f:\tools\msedgedriver.exe')
        self.driver = webdriver.Edge(options = options,service=service)
        self.driver.implicitly_wait(5)  # 隐式等待,后续所有的 find_element
        for handle in self.driver.window_handles:
            self.driver.switch_to.window(handle)
            if self.driver.title=="中国农业银行企业网银":
                break
        self.handle = self.driver.current_window_handle


    def checking(self):
        '''对账'''
        self.driver.switch_to.frame("indexFrame")
        self.driver.find_element(By.XPATH,
            '//*[@id="height-ctrl"]/div[2]/div[1]/ul/span/div[3]/div[2]/div[2]/div[2]').click() # 对账单
        self.driver.switch_to.window(self.handle)
        self.driver.switch_to.frame("contentFrame")
        self.driver.find_element(By.XPATH,
            '//*[@id="bodyb"]/div/div/div/div[2]/div[1]/div[3]/table/tbody/tr/td[2]/div/div/button').click() # 账单编号
        self.driver.switch_to.window(self.handle)
        self.driver.switch_to.frame("contentFrame")
        self.driver.find_element(By.XPATH,
            '//*[@id="bodyb"]/div/div/div/div[3]/div[1]/div[2]/table/thead/tr/th[7]/div/label/span[1]/span').click() # 全部相符
        self.driver.find_element(By.XPATH,
            '//*[@id="bodyb"]/div/div/div/div[4]/div[1]/button').click() # 提交


el=Element()
el.checking()
