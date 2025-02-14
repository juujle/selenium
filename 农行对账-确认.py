"""农行对账辅助"""
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
        #self.driver.find_element(By.ID,"ok").click() # 确认密码提示
        self.driver.switch_to.frame("indexFrame")
        self.driver.find_element(By.ID,"bill").click() # 查询
        self.driver.switch_to.window(self.handle)
        self.driver.switch_to.frame("contentFrame")
        self.driver.find_element(By.XPATH,'//*[@id="table1"]/tbody/tr[2]/td[2]').click() # 对账单
        self.driver.switch_to.window(self.handle)
        self.driver.switch_to.frame("contentFrame")
        self.driver.find_element(By.CSS_SELECTOR,"input[type=radio]").click() # 通过
        self.driver.find_element(By.XPATH,"/html/body/div[3]/input[1]").click() # 确定


el = Element()
el.checking()
