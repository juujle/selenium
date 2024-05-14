"""财务公布报表自动化处理及打印"""
import time
import datetime
import pyautogui as pag
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service


def get_month() -> int:
    """返回会计月份(当前月份的上一个月)"""
    month = datetime.date.today().month
    month = ((month + 10) % 12) + 1  # 当前月份的前一个月
    return month


def get_list():
    """返回要打印及修改报表列表"""
    month = get_month()
    if month == 3:
        rep_list = (0, 14, 15, 9, 10, 11)
        modi_list = (9, 10, 11, 14)
    elif month in (6, 9):
        rep_list = (12, 13, 8, 9, 10)
        modi_list = (8, 9, 10, 12)
    elif month == 12:
        rep_list = (0, 12, 14, 15, 9, 10, 11)
        modi_list = (9, 10, 11, 14)
    else:
        rep_list = (11, 8, 9)
        modi_list = (8, 9)
    return rep_list, modi_list


class Form:
    """封装成类"""

    def __init__(self):
        options = webdriver.EdgeOptions()
        options.add_experimental_option("debuggerAddress", "127.0.0.1:5001")
        self.driver = webdriver.Edge(options=options,service=Service(r'f:\tools\msedgedriver.exe'))
        self.driver.implicitly_wait(5)  # 隐式等待,后续所有的 find_element
        for handle in self.driver.window_handles:
            self.driver.switch_to.window(handle)
            if self.driver.title == "黄埔区农村集体三资平台":
                break
        self.handle = self.driver.current_window_handle
        self.driver.switch_to.frame("frame_content")

    def switch(self, index: int):
        """按索引号转到某公布表"""
        select = Select(self.driver.find_element(By.ID, "replst"))
        select.select_by_index(index)
        time.sleep(1)

    def modify(self):
        """数据编辑增加空白行"""
        self.driver.find_element(By.XPATH, '//*[@id="topbtn"]/button[3]').click()  # 数据编辑
        time.sleep(1)
        # 切换到动态iframe
        frame = self.driver.find_elements(By.TAG_NAME, 'iframe')
        self.driver.switch_to.frame(frame[-1].get_attribute('id'))
        # 点击第二栏及输入" "
        self.driver.find_element(By.XPATH,
            '/html/body/div[2]/div/div/div/div/div[2]/div[2]/table/tbody/tr/td[2]').click()
        self.driver.find_elements(By.XPATH, '//tbody/tr/td/input')[0].send_keys(' ')  # 输入
        self.driver.find_element(By.XPATH, '//*[@id="btnEdtsave"]/tbody/tr/td[2]/em/button').click()  # 保存
        self.driver.find_element(By.XPATH, '/html/body/div[6]/div[2]/div[4]/a/span/span').click()  # 确认
        self.driver.find_element(By.XPATH, '//*[@id="btnEdtClose"]/tbody/tr/td[2]/em/button').click()  # 返回
        time.sleep(1)
        self.driver.switch_to.parent_frame()  # 切回主iframe

    def print(self, xpath: str, second: int):
        """打印报表: xpath为批打印或否; second为等待秒数"""
        self.driver.find_element(By.ID, "btnPrt").click()  # 打印
        self.driver.find_element(By.XPATH, xpath).click()  # 数据打印或批打印
        time.sleep(second)
        self.driver.switch_to.window(self.driver.window_handles[-1])
        pag.keyDown("ctrl")  # 快捷键Ctrl+P调用打印
        pag.press("p")
        pag.keyUp("ctrl")
        time.sleep(2)
        pag.hotkey("Enter")  # 确认打印
        time.sleep(1)
        self.driver.close()
        # 恢复到默认窗口及框架
        self.driver.switch_to.window(self.handle)
        self.driver.switch_to.frame("frame_content")


form = Form()
rep_lists, modi_lists = get_list()

form.switch(0)
form.print('//*[@id="linkprt"]/div[2]', 7)  # 批打印

for rep in rep_lists:  # 逐个表打印
    form.switch(index=rep)
    if rep in modi_lists:  # 是否需要修改表
        form.modify()
    form.print('//*[@id="linkprt"]/div[1]', 5)
