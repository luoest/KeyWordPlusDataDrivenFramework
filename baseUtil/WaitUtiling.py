#encoding = utf-8
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

class WaitUtiling():
    def __init__(self, driver):
        self.driver = driver
        self.timeout = 10
        self.wait = WebDriverWait(self.driver, self.timeout)
        self.locDict = {'id': By.ID,
                        'link text': By.LINK_TEXT,
                        'partial link text': By.PARTIAL_LINK_TEXT,
                        'css': By.CSS_SELECTOR}

    def getPresentElement(self, locType, loc):
        try:
            if locType.lower() in self.locDict:
                element = self.wait.until(EC.presence_of_element_located((self.locDict[locType.lower()], loc)))
                return element
            else:
                print('没有发现匹配的定位方式。')
        except Exception as e:
            print(e)

    def getVisibleElement(self, locType, loc):
        try:
            if locType.lower() in self.locDict:
                element = self.wait.until(EC.visibility_of_element_located((self.locDict[locType.lower()], loc)))
                return element
            else:
                print('没有发现匹配的定位方式。')
        except Exception as e:
            print(e)


if __name__ == '__main__':
    from selenium import webdriver
    driver = webdriver.Chrome()
    driver.get('https://www.baidu.com')
    wu = WaitUtiling(driver)
    wu.getPresentElement('id', 'kw').send_keys('python')
    wu.getVisibleElement('id', 'su').click()
    wu.getVisibleElement('partial link text', 'Welcome').click()

