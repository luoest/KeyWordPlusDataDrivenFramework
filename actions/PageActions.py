#encoding = utf-8

from selenium import webdriver
from baseUtil.WaitUtiling import WaitUtiling
from baseUtil.PageObject import getElement
import time, os
from baseUtil.PathAndDir import exception_dir
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

def getBrowser(browserName):
    global driver, wait, browser_name
    if browserName.lower() == 'edge':
        driver = webdriver.Edge()
    elif browserName.lower() == 'chrome':
        driver = webdriver.Chrome()
    else:
        driver = webdriver.Firefox()
    wait = WaitUtiling(driver)
    browser_name = browserName.lower()

def getMaximizedBrowser():
    driver.maximize_window()

def getUrl(url):
    driver.get(url)

def getClear(locType, loc):
    try:
        getElement(driver, locType, loc).clear()
    except Exception as e:
        print(e)

def getSendKeys(locType, loc, value):
    getElement(driver, locType, loc).send_keys(value)

def getClick(locType, loc):
    getElement(driver, locType, loc).click()

def getSleep(sleepSeconds):
    time.sleep(sleepSeconds)

def getClose():
    driver.quit()

def getPageSource():
    return driver.page_source

def getAssertInPageSource(element):
    try:
        assert str(element) in getPageSource(), '页面中没有发现目标元素：' + str(element)
    except Exception as e:
        raise e

def getTitle():
    return driver.title

def getAssertElementInTitle(element):
    try:
        assert str(element) in getTitle(), '页面标题中没有发现元素：' + str(element)
    except Exception as e:
        raise e

def getOpenLocalFile(filePathAndName):
    os.startfile(filePathAndName)

def getImage():
    now = time.strftime('%Y-%m-%d %H-%M-%S')
    imageName = exception_dir + '/' + now + '.png'
    driver.get_screenshot_as_file(imageName)
    return imageName

def getPressEnterKey():
    '''
    模拟回车键
    '''
    ActionChains(driver).key_down(Keys.CONTROL).send_keys(Keys.ENTER).key_up(Keys.CONTROL).perform()



if __name__ == '__main__':
    try:
        getBrowser('chrome')
        getMaximizedBrowser()
        getUrl('https://www.baidu.com')
        getSendKeys('id', 'kw', 2019)
        getSleep(1)
        getPressEnterKey()

    finally:
        getSleep(3)
        getClose()
