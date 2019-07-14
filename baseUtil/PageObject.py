#encoding = utf-8
from selenium.webdriver.support.wait import WebDriverWait

# 高亮正在执行的目标元素
def getHighLightElement(driver, element):
    driver.execute_script('arguments[0].setAttribute("style", arguments[1]);', element,
                          'background: yellow; border: 2px solid red;')

# 获取元素对象（单个）
def getElement(driver, locType, loc):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(lambda driver: driver.find_element(locType, loc))
        getHighLightElement(driver, element)
        return element
    except Exception as e:
        raise e


if __name__ == '__main__':
    from selenium import webdriver
    driver = webdriver.Firefox()
    driver.get('https://www.baidu.com')
    getElement(driver, 'id', 'kw').send_keys(2019)
    getElement(driver, 'id', 'su').click()
