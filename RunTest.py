from testCases.TestCtripSearch import TestCtripSearh

if __name__ == '__main__':
    from actions.PageActions import getOpenLocalFile, getClose
    from baseUtil.PathAndDir import excel_path
    TestCtripSearh()
    getClose()
    getOpenLocalFile(excel_path)
