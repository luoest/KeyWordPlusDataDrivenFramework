#encoding = utf-8

from BSTestRunner import BSTestRunner
import unittest
from baseUtil.PathAndDir import report_dir
from actions.PageActions import getOpenLocalFile, getClose
from baseUtil.PathAndDir import excel_path

if __name__ == '__main__':

    # 生成HTML报告
    test_dir = './testCases'
    discover = unittest.defaultTestLoader.discover(test_dir, 'test_ctripSearch.py')
    print(discover)
    reportName = report_dir + '/report.html'
    with open(reportName, 'w', encoding='utf-8') as f:
        runner = BSTestRunner(stream=f, verbosity=2, title='关键字+数据混合驱动测试', description='测试携程网站搜索机票和酒店')
        runner.run(discover)
        f.close()

    getClose()
    getOpenLocalFile(excel_path)
