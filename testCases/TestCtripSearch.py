from testCases.WriteTestResults import *
from testCases.SearchData import DataDrivenForSearching

def TestCtripSearh():
    executing_method = ''
    caseChoices = excelObj.getColValue(caseSheet, caseSheetChoicNo)
    for idx, caseChoice in enumerate(caseChoices[1:]):
        if caseChoice.value.upper() == 'Y':
            caseType = excelObj.getRowValue(caseSheet, idx+2)[caseSheetTypeNo - 1].value
            if caseType == '关键字':
                keySheetName = excelObj.getRowValue(caseSheet, idx+2)[caseSheetNameNo-1].value
                keySheetObj = excelObj.getSheet(keySheetName)

                keyRowNum = excelObj.getRowAndColNum(keySheetObj)[0]
                stepNo = 0
                for step in range(2, keyRowNum+1):
                    methodName = excelObj.getRowValue(keySheetObj, step)[keyMethodNo-1].value
                    locType = excelObj.getRowValue(keySheetObj, step)[keyLocTypeNo-1].value
                    loc = excelObj.getRowValue(keySheetObj, step)[keyLocNo-1].value
                    param = excelObj.getRowValue(keySheetObj, step)[keyValueNo-1].value
                    if methodName and param and locType is None and loc is None:
                        if isinstance(param, int):
                            executing_method = methodName + '(%d)'%param
                        else:
                            executing_method = methodName + '("%s")'%param
                    if methodName and locType and loc and param:
                        executing_method = methodName + '("%s", "%s", "%s")'%(str(locType), str(loc), str(param))
                    elif methodName and locType and loc and param is None:
                        executing_method = methodName + '("%s", "%s")'%(str(locType), str(loc))
                    elif methodName and locType is None and loc is None and param is None:
                        executing_method = methodName + '()'

                    try:
                        eval(executing_method)
                        msg = '通过：%s'%str(executing_method)
                        print(msg)
                        logging.info(msg)
                        writeResult(keySheetObj, msg, rowNo=step, colNo=keyRsultNo, style='blue', timeNo=keyTimeNo,
                                    erro=' ', erroNo=keyErroNo, picPath=' ', picNo=keyPicNo, nowTime=excelObj.now)
                        stepNo += 1
                    except Exception as e:
                        msg = '未通过：%s'%str(methodName)
                        print(msg)
                        errorInfo = traceback.format_exc()
                        logging.debug(msg)
                        errorPicPath = getImage()
                        writeResult(keySheetObj, result=msg, rowNo=step, colNo=keyRsultNo, timeNo=keyTimeNo, style='red',
                                    erro=str(errorInfo), erroNo=keyErroNo, picPath=errorPicPath, picNo=keyPicNo, nowTime=excelObj.now)

                if stepNo == keyRowNum-1:
                    msg = '<成功>用例：“%s”共有步骤：“%d”，执行成功步骤：“%d”'%(keySheetObj.title, keyRowNum-1, stepNo)
                    print(msg)
                    logging.info(msg)
                    writeResult(caseSheet,msg,idx+2,caseSheetResultNo, timeNo=caseSheetTimeNo, style='blue', nowTime=excelObj.now)
                else:
                    msg = '<失败>用例：“%s”共有步骤：“%d”，执行成功步骤：“%d”' % (keySheetObj.title, keyRowNum - 1, stepNo)
                    print(msg)
                    logging.info(msg)
                    writeResult(caseSheet, msg, idx + 2, caseSheetResultNo, timeNo=caseSheetTimeNo, style='red', nowTime=excelObj.now)

            elif caseType == '数据':
                stepSheetName = excelObj.getRowValue(caseSheet, idx+2)[caseSheetNameNo-1].value
                stepSheetObj = excelObj.getSheet(stepSheetName)
                dataSheetName = excelObj.getRowValue(caseSheet, idx + 2)[caseSheetNameNo].value
                dataSheetObj = excelObj.getSheet(dataSheetName)
                DataDrivenForSearching(dataSheetObj, stepSheetObj, idx+2)

            # caseType不需要执行的用例，将之前执行结果和执行时间清空
            else:
                writeResult(caseSheet, result=' ', rowNo=idx + 2, colNo=caseSheetResultNo, nowTime=' ',timeNo=caseSheetTimeNo, style='red')
        # caseChoice.value.upper() 不等于 'Y'的情况下，清空之前的执行结果和执行时间
        else:
            writeResult(caseSheet, result=' ', rowNo=idx + 2, colNo=caseSheetResultNo, nowTime=' ',
                        timeNo=caseSheetTimeNo, style='red')




if __name__ == '__main__':
    from actions.PageActions import getOpenLocalFile
    TestCtripSearh()
    getOpenLocalFile(r'D:\PycharmProjects\KeyWordPlusDataDrivenFramework\testData\携程查询机票酒店.xlsx')

















