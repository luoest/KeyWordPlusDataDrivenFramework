from testCases.WriteTestResults import *

def DataDrivenForSearching(dataSheetObj, stepSheetObj, caseRowNo):
    executing_method = ''
    error = ''
    errorPicPath = ''
    target_cases = 0
    successful_cases = 0

    dataChoices = excelObj.getColValue(dataSheetObj, dataChoiceNo)
    for idx, dataChoice in enumerate(dataChoices[1:]):
        if dataChoice.value.upper() == 'Y':
            target_cases += 1
            stepRowNum = excelObj.getRowAndColNum(stepSheetObj)[0]
            stepNum = 0

            for step in range(2, stepRowNum+1):
                methodName = excelObj.getRowValue(stepSheetObj, step)[stepMethodNo-1].value
                locType = excelObj.getRowValue(stepSheetObj, step)[stepLocTypeNo-1].value
                loc = excelObj.getRowValue(stepSheetObj, step)[stepLocNo - 1].value
                cellParam = excelObj.getRowValue(stepSheetObj, step)[stepCellNameNo - 1].value

                if cellParam:
                    try:
                        cellParam = int(cellParam)
                        executing_method = methodName + '(%d)'%(cellParam)
                    except:
                        cellParam += str(idx+2)
                        cellValue = excelObj.getCellValue(dataSheetObj, cellParam)
                        if methodName and locType and loc:
                            executing_method = methodName + '("%s","%s","%s")'%(locType, loc, cellValue)
                        elif methodName and locType is None and loc is None:
                            executing_method = methodName + '("%s")'%(cellValue)
                elif methodName and locType is None and loc is None and cellParam is None: # 没有任何参数的情况
                    executing_method = methodName + '()'
                else:
                    executing_method = methodName + '("%s","%s")' % (locType, loc)
                try:
                    eval(executing_method)
                    msg = '通过：%s' % str(executing_method)
                    print(msg)
                    logging.info(msg)
                    stepNum+=1
                except Exception as e:
                    msg = '未通过：%s' % str(executing_method)
                    error = traceback.format_exc()
                    errorPicPath = getImage()
                    print(msg)
                    logging.debug(error)
                    print(e, error)
            if stepNum == stepRowNum - 1:
                msg = '<成功>用例：“%s”共有步骤：“%d”，执行成功步骤：“%d”' % (dataSheetObj.title, stepRowNum - 1, stepNum)
                print(msg)
                logging.info(msg)
                writeResult(dataSheetObj, msg, idx+2, colNo=dataResultNo, timeNo=dataTimeNo, style='blue',
                            nowTime=excelObj.now, erro=' ', erroNo=dataErroNo, picPath=' ',picNo=dataPicNo)
                successful_cases += 1
            else:
                msg = '<失败>用例：“%s”共有步骤：“%d”，执行成功步骤：“%d”' % (dataSheetObj.title, stepRowNum - 1, stepNum)
                print(msg)
                logging.info(msg)
                writeResult(dataSheetObj, msg, idx+2, dataResultNo, timeNo=dataTimeNo, style='red',
                            nowTime=excelObj.now, erro=error, erroNo=dataErroNo, picPath=errorPicPath,picNo=dataPicNo)

    if target_cases == successful_cases:
        msg = '<成功>用例：“%s”共有用例：“%d”，执行成功：“%d”' % (dataSheetObj.title, target_cases, successful_cases)
        print(msg)
        logging.info(msg)
        writeResult(caseSheet, msg, caseRowNo, caseSheetResultNo, timeNo=dataTimeNo, style='blue',
                    nowTime=excelObj.now)
    else:
        msg = '<失败>用例：“%s”共有用例：“%d”，执行成功：“%d”' % (dataSheetObj.title, target_cases, successful_cases)
        print(msg)
        logging.info(msg)
        writeResult(caseSheet, msg, caseRowNo, caseSheetResultNo, timeNo=dataTimeNo, style='red',
                    nowTime=excelObj.now)

if __name__ == '__main__':
    from testCases.TestCtripSearch import TestCtripSearh
    TestCtripSearh()
    dataSheetObj = excelObj.getSheet('机票查询数据')
    stepSheetObj = excelObj.getSheet('机票查询')
    DataDrivenForSearching(dataSheetObj, stepSheetObj, 3)













