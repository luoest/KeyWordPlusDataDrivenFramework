from baseUtil.Excelling import ExcelParsing
from actions.PageActions import *
import traceback
from baseUtil.myLog import *
from baseUtil.PathAndDir import excel_path

excelObj = ExcelParsing()
excelObj.getWorkBook(excel_path)
caseSheet = excelObj.getSheet('用例选项')

# “用例选项”对应单元格序号
caseSheetTypeNo = 3
caseSheetNameNo = 4
caseSheetDataSheetNo = 5
caseSheetChoicNo = 6
caseSheetTimeNo = 7
caseSheetResultNo = 8

# “登录携程”对应单元格序号
keyMethodNo = 2
keyLocTypeNo = 3
keyLocNo = 4
keyValueNo = 5
keyTimeNo = 6
keyRsultNo = 7
keyErroNo = 8
keyPicNo = 9

# “机票查询数据”对应单元格序号
dataChoiceNo = 6  # 添加新的用例，如酒店查询，是否执行及后面的选项都应放在相应单元格
dataTimeNo = 7
dataResultNo = 8
dataErroNo = 9
dataPicNo = 10
# “机票查询”对应单元格序号
stepMethodNo = 2
stepLocTypeNo = 3
stepLocNo = 4
stepCellNameNo = 5

def writeResult(sheet, result, rowNo, colNo, nowTime=None, timeNo=None, erro=None, erroNo = None, picPath=None, picNo = None, style=None):
    if erro:
        excelObj.getWriteByRowAndCol(sheet, content=erro, rowNo=rowNo, colNo=erroNo, style=style)
        excelObj.getWriteByRowAndCol(sheet, content=picPath, rowNo=rowNo, colNo=picNo, style=style)
    excelObj.getWriteByRowAndCol(sheet, result, rowNo, colNo, style)
    excelObj.getWriteTimeByRowAndCol(sheet, nowTime=nowTime, rowNo=rowNo, colNo=timeNo, style=style)