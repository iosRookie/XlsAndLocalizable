# coding=utf-8
import xlrd
import xlsxwriter

from XlsOperationUtil import XlsOperationUtil


def generateiOSandAndroidCommonKey(iosFile, androidFile):
    (ikeys, ivalues) = XlsOperationUtil.getIOSKeysAndValues(iosFile)
    (akeys, avalues) = XlsOperationUtil.getAndroidKeysAndValues(androidFile)
    temp = []
    for value in avalues:
        if value not in temp:
            temp.append(value)
        else:
            print value

    commonIOSKeys = []
    commonAndroidKeys = []
    commonChineseValues = []

    onlyAndroidKeys = []
    onlyAndroidValues = []

    onlyIOSKeys = []
    onlyIOSValues = []

    for index, key in enumerate(ikeys):
        if ivalues[index] in avalues:
            commonIOSKeys.append(key)
            commonChineseValues.append(ivalues[index])
            commonAndroidKeys.append(akeys[avalues.index(ivalues[index])])
        else:
            onlyIOSKeys.append(key)
            onlyIOSValues.append(ivalues[index])
    for index, key in enumerate(akeys):
        if key not in commonAndroidKeys:
            onlyAndroidKeys.append(key)
            onlyAndroidValues.append(avalues[index])
    return commonChineseValues
    # XlsOperationUtil.writeToFile("/Users/yyg/Desktop/temp", "common.xlsx", ["iOS-Key", "android-Key", "commonValues"], [commonIOSKeys, commonAndroidKeys, commonValues])
    # XlsOperationUtil.writeToFile("/Users/yyg/Desktop/temp", "iosSingle.xlsx", ["Key", "Value"], [onlyIOSKeys, onlyIOSValues])
    # XlsOperationUtil.writeToFile("/Users/yyg/Desktop/temp", "androidSingle.xlsx", ["Key", "Value"], [onlyAndroidKeys, onlyAndroidValues])
    # XlsOperationUtil.writeToFile("/output", "difference.xlsx", ["iOS-Key", "iOS-Value", "android-Key", "android-Value"], [onlyIOSKeys, onlyIOSValues, onlyAndroidKeys, onlyAndroidValues])

def commonWithXlsx(iosXls, androidXls, commonChineseValues):
    ioscommon = {}
    ios = []
    androidcommon = {}
    android = []
    iosBook = xlrd.open_workbook(iosXls)
    androidBook = xlrd.open_workbook(androidXls)
    for iosSheet in iosBook.sheets():
        if iosSheet.name == "Localizable.strings":
            for index, row in enumerate(iosSheet.get_rows()):
                if len(row) >= 2 and (row[2].value in commonChineseValues):
                    # 中文简体相等的值
                    ioscommon[row[2].value] = row
                else:
                    ios.append(row)
    for androidSheet in androidBook.sheets():
        if androidSheet.name == "strings.xml":
            for index, row in enumerate(androidSheet.get_rows()):
                if len(row) >= 2 and (row[2].value in commonChineseValues):
                    # 中文简体相等的值
                    androidcommon[row[2].value] = row
                else:
                    android.append(row)
    # ios独有国际化
    wiosBook = xlsxwriter.Workbook("iosingle.xlsx")
    wiosworksheet = wiosBook.add_worksheet()
    wiostitleBold = wiosBook.add_format({
        'bold': True,
        'border': 1,
        'align': 'left',
        'valign': 'vcenter',
        'fg_color': '#00FF00',
    })
    for index, row in enumerate(ios):
        for cindex, cell in enumerate(row):
            if index == 0:
                wiosworksheet.write(index, cindex, cell.value, wiostitleBold)
            else:
                wiosworksheet.write(index, cindex, cell.value)
    wiosBook.close()
    # android独有国际化
    wandroidBook = xlsxwriter.Workbook("androidSingle.xlsx")
    wandroidworksheet = wandroidBook.add_worksheet()
    wandroidtitleBold = wandroidBook.add_format({
        'bold': True,
        'border': 1,
        'align': 'left',
        'valign': 'vcenter',
        'fg_color': '#00FF00',
    })
    for index, row in enumerate(android):
        for cindex, cell in enumerate(row):
            if index == 0:
                wandroidworksheet.write(index, cindex, cell.value, wandroidtitleBold)
            else:
                wandroidworksheet.write(index, cindex, cell.value)
    wandroidBook.close()
    # 公共国际化
    commonBook = xlsxwriter.Workbook("common.xlsx")
    commonsheet = commonBook.add_worksheet()
    commontitleBold = commonBook.add_format({
        'bold': True,
        'border': 1,
        'align': 'left',
        'valign': 'vcenter',
        'fg_color': '#00FF00',
    })
    commonsheet.write(0, 0, "iOSKey", commontitleBold)
    commonsheet.write(0, 1, "androidKey", commontitleBold)
    commonsheet.write(0, 2, "en", commontitleBold)
    commonsheet.write(0, 3, "zh", commontitleBold)
    commonsheet.write(0, 4, "zhs", commontitleBold)
    for index, key in enumerate(commonChineseValues):
        iosrow = ioscommon.get(key)
        androidrow = androidcommon.get(key)
        commonsheet.write(index + 1, 0, iosrow[0].value)
        commonsheet.write(index + 1, 1, androidrow[0].value)
        commonsheet.write(index + 1, 2, iosrow[1].value)
        commonsheet.write(index + 1, 3, iosrow[2].value)
        commonsheet.write(index + 1, 4, iosrow[3].value)
    commonBook.close()


def main():
    # 提取中文相等的
    iosfile = "/Users/yyg/Documents/ukelink/fork/yyg/simbox-app/platforms/iOS/SIMBOX+/Resources/zh-Hans.lproj/Localizable.strings"
    androidfile = "/Users/yyg/Documents/ukelink/fork/yyg/android/simbox-app/platforms/android/app/src/main/res/values-zh-rCN/strings.xml"
    commonChineseValues = generateiOSandAndroidCommonKey(iosfile, androidfile)

    iOSXls = "/Users/yyg/Desktop/temp/ios.xlsx"
    androidXls = "/Users/yyg/Desktop/temp/android.xlsx"
    commonWithXlsx(iOSXls, androidXls, commonChineseValues)
main()
