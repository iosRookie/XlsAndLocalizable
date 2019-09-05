# coding=utf-8
import os
from optparse import OptionParser

import xlsxwriter

from Log import Log
from XlsOperationUtil import XlsOperationUtil


def addParser():
    parser = OptionParser()

    parser.add_option("-r", "--resource",
                      help="android or iOS resource path.",
                      metavar="resource")

    parser.add_option("-s", "--save",
                      help="The directory where the xlsx files will be saved.",
                      metavar="save")

    parser.add_option("-t", "--type",
                      type="string",
                      metavar="type")

    (options, args) = parser.parse_args()
    return options


def stringsToXlsx(resourceDir, savePath):
    if resourceDir.endswith("/"):
        resourceDir = resourceDir[:-1]
    workbook = xlsxwriter.Workbook(savePath + "/" + "ios.xlsx")
    titleBold = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'left',
        'valign': 'vcenter',
        'fg_color': '#00FF00',
    })

    for _, dirName, _ in os.walk(resourceDir):
        lprojDirs = sorted([di for di in dirName if di.endswith(".lproj")])
        for index, lproj in enumerate(lprojDirs):
            for _, _, files in os.walk(resourceDir + "/" + lproj):
                stringFiles = sorted([file for file in files if file.endswith(".strings")])
                for strFile in stringFiles:
                    (keys, values) = XlsOperationUtil.getIOSKeysAndValues(resourceDir + "/" + lproj + "/" + strFile)
                    if index == 0:
                        worksheet = workbook.add_worksheet(strFile)
                        worksheet.write(0, 0, "Keys", titleBold)
                        worksheet.write(0, 1, lproj.replace(".lproj", ""), titleBold)
                        worksheet.write_column("A2", keys)
                        worksheet.write_column("B2", values)
                    else:
                        sheets = workbook.worksheets()
                        for sheet in sheets:
                            if sheet.name == strFile:
                                column = str(chr(ord("A") + 1 + index)) + "2"
                                sheet.write(0, 1 + index, lproj.replace(".lproj", ""), titleBold)
                                sheet.write_column(column, values)
    workbook.close()


def xmlToXlsx(resourceDir, savePath):
    if resourceDir.endswith("/"):
        resourceDir = resourceDir[:-1]
    workbook = xlsxwriter.Workbook(savePath + "/" + "android.xlsx")
    worksheet = workbook.add_worksheet("strings.xml")
    titleBold = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'left',
        'valign': 'vcenter',
        'fg_color': '#00FF00',
    })

    for _, dirnames, _ in os.walk(resourceDir):
        valuesDirs = sorted([di for di in dirnames if
                             (di.startswith("values") and os.path.exists(
                                 resourceDir + "/" + di + "/" + "strings.xml"))])
        for index, value in enumerate(valuesDirs):
            (keys, values) = XlsOperationUtil.getAndroidKeysAndValues(resourceDir + '/' + value + '/' + "strings.xml")
            if index == 0:
                worksheet.write(0, 0, "Keys", titleBold)
                worksheet.write(0, 1, "value", titleBold)
                worksheet.write_column("A2", keys)
                worksheet.write_column("B2", values)
            else:
                column = str(chr(ord("A") + 1 + index)) + "2"
                worksheet.write(0, 1 + index, value, titleBold)
                worksheet.write_column(column, values)
    workbook.close()


def main():
    options = addParser()
    if options.type is None:
        Log.info("Change type can not be empty!")
        return
    if options.resource is None:
        Log.info("xls file can not be empty!")
        return
    if options.save is None:
        Log.info("save path can not be empty!")
        return

    if not os.path.exists(options.save):
        os.makedirs(options.save)

    if options.type == "ios":
        stringsToXlsx(options.resource, options.save)
        Log.info("Completed iOS to xls ! " + options.save)
    #     stringsToXlsx("/Users/yyg/Documents/ukelink/fork/yyg/simbox-app/platforms/iOS/SIMBOX+/Resources",
    #               "/Users/yyg/Desktop/temp")
    if options.type == "android":
        xmlToXlsx(options.resource, options.save)
        Log.info("Completed android to xls ! " + options.save)
    # xmlToXlsx("/Users/yyg/Documents/ukelink/fork/yyg/android/simbox-app/platforms/android/app/src/main/res", "/Users/yyg/Desktop/temp")


main()
