# coding=utf-8
import os
import re
from optparse import OptionParser
from Log import Log
import xlrd

def addParser():
    parser = OptionParser()

    parser.add_option("-f", "--file",
                      help="Xls file.",
                      metavar="file")

    parser.add_option("-s", "--save",
                      help="The directory where the strings or xml files will be saved.",
                      metavar="save")

    parser.add_option("-t", "--type",
                      type="string",
                      metavar="type")

    (options, args) = parser.parse_args()
    return options


def XlsToString(filePath, saveDir):
    workbook = xlrd.open_workbook(filePath)
    for sheet in workbook.sheets():
        # 创建国际化文件夹
        dirNames = [saveDir + "/" + dirName.value + ".lproj" for dirName in sheet.get_rows().next()]
        # 去除第一个Key的文件夹路径
        if len(dirNames) > 0: dirNames.pop(0)
        for dirName in dirNames:
            if not os.path.exists(dirName):
                os.makedirs(dirName, mode=0777)
        # 按类型将字符串写入
        for index, dirName in enumerate(dirNames):
            iosDestFilePath = dirName + "/" + sheet.name
            iosFileManager = open(iosDestFilePath, "wb")
            for vIndex, row in enumerate(sheet.get_rows()):
                if vIndex == 0 and row[0].value == "Keys":
                    continue
                content = "\"" + row[0].value + "\" " + \
                          "= " + "\"" + row[1 + index].value + "\";\n"
                iosFileManager.write(content.encode('utf-8'))
            iosFileManager.close()

def XlsToXml(filePath, saveDir):
    workbook = xlrd.open_workbook(filePath)
    for sheet in workbook.sheets():
        # 创建国际化文件夹
        dirNames = [saveDir + "/" + dirName.value for dirName in sheet.get_rows().next()]
        # 去除第一个Key的文件夹路径
        if len(dirNames) > 0: dirNames.pop(0)
        for dirName in dirNames:
            if not os.path.exists(dirName):
                os.makedirs(dirName)
        # 按类型将字符串写入
        keys = sheet.col_values(0)
        for index, dirName in enumerate(dirNames):
            values = sheet.col_values(index + 1)
            iosDestFilePath = dirName + "/" + "strings.xml"
            iosFileManager = open(iosDestFilePath, "wb")
            stringEncoding = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<resources>\n"
            iosFileManager.write(stringEncoding)
            stringArrayValue = []
            stringArrayKey = ""
            for kIndex, key in enumerate(keys):
                if kIndex == 0 and key == "Keys":
                    continue
                # 匹配string-array
                match = re.match(r'(.*)\[(\d+)\]', key)
                if match is not None:
                    if len(stringArrayKey) == 0:
                        stringArrayKey = re.split(r'\[(.*)\]', key)[0].encode('utf-8')
                    else:
                        if key.startswith(stringArrayKey):
                            stringArrayValue.append(values[kIndex].encode('utf-8'))
                        else:
                            #写入string-array
                            iosFileManager.write(XmlStringArray(stringArrayKey, stringArrayValue))
                            #清空上次记录
                            stringArrayValue = []
                            #复制当前
                            stringArrayKey = re.split(r'\[(.*)\]', key)[0].encode('utf-8')
                            stringArrayValue.append(values[kIndex].encode('utf-8'))
                else:
                    # 写入string-array
                    if len(stringArrayKey) > 0 and len(stringArrayValue) > 0:
                        iosFileManager.write(XmlStringArray(stringArrayKey, stringArrayValue))
                        # 清空string-array数据
                        stringArrayValue = []
                        stringArrayKey = ""
                    # 写入string
                    cKey = key.encode('utf-8')
                    cValue = values[kIndex].encode('utf-8')
                    if cValue is None or cValue == '':
                        # print("Key:" + cKey + "\'s value is None. Index:" + str(kIndex + 1))
                        continue
                    content = "   <string name=\"" + cKey + "\">" + cValue + "</string>\n"
                    iosFileManager.write(content)
            iosFileManager.write("</resources>")
            iosFileManager.close()

def XmlStringArray(key, values):
    if len(key) > 0 and len(values) > 0:
        content = "   <string-array name=\"" + key + "\">\n"
        for item in values:
            content += "      <item>" + item + "</item>\n"
        content += "   </string-array>\n"
        return content

def main():
    options = addParser()
    if options.type is None:
        Log.error("type can not be empty!")
        return
    if options.file is None:
        Log.error("xls file can not be empty!")
        return
    if options.save is None:
        Log.error("save path can not be empty!")
        return

    if options.type == "ios":
        XlsToString(options.file, options.save)
        Log.info("Completed xls to iOS ! " + options.save)
    # XlsToString("/Users/yyg/Desktop/temp/ios.xlsx", "/Users/yyg/Desktop/temp")
    if options.type == "android":
        XlsToXml(options.file, options.save)
        Log.info("Completed xls to android ! " + options.save)
    # XlsToXml("/Users/yyg/Desktop/未命名文件夹/android.xlsx", "")

main()