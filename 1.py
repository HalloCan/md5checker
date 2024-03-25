# automatic file integrity verification script(MD5)
# Coded by HalloCandy, 2024-3-25 Version 1.0

import hashlib
import os
import xlrd
import xlwt


def MD5(filename):
    with open(filename, 'rb') as file:
        data = file.read()
    file_md5 = hashlib.md5(data).hexdigest()
    return file_md5


def readExcel(filename, status = 0):    # status == 0:default  ==1:for varify files.
    try:
        f = open(filename)
        f.close()
    except IOError:
        print("No such file:" + filename)
        exit(2)
    if status == 0:
        wb = xlrd.open_workbook(filename=filename)
        sheet = wb.sheet_by_index(0)
        datas = sheet.col_values(0)
        return datas
    elif status == 1:
        work_file = xlrd.open_workbook(filename=filename)
        sheet = work_file.sheet_by_index(0)
        paths = sheet.col_values(0)
        values = sheet.col_values(1)
        paths.pop(0)
        values.pop(0)
        return values, paths


def writeExcel(filename, datas, statuscode): # statuscode: 0:save the MD5; 1:save the file path;
                                             # 2:save the unmatched files(not now)
    if statuscode == 0:
        wb = xlwt.Workbook(filename)
        sheet = wb.add_sheet("Sheet1")
        # Set the column width 设置列列宽
        sheet.col(0).width = 512 * 20
        sheet.col(1).width = 512 * 20
        sheet.write(0, 0, "filename")
        sheet.write(0, 1, "MD5")
        md5list = list(datas.items())
        for i in range(1, len(md5list) + 1):
            sheet.write(i, 0, md5list[i - 1][0])
            sheet.write(i, 1, md5list[i - 1][1])
        wb.save(filename)
    elif statuscode == 1:
        wb = xlwt.Workbook(filename)
        sheet = wb.add_sheet("Sheet1")
        for i in range(0, len(datas)):
            sheet.write(i, 0, datas[i])
        wb.save(filename)


def setDB():
    print("*" * 64)
    print("Setting Database.")
    print("You should have an xls file containing your files.")
    print("If ready, press enter.")
    input("Press enter to continue...")
    print("OK!")


def getFileList():
    print("*" * 64)
    print("Getting File List.")
    inputPath = input("Input the path of the files(default:'./'):")
    if inputPath == "":
        path = "./"
    else:
        if os.path.exists(inputPath):
            path = inputPath
        else:
            print("Invalid Path")
            path = "./"

    # 获取当前目录下的所有文件
    fileList = [os.path.join(path, file) for file in os.listdir(path)]
    ls = []
    for file in fileList:
        if not os.path.isdir(file):
            ls.append(file)
    writeExcel('fileList.xls', ls, 1)


def getMD5():
    print("*" * 64)
    print("Getting MD5")
    print("You should have an xls file containing your files named 'fileList.xls'.")
    print("If ready, press enter.")
    input("Press enter to continue...")
    try:
        f = open(filename := "./fileList.xls")
        f.close()
    except IOError:
        print("No such file:" + filename)
        exit(2)
    fileList = readExcel(filename)
    dict = {fileList[i]:MD5(fileList[i]) for i in range(len(fileList))}
    writeExcel('hash.xls', dict, 0)


def verify():
    print("*" * 64)
    print("You should have an xls file containing your files named 'hash.xls'.")
    print("If ready, press enter.")
    input("Press enter to continue...")
    try:
        f = open(filename := "./hash.xls")
        f.close()
    except IOError:
        print("No such file:" + filename)
        exit(2)
    orighash, origpath = readExcel(filename, status=1)
    erroritem = []
    notmatchitem = []
    notmatchitem_md5 = []
    for i in range(len(orighash)):
        try:
            f = open(origpath[i])
            f.close()
        except IOError:
            print("File: " + origpath[i] + " is not exist!")
            erroritem.append(origpath[i])
            continue
        newmd5 = MD5(origpath[i])
        if orighash[i] != newmd5 :
            notmatchitem.append(origpath[i])
            notmatchitem_md5.append(newmd5)
    if len(notmatchitem) == 0 and len(erroritem) == 0:
        print("\033[1;32mVerification passed\033[0m")
    else:
        print("\033[1;31mVerification failed:\033[0m")
        print("Not matched: " + str(len(notmatchitem)))
        for i in range(len(notmatchitem)):
            print(notmatchitem[i])
        print("*" * 32)
        print("No longer existed: " + str(len(erroritem)))
        for i in range(len(erroritem)):
            print(erroritem[i])


if __name__ == '__main__':
    print("WELCOME TO MD5 Verification Program")
    print("Set a database:A   Verify the file hash:B， Save the file list:C")
    while True:
        which = input("Please choose an option:")
        if which == "A" or which == "a":
            print("choose A")
            getMD5()
            break
        elif which == "B" or which == "b":
            print("choose B")
            verify()
            break
        elif which == "C" or which == "c":
            print("choose C")
            getFileList()
            break
        else:
            print("Invalid Input")

