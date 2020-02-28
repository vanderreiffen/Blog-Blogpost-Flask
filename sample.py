import os
import xlwings.constants
import datetime
import time
from shutil import copy
import shutil
import zipfile

import simplejson as json
import win32com.client as xlz


def convertColSTR(sht, column):
    c = int(column)
    if c>0:
        tmpCol = sht.Cells(1, c).Address
        col = tmpCol.split("$")[1]
        return col

def UpdateOneDateOneValue(scData, ik, varDate, obsSHT):

    currDate = obsSHT.Cells(1, 14).Value

    if str(obsSHT.Name).lower().find("d") > -1 or str(obsSHT.Name).lower().find("w") > -1 :
        if currDate is None:
            rangeObj = obsSHT.Range("N:N")
            rangeObj.EntireColumn.Insert()
            obsSHT.Cells(1, 14).Value = varDate.strftime('%m-%d-%Y')
            obsSHT.Cells.Range("N1").NumberFormat = "dd-mmm-yyyy"
            obsSHT.Columns("N:N").EntireColumn.AutoFit()
        else:
            currDate = obsSHT.Cells(1, 14).Value.replace(tzinfo=None)
            if currDate < varDate:
                rangeObj = obsSHT.Range("N:N")
                rangeObj.EntireColumn.Insert()
                obsSHT.Cells(1, 14).Value = varDate.strftime('%m-%d-%Y')
                obsSHT.Cells.Range("N1").NumberFormat = "dd-mmm-yyyy"
                obsSHT.Columns("N:N").EntireColumn.AutoFit()

        fndDate = obsSHT.Cells.Find(What=varDate.date().strftime("%#m/%#d/%Y"), LookAt=xlwings.constants.LookAt.xlWhole)
    else:
        if currDate is None:
            rangeObj = obsSHT.Range("N:N")
            rangeObj.EntireColumn.Insert()
            obsSHT.Cells(1, 14).Value = varDate.strftime('%m-%d-%Y')
            obsSHT.Cells.Range("N1").NumberFormat = "mmm-yyyy"
            obsSHT.Columns("N:N").EntireColumn.AutoFit()
        else:
            currDate = obsSHT.Cells(1, 14).Value.replace(tzinfo=None)
            if currDate < varDate:
                rangeObj = obsSHT.Range("N:N")
                rangeObj.EntireColumn.Insert()
                obsSHT.Cells(1, 14).Value = varDate.strftime('%m-%d-%Y')
                obsSHT.Cells.Range("N1").NumberFormat = "mmm-yyyy"
                obsSHT.Columns("N:N").EntireColumn.AutoFit()

        # fndDate = obsSHT.Cells.Find(What=varDate.timetuple(), LookAt=xlwings.constants.LookAt.xlPart)
        #fndDate = obsSHT.Cells.Find(What=varDate.date().strftime("%#m/%#d/%Y"), LookAt=xlwings.constants.LookAt.xlPart)
        fndDate = obsSHT.Cells.Find(What=varDate.date().strftime("%#m/%#d/%Y"), LookAt=xlwings.constants.LookAt.xlWhole)
    if fndDate is not None:
        if scData is not None:
            if str(scData).strip() != '':
                obsSHT.Cells(ik, fndDate.Column).Value = scData
        #else:
        #    obsSHT.Cells(ik, fndDate.Column).Value = ""


def UpdateOneDateOneValueOptimized(scData, ik, varDate, obsSHT, dictdatecol):
    colvar = 0
    if varDate in dictdatecol:
        # already have in dict
        colvar = int(dictdatecol[varDate])
    else:
        fndDate = obsSHT.Cells.Range("1:1").Find(What=varDate.date().strftime("%#m/%#d/%Y"),
                                                 LookAt=xlwings.constants.LookAt.xlWhole,
                                                 LookIn=xlwings.constants.FindLookIn.xlFormulas)
        fndDate = obsSHT.Cells.Range("1:1").Find(What=varDate.date().strftime("%#m/%#d/%Y"), LookAt=xlwings.constants.LookAt.xlWhole)
        # fndDate = obsSHT.Cells.Range("1:1").Find(What=varDate.date().strftime("%d/%m/%Y"), LookAt=xlwings.constants.LookAt.xlWhole)
        if fndDate is not None:
            colvar = fndDate.Column
            dictdatecol.update({varDate: int(colvar)})
        else:
            # case need insert new date
            currDate = obsSHT.Cells(1, 14).Value
            if currDate is None:
                # case first and only date
                rangeObj = obsSHT.Range("N:N")
                rangeObj.EntireColumn.Insert()
                obsSHT.Cells(1, 14).Value = varDate.strftime('%m-%d-%Y')
                if str(obsSHT.Name).lower().find("d") > -1 or str(obsSHT.Name).lower().find("w") > -1:
                    obsSHT.Cells.Range("N1").NumberFormat = "dd-mmm-yyyy"
                else:
                    obsSHT.Cells.Range("N1").NumberFormat = "mmm-yyyy"
                obsSHT.Columns("N:N").EntireColumn.AutoFit()
                colvar = 14
                dictdatecol.update({varDate: int(colvar)})
            else:
                # find the column to insert
                currDate = obsSHT.Cells(1, 14).Value.replace(tzinfo=None)
                currDateRange = obsSHT.Cells(1, 14)
                if currDate < varDate:
                    # case new date is the latest
                    dictdatecol.clear()
                    rangeObj = obsSHT.Range("N:N")
                    rangeObj.EntireColumn.Insert()
                    obsSHT.Cells(1, 14).Value = varDate.strftime('%m-%d-%Y')
                    if str(obsSHT.Name).lower().find("d") > -1 or str(obsSHT.Name).lower().find("w") > -1:
                        obsSHT.Cells.Range("N1").NumberFormat = "dd-mmm-yyyy"
                    else:
                        obsSHT.Cells.Range("N1").NumberFormat = "mmm-yyyy"
                    obsSHT.Columns("N:N").EntireColumn.AutoFit()
                    colvar = 14
                    dictdatecol.update({varDate: int(colvar)})
                else:
                    while True:
                        currDateRange = currDateRange.Offset(1,2)
                        currDate = currDateRange.Value
                        if currDate is None:
                            dictdatecol.clear()
                            colvar = currDateRange.Column
                            rangeObj = obsSHT.Range(convertColSTR(obsSHT, colvar)+ ':' + convertColSTR(obsSHT, colvar))
                            rangeObj.EntireColumn.Insert()
                            obsSHT.Cells(1, colvar).Value = varDate.strftime('%m-%d-%Y')
                            if str(obsSHT.Name).lower().find("d") > -1 or str(obsSHT.Name).lower().find("w") > -1:
                                obsSHT.Cells.Range(convertColSTR(obsSHT, colvar) + "1").NumberFormat = "dd-mmm-yyyy"
                            else:
                                obsSHT.Cells.Range(convertColSTR(obsSHT, colvar) + "1").NumberFormat = "mmm-yyyy"
                            obsSHT.Columns(convertColSTR(obsSHT, colvar)+":"+convertColSTR(obsSHT, colvar)).EntireColumn.AutoFit()
                            dictdatecol.update({varDate: int(colvar)})
                            break
                        elif currDate.replace(tzinfo=None) < varDate:
                            dictdatecol.clear()
                            colvar = currDateRange.Column
                            rangeObj = obsSHT.Range(convertColSTR(obsSHT, colvar) + ':' + convertColSTR(obsSHT, colvar))
                            rangeObj.EntireColumn.Insert()
                            obsSHT.Cells(1, colvar).Value = varDate.strftime('%m-%d-%Y')
                            if str(obsSHT.Name).lower().find("d") > -1 and str(obsSHT.Name).lower().find("w") > -1:
                                obsSHT.Cells.Range(convertColSTR(obsSHT, colvar) + "1").NumberFormat = "dd-mmm-yyyy"
                            else:
                                obsSHT.Cells.Range(convertColSTR(obsSHT, colvar) + "1").NumberFormat = "mmm-yyyy"
                            obsSHT.Columns(convertColSTR(obsSHT, colvar) + ":" + convertColSTR(obsSHT,colvar)).EntireColumn.AutoFit()
                            dictdatecol.update({varDate: int(colvar)})
                            break

    if colvar > 0:
        if scData is not None:
            if str(scData).strip() != '':
                while True:
                    try:
                        obsSHT.Cells(ik, colvar).Value = scData
                        break
                    except:
                        time.sleep(0.1)


        #else:
        #    obsSHT.Cells(ik, fndDate.Column).Value = ""

def clearCellFormat(sht):
    last_Row = sht.UsedRange.Rows.Count
    for i in range(2, last_Row + 1):
        sheet = sht.Range(str(i) + ':' + str(i))
        sheet.Font.Bold = False
        sheet.Interior.ColorIndex = None

def copyInputFile(inputpath, targetpath):
   #eg: targetpath = X:\Macros\Programmer Team\New Automation\MAC\Python
   #eg: inputpath= X:\Macros\Programmer Team\New Automation\MAC\Python\AMCM\Interbank Middle Rates\MAC-AMCM-1430038.xlsx
    obsfullpath = ""
    inputname = str(os.path.basename(inputpath))
   #remove input file if exists
    if os.path.exists(targetpath + "/" + inputname):
        try:
            os.remove(targetpath + "/" + inputname)  # try to remove it directly
        except OSError as e:
            message = "Failed to copy the file"
            # print(message)


    if os.path.exists(targetpath):
        copy(inputpath, targetpath)
        obsfullpath = targetpath + "/" + inputname

    return obsfullpath


def get_number(num1):
    tempNum = ""
    for i in range(0, len(num1)):
        tmp = str(num1[i])
        if tmp.isdigit():
            if tempNum != "":
                tempNum = tempNum + str(tmp)
            else:
                tempNum = str(tmp)

    return tempNum

def getTime():
    timenow = datetime.datetime.utcnow().strftime("%I:%M:%S%p")
    return timenow

import importlib, importlib.util #module_from_file
#copy from https://docs.python.org/3/library/importlib.html#importing-a-source-file-directly
def module_from_file(module_name, file_path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

def getmonth(inputtmpMonth):
    tmpMonth = inputtmpMonth.lower()
    if tmpMonth.find("jan") >= 0 or tmpMonth.find("january") >= 0:
        month = 1
    elif tmpMonth.find("feb") >= 0 or tmpMonth.find("february") >= 0:
        month = 2
    elif tmpMonth.find("mar") >= 0 or tmpMonth.find("march") >= 0 or tmpMonth.find("1q") >= 0:
        month = 3
    elif tmpMonth.find("apr") >= 0 or tmpMonth.find("april") >= 0:
        month = 4
    elif tmpMonth.find("may") >= 0 or tmpMonth.find("may") >= 0:
        month = 5
    elif tmpMonth.find("jun") >= 0 or tmpMonth.find("june") >= 0 or tmpMonth.find("2q") >= 0:
        month = 6
    elif tmpMonth.find("jul") >= 0 or tmpMonth.find("july") >= 0:
        month = 7
    elif tmpMonth.find("aug") >= 0 or tmpMonth.find("august") >= 0:
        month = 8
    elif tmpMonth.find("sep") >= 0 or tmpMonth.find("sep") >= 0 or tmpMonth.find("3q") >= 0:
        month = 9
    elif tmpMonth.find("oct") >= 0 or tmpMonth.find("october") >= 0:
        month = 10
    elif tmpMonth.find("nov") >= 0 or tmpMonth.find("nov") >= 0:
        month = 11
    elif tmpMonth.find("dec") >= 0 or tmpMonth.find("december") >= 0 or tmpMonth.find("4q") >= 0:
        month = 12
    else:
        month = 0

    return int(month)

# sh is worksheet,
def paste_htmltable_excel(sh, tabhtml, startrow):
    workingrow = startrow
    for rowhtml in tabhtml.find_elements_by_tag_name("tr"):
        workingcol = 1
        for cellhtml in rowhtml.find_elements_by_xpath("child::*"):
            sh.Cells(workingrow, workingcol).Value = str(cellhtml.text).strip()
            workingcol = workingcol + 1
        workingrow = workingrow + 1

    return workingrow + 1

# sh is worksheet, using get_attribute
def paste_htmltable_excel_2(sh, tabhtml, startrow):
    workingrow = startrow
    for rowhtml in tabhtml.find_elements_by_tag_name("tr"):
        workingcol = 1
        for cellhtml in rowhtml.find_elements_by_xpath("child::*"):
            sh.Cells(workingrow, workingcol).Value = str(cellhtml.get_attribute('innerText')).strip()
            workingcol = workingcol + 1
        workingrow = workingrow + 1

    return workingrow + 1

def paste_htmltable_excel_colrowspan(sh, tabhtml, startrow):
    workingrow = startrow
    dicttoskip = {}

    for rowhtml in tabhtml.find_elements_by_tag_name("tr"):
        workingcol = 1
        for cellhtml in rowhtml.find_elements_by_xpath("child::*"):
            mycolspan = 1
            myrowspan = 1
            while True:
                if str(workingrow) + ":" + str(workingcol) in dicttoskip:
                    workingcol = workingcol + 1
                else:
                    break

            rowspanatt = cellhtml.get_attribute("rowspan")
            colspanatt = cellhtml.get_attribute("colspan")

            sh.Cells(workingrow, workingcol).Value = str(cellhtml.text).strip()
            if rowspanatt is None:
                if colspanatt is not None:
                    mycolspan = int(colspanatt)
            else:
                # need store to dictionary
                if colspanatt is not None:
                    mycolspan = int(colspanatt)
                myrowspan = int(rowspanatt)

            for i in range(workingrow, workingrow + myrowspan):
                for j in range(workingcol, workingcol + mycolspan):
                    tostore = str(i) + ":" + str(j)
                    if tostore not in dicttoskip:
                        dicttoskip.update({tostore:''})

            workingcol = workingcol + 1
        workingrow = workingrow + 1

    return workingrow + 1


def getSize(filename):
    if os.path.isfile(filename):
        st = os.stat(filename)
        return st.st_size
    else:
        return -1

def wait_download(file_path):
    current_size = getSize(file_path)
    time.sleep(5)  # wait download
    while current_size !=getSize(file_path) or getSize(file_path)==0:
        current_size =getSize(file_path)
        time.sleep(5)# wait download

def wait_file_complete(folder_path):
    # check file exist
    fail_count = 0
    while True:
        if os.listdir(folder_path) == []:
            time.sleep(1)
            fail_count = fail_count + 1
            if fail_count > 30:
                return False
        else:
            myfilename = folder_path + "\\" + os.listdir(folder_path)[0]
            break

    # check until the size do not change
    wait_download(myfilename)
    return True

def count_rows_htmltable(tabhtml):
    ctr =0
    for rowhtml in tabhtml.find_elements_by_tag_name("tr"):
        ctr += 1

    return int(ctr)


def sort_dict_by_key(serie_dict):
    dict = {}
    temp_dict= {}
    dict = serie_dict
    arr = [None] * len(dict)
    try:
        # for i in range(0,len(dict)-1):
        #     arr[i] = dict.keys()[i]
        i =0
        for var in dict.keys():
            arr[i] = var
            i += 1
        # bubble sort
        for i in range(0, len(arr)):
            for j in range(0,len(arr)):
                tmp1 = str(arr[i]).split(':')[0].strip()
                tmp2 = str(arr[j]).split(':')[0].strip()
                if tmp1 < tmp2:
                    tmp = arr[j]
                    arr[j] = arr[i]
                    arr[i] = tmp

        for i  in range(0, len(arr)):
            keyval = arr[i]
            temp_dict.update({keyval: i+ 1})
    except Exception as e:
        s = str(e)

    return temp_dict

def getmonth_TH(tmpSTRMnth):
    tmpMonth = str(tmpSTRMnth).lower().strip()

    if tmpMonth == 'ม.ค.' or tmpMonth == 'มกราคม' or tmpMonth == 'มกรำคม':
        month = '1'

    elif tmpMonth == 'ก.พ.' or tmpMonth == 'กมภาพนธ' or tmpMonth == 'กมภำพนธ' \
            or tmpMonth == 'กุมภาพันธ์':
        month = '2'

    elif tmpMonth == 'มี.ค.' or tmpMonth == 'มนาคม' or tmpMonth == 'มนำคม' \
            or tmpMonth == 'มีนาคม':
        month = '3'

    elif tmpMonth == 'เม.ย.' or tmpMonth == 'เมษายน' or tmpMonth == 'เมษำยน':
        month = '4'

    elif tmpMonth == 'พ.ค.' or tmpMonth == 'พฤษภาคม' or tmpMonth == 'พฤษภำคม':
        month = '5'

    elif tmpMonth == 'ม.ย.' or tmpMonth == 'มิ.ย.' or tmpMonth == 'มถนายน' \
            or tmpMonth == 'มถนำยน' or tmpMonth == 'มิถุนายน':
        month = '6'

    elif tmpMonth == 'ก.ค.' or tmpMonth == 'กรกฎาคม' or tmpMonth == 'กรกฎำคม' \
            or tmpMonth == 'เดอนกรกฎาคม':
        month = '7'

    elif tmpMonth == 'ส.ค.' or tmpMonth == 'สิงหาคม' or tmpMonth == 'สงหาคม':
        month = '8'

    elif tmpMonth == 'ก.ย.' or tmpMonth == 'กนยายน' or tmpMonth == 'กนยำยน' \
            or tmpMonth == 'ประจำเดือนสิงหาคม' or tmpMonth == 'กันยายน' or tmpMonth == 'เดอนกนยายน':
        month = '9'

    elif tmpMonth == 'ต.ค.' or tmpMonth == 'ตลำคม' or tmpMonth == 'ตุลาคม' or tmpMonth == 'ตลาคม':
        month = '10'

    elif tmpMonth == 'พ.ย.' or tmpMonth == 'พฤศจกายน' or tmpMonth == 'พฤศจกำยน' \
            or tmpMonth == 'พฤศจิกายน':
        month = '11'

    elif tmpMonth == 'ธ.ค.' or tmpMonth == 'ธนวาคม' or tmpMonth == 'ธนวำคม' \
            or tmpMonth == 'เดือนธันวาคม' or tmpMonth == 'ธันวาคม':
        month = '12'

    else:
        month = '0'

    return int(month)

def getmonth_french(inputtmpMonth):
    tmpMonth = inputtmpMonth.lower().strip()
    tmpMonth = tmpMonth[:4]
    if tmpMonth.find("jan") >= 0 or tmpMonth.find("janv") >= 0:
        month = 1
    elif tmpMonth.find("fév") >= 0 or tmpMonth.find("fev") >= 0 or tmpMonth.find("febr") >= 0:
        month = 2
    elif tmpMonth.find("mar") >= 0 or tmpMonth.find("mars") >= 0:
        month = 3
    elif tmpMonth.find("avr") >= 0 or tmpMonth.find("apr") >= 0:
        month = 4
    elif tmpMonth.find("mai") >= 0 or tmpMonth.find("may") >= 0:
        month = 5
    elif tmpMonth.find("juin") >= 0 or tmpMonth.find("june") >= 0:
        month = 6
    elif tmpMonth.find("juil") >= 0 or tmpMonth.find("july") >= 0:
        month = 7
    elif tmpMonth.find("aoû") >= 0 or tmpMonth.find("aug") >= 0:
        month = 8
    elif tmpMonth.find("sep") >= 0:
        month = 9
    elif tmpMonth.find("oct") >= 0:
        month = 10
    elif tmpMonth.find("nov") >= 0:
        month = 11
    elif tmpMonth.find("déc") >= 0 or tmpMonth.find("dec") >= 0:
        month = 12
    else:
        month = 0

    return int(month)

def getJPNstartYear(tmpYear, period):

    ctr = 1; startYear = 0; currYear = 0
    if str(period).lower() == 'reiwa' or str(period).lower().strip().find('r') == 0 or period.strip().find('令和')>=0:
        startYear = 2019
    elif str(period).lower() == 'heisei' or str(period).lower().strip().find('h') == 0 or period.strip().find('平成')>=0:
        startYear = 1989
    elif str(period).lower() == 'showa' or str(period).lower().strip().find('s') == 0 or period.strip().find('昭和')>=0:
        startYear = 1926
    elif str(period).lower() == 'taisho':
        startYear = 1912
    elif str(period).lower() == 'meiji':
        startYear = 1868
    else:
        startYear = 0

    if startYear > 0:
        currYear = convGregorian_JPN(tmpYear, startYear)

    return (currYear)

######################    1      2019
def convGregorian_JPN(JpnYear, starYear):
    ctr = 1;startYear = 0;currYear = 0
    currYear = starYear
    if JpnYear == "":
        JpnYear = 1
    elif JpnYear.isnumeric():
        JpnYear = int(JpnYear)

    while ctr < JpnYear:
        currYear = currYear + 1
        ctr = ctr + 1

    return (currYear)

##########################################################################################
## sourcefile = source file path
## filenameinzipfile = file name in the zip folder
## if filenameinzipfile empty, it will extract all file in zip folder
##########################################################################################

def unzipfile(sourcefile, filenameinzipfile):
    newSourceFile = ""
    try:
        newFileName = os.path.basename(sourcefile).replace(".zip", "")
        tmpPath = os.path.dirname(sourcefile) + "\\temp_" + datetime.datetime.utcnow().strftime("%Y%m%d_%H%M%S")
        newSourceFile = ""
        if not os.path.exists(tmpPath):
            os.makedirs(tmpPath)

        zip_ref = zipfile.ZipFile(sourcefile, 'r')
        # zip_ref.infolist()

        zip_ref.extractall(tmpPath)
        zip_ref.close()
        filename = pathfile = curFile = newPath = allfilePath = ""
    except:
        newSourceFile = ""
        return newSourceFile

    if filenameinzipfile != "":
        try:
            for root, dirs, files in os.walk(tmpPath):
                if str(files).lower().find(str(filenameinzipfile).lower()) >= 0:
                    for file in files:
                        if filenameinzipfile.strip() == file.split(".")[0].strip():
                            filename = file
                            pathfile = root
                            break
                    if filename != "" and pathfile != "":
                        break

            if filename != "" and pathfile != "":
                format = "." + filename.split(".")[1].strip()
                if os.path.exists(pathfile +"\\" + filename):
                    curName = pathfile + "\\" + filename
                    newName = pathfile + "\\" + newFileName + format

                    try:
                        os.rename(curName, newName)
                    except:
                        raise

                    newSourceFile = os.path.dirname(sourcefile) + '/' + newFileName + format

                    if os.path.exists(newSourceFile):
                        os.remove(newSourceFile)

                    shutil.copy(newName, os.path.dirname(sourcefile))
        except:
            newSourceFile = ""
    else:
        try:
            firstTime = True
            for root, dirs, files in os.walk(tmpPath):
                if files != []:
                    for file in files:
                        newFolder = newFileName
                        currFile = root + "\\" + file

                        if firstTime == True:
                            newPath = os.path.dirname(sourcefile) + "\\" + newFolder
                            if os.path.exists(newPath):
                                shutil.rmtree(newPath)

                            if not os.path.exists(newPath):
                                os.makedirs(newPath)

                            firstTime = False

                        shutil.copy(currFile, newPath)

                        if os.path.exists(newPath + "\\" + file):

                            if allfilePath == "":
                                allfilePath = newPath + "\\" + file
                            else:
                                allfilePath = allfilePath + ";" + newPath + "\\" + file

            if allfilePath != "":
                newSourceFile = allfilePath

        except:
            newSourceFile = ""

    if os.path.exists(tmpPath):
        shutil.rmtree(tmpPath)

    return newSourceFile

# def unzipfile(sourcefile, filenameinzipfile):
#     newFileName = os.path.basename(sourcefile).replace(".zip", "")
#     tmpPath = os.path.dirname(sourcefile) + "\\temp_" + datetime.datetime.utcnow().strftime("%Y%m%d_%H%M%S")
#     newSourceFile = ""
#     if not os.path.exists(tmpPath):
#         os.makedirs(tmpPath)
#
#     zip_ref = zipfile.ZipFile(sourcefile, 'r')
#     # zip_ref.infolist()
#
#     zip_ref.extractall(tmpPath)
#     zip_ref.close()
#     filename = ""
#     if filenameinzipfile != "":
#         for root, dirs, files in os.walk(tmpPath):
#             if str(files).lower().find(str(filenameinzipfile).lower()) >=0:
#                 filename = str(files[0])
#                 break
#
#         if filename != "":
#             format = "." + filename.split(".")[1].strip()
#             if os.path.exists(tmpPath +"\\" + filename):
#                 curName = tmpPath + "\\" + filename
#                 newName = tmpPath + "/" + newFileName + format
#
#                 try:
#                     os.rename(curName, newName)
#                 except:
#                     raise
#
#                 newSourceFile = os.path.dirname(sourcefile) + '/' + newFileName + format
#
#                 if os.path.exists(newSourceFile):
#                     os.remove(newSourceFile)
#
#                 shutil.copy(newName, os.path.dirname(sourcefile))
#     else:
#         for root, dirs, files in os.walk(tmpPath):
#             filename = str(files[0])
#             break
#
#         if filename != "":
#             format = "." + filename.split(".")[1].strip()
#             if os.path.exists(tmpPath +"\\" + filename):
#                 curName = tmpPath + "\\" + filename
#                 newName = tmpPath + "/" + newFileName + format
#
#                 try:
#                     os.rename(curName, newName)
#                 except:
#                     raise
#
#                 newSourceFile = os.path.dirname(sourcefile) + '/' + newFileName + format
#
#                 if os.path.exists(newSourceFile):
#                     os.remove(newSourceFile)
#
#                 shutil.copy(newName, os.path.dirname(sourcefile))
#
#     if os.path.exists(tmpPath):
#         shutil.rmtree(tmpPath)
#
#     return newSourceFile

def json_convert_1(s, strscbk):
    date_dict = {}
    seri_dict = {}
    val_dict = {}
    res =False
    scbk = None
    sh = None
    d_ctr = 0; s_ctr = 0
    # dataform = str(s).strip("'<>() ").replace('\'','\"')
    # dataform = s.decode('utf-8').replace('\0','')
    try:
        res = json.loads(s)
        iyr = "" ; im =""
        for it in res['records']:
            # dates
            tmp = str(it['time']).strip()
            if is_number(left(tmp,4)):
                iyr = left(tmp,4)

            if tmp.find('H1') >= 0:
                im = 6
            elif tmp.find('H2') >= 0:
                im = 12
            elif is_number(right(tmp,4)):
                im = 12
            else:
                im = getmonth(tmp.split(' ')[1].strip())


            if im != "" and iyr != "":
                idate= datetime.datetime(int(iyr), int(im), 1)
                if it['time'] not in date_dict:
                    d_ctr = d_ctr + 1
                    date_dict.update({it['time']: str(d_ctr) + '|' + str(idate)})

            # series names
            tmp = it['variableCode'] + ':' + it["variableName"]
            if tmp != "":
                if tmp not in seri_dict:
                    s_ctr = s_ctr + 1
                    seri_dict.update({tmp : s_ctr})
            # values
            tmp = it["value"]
            tmp2 = it['variableCode'] + ':' + it["variableName"] + '|' + it['time']
            tmp2 = str(tmp2).strip()
            if tmp2 not in val_dict:
                val_dict.update({tmp2: tmp})
                res = True


        if res == True:
            sf = xlz.gencache.EnsureDispatch("Excel.Application")
            scbk = sf.Workbooks.Add()
            sh = scbk.Sheets.Add()
            sh.Name = "Data"
            tic = 2;
            tir = 2
            scbk.SaveAs(strscbk.replace("/", '\\'))
            sh.Cells(1, 1).Value = "Variables"
        else:
            return False

        seri_dict = sort_dict_by_key(seri_dict)
        if len(val_dict) > 0 and  len(seri_dict) > 0 and len(date_dict) > 0:
            res= False
            for tvar in (date_dict.keys()):
                str_temp =  date_dict[tvar]
                sh.Cells(1,tic).Value = tvar
                tic += 1

            for tvar in seri_dict.keys():
                str_temp = str(tvar).split(':')[1].strip()
                sh.Cells(tir,1).Value = str_temp
                tir += 1

            for tvar in val_dict.keys():
                str_temp = str(tvar).split('|')[0].strip()
                str_temp2 = str(tvar).split('|')[1].strip()
                try:
                    str_temp = seri_dict[str_temp]
                    str_temp2 = date_dict[str_temp2]
                    str_temp2 = str(str_temp2).split('|')[0].strip()
                    tir = int(str_temp) + 1
                    tic = int(str_temp2) + 1
                    sh.Cells(tir,tic).Value = val_dict[tvar]
                    res = True
                except KeyError:
                    res = False
            if res == True:
                scbk.Close(SaveChanges=True)
                return res

    except Exception as e:
        str(e)
        return False

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

def left(s, amount):
    return s[:amount]

def right(s, amount):
    return s[-amount:]

def mid(s, offset, amount):
    return s[offset:offset+amount]


def PrepareHistoricalDateColumns(obsSHT, dictdatestoinsert):
    # only fill for date earlier than oldest existing
    # need to call before any updateonedateonevalue for series
    currDate = obsSHT.Cells(1, 14).Value
    dictsortreversed = sorted(dictdatestoinsert,  reverse=True)

    if currDate is None:
        currDateRange = obsSHT.Cells(1, 14)
        for varDate in dictsortreversed:
            # case we found a date which is smaller than the earliest date
            colvar = currDateRange.Column
            rangeObj = obsSHT.Range(convertColSTR(obsSHT, colvar) + ':' + convertColSTR(obsSHT, colvar))
            rangeObj.EntireColumn.Insert()
            obsSHT.Cells(1, colvar).Value = varDate.strftime('%m-%d-%Y')
            if str(obsSHT.Name).lower().find("d") > -1 or str(obsSHT.Name).lower().find("w") > -1:
                obsSHT.Cells.Range(convertColSTR(obsSHT, colvar) + "1").NumberFormat = "dd-mmm-yyyy"
            else:
                obsSHT.Cells.Range(convertColSTR(obsSHT, colvar) + "1").NumberFormat = "mmm-yyyy"
            obsSHT.Columns(
                convertColSTR(obsSHT, colvar) + ":" + convertColSTR(obsSHT, colvar)).EntireColumn.AutoFit()
            #currDateRange = currDateRange.Offset(1, 2)
    else:
        # find the last column
        currDate = obsSHT.Cells(1, 14).Value.replace(tzinfo=None)
        currDateRange = obsSHT.Cells(1, 14)

        while True:
            currDateRange = currDateRange.Offset(1, 2)
            currDate = currDateRange.Value
            if currDate is None:
                break

        currDateRange = currDateRange.Offset(1,0)
        # find the latest date in dictionary which is older than the earliest existing date
        oldestdate = currDateRange.Value.replace(tzinfo=None)
        currDateRange = currDateRange.Offset(1, 2)
        for varDate in dictsortreversed:
            if varDate < oldestdate:
                # case we found a date which is smaller than the earliest date
                colvar = currDateRange.Column
                rangeObj = obsSHT.Range(convertColSTR(obsSHT, colvar) + ':' + convertColSTR(obsSHT, colvar))
                rangeObj.EntireColumn.Insert()
                obsSHT.Cells(1, colvar).Value = varDate.strftime('%m-%d-%Y')
                if str(obsSHT.Name).lower().find("d") > -1 or str(obsSHT.Name).lower().find("w") > -1:
                    obsSHT.Cells.Range(convertColSTR(obsSHT, colvar) + "1").NumberFormat = "dd-mmm-yyyy"
                else:
                    obsSHT.Cells.Range(convertColSTR(obsSHT, colvar) + "1").NumberFormat = "mmm-yyyy"
                obsSHT.Columns(
                    convertColSTR(obsSHT, colvar) + ":" + convertColSTR(obsSHT, colvar)).EntireColumn.AutoFit()
                #currDateRange = currDateRange.Offset(1, 2)


## added recently dec 19 @kevin
def find_end(sht, rng, dd='right'):
    if dd == 'right':
        rng = sht.Cells(rng.Row, rng.Column + 1)
        while rng.Text == '' and rng.Column < 1000:
            rng = sht.Cells(rng.Row, rng.Column + 1)
    elif dd== 'left':
        rng = sht.Cells(rng.Row, rng.Column - 1)
        while rng.Text == '' and rng.Column > 0:
            rng = sht.Cells(rng.Row, rng.Column - 1)
    elif dd== 'down':
        rng = sht.Cells(rng.Row+1, rng.Column)
        while rng.Text == '' and rng.Row < 1000:
            rng = sht.Cells(rng.Row+1, rng.Column)
    elif dd == 'up':
        rng = sht.Cells(rng.Row-1, rng.Column)
        while rng.Text == '' and rng.Row > 0:
            rng = sht.Cells(rng.Row-1, rng.Column)
    return rng

def only_digits(s):
    tmpstr = str(s)
    scdata = ""
    for ctr in range(0, len(tmpstr)):
        if is_number(str(tmpstr)[ctr]):
            scdata = scdata + s[ctr]

    return scdata

def is_year(iyr):
    try:
        if len((only_digits(iyr))) == 4:
            return True
        else:
            return  False
    except ValueError:
        return  False


def find_mapping(scsht, tofind, delim=':', st_row=1, srch_order='by_row', look_at='part',
                 srch_direction='next', rngFind = None, rngFindBool = False):
    dc = None
    if rngFindBool:
        dc = rngFind.Cells(1, 1)
    else:
        dc = scsht.Cells(st_row, 1)

    search_order_dict = {'by_row': xlwings.constants.SearchOrder.xlByRows,
                         'by_col': xlwings.constants.SearchOrder.xlByColumns}
    lookin_dict = {'values': xlwings.constants.FindLookIn.xlValues, 'formulas': xlwings.constants.FindLookIn.xlFormulas,
                   'comments': xlwings.constants.FindLookIn.xlComments}
    after_dict = {'next': xlwings.constants.SearchDirection.xlNext,
                  'prev': xlwings.constants.SearchDirection.xlPrevious}
    lookat_dict = {'part': xlwings.constants.LookAt.xlPart, 'whole': xlwings.constants.LookAt.xlWhole}

    for tmpdata in tofind.split(delim):
        temp = str(tmpdata).strip().lower().replace('(-)','') #remove the offset indicator
        if rngFindBool:
            if dc is not  None:
                dc = rngFind.Find(What=temp.strip(), LookAt=lookat_dict[look_at],
                                      SearchOrder=search_order_dict[srch_order],
                                      MatchCase=False, After=dc, SearchDirection=after_dict[srch_direction],
                                      )

        else:
            if dc is not None:
                dc = scsht.Cells.Find(What=temp.strip(), LookAt=lookat_dict[look_at],
                                      SearchOrder=search_order_dict[srch_order],
                                      MatchCase=False, After=dc, SearchDirection= after_dict[srch_direction],
                                         )

        for ctr in range(str(tmpdata).count('(-)')): #move it to left
            dc = rng_offset(scsht,dc)

    outputRng = None
    if dc is not None:
        outputRng = dc
    return outputRng


def range_limit_pdf(sh, find_start, find_end,st_row = 1):
    #start range
    res = False
    try:
        rngstart = find_mapping(sh,find_start,st_row = st_row)
        rngend = find_mapping(sh, find_end, st_row = rngstart.Row + 1, srch_order='by_row')
        if rngend.Row < rngstart.Row + 10:
            rngend = find_mapping(sh, find_end, st_row = rngend.Row + 1, srch_order='by_row')
        outrng = sh.Range( str(rngstart.Row) + ":" + str(rngend.Row))
        res = True
        return  (res ,outrng)
    except Exception as e:
        print(str(e))
        return (res,None)

def rev_srch(sh, tofind, st=1):
    dc = None
    dc = sh.Cells(st, 1)
    dc = sh.Cells.Find(What=tofind.strip(), LookAt=xlwings.constants.LookAt.xlPart,
                          SearchOrder=xlwings.constants.SearchOrder.xlByRows,
                          MatchCase=False, After=dc, SearchDirection= xlwings.constants.SearchDirection.xlPrevious)
    return dc




def rng_offset(sh, rng, direction='left', offset_count =1):
    if direction == 'left':
        rng = sh.Cells(rng.Row, rng.Column - offset_count)
        return  rng
    if direction == 'right':
        rng = sh.Cells(rng.Row, rng.Column + offset_count)
        return  rng
    if direction == 'up':
        rng = sh.Cells(rng.Row - offset_count, rng.Column )
        return  rng
    if direction == 'down':
        rng = sh.Cells(rng.Row + 1, rng.Column)
        return  rng


def is_date(iyr, im):
    try:
        if (int(iyr) >= 1900 and int(iyr) <= 2500) and (int(im) >= 1 and int(im) <= 12):
            return True
        else:
            return False
    except ValueError:
        return False


def find_cell(scsht, tofind, srow=1, scol=1, lrow=1, lcol=1, search_order='by_row', lookat='part',
              search_sheet='advance', lookin='formulas', after_cell='next', delim=':'):
    """
    :param scsht: source sheet object #cannot be None
    :param tofind: string #cannot be None
    :param srow: start row, integer #default 1
    :param scol: start column, integer #default 1
    :param lrow: last row, integer, #default 1
    :param lcol: last column, integer #default 1
    :param search_order: 'by_row' or 'by_col', string #default 'by_row'
    :param search_sheet: 'advance' or 'basic' or 'iter' if basic no lrow or lcol needed. if iter, the find has delimiter #default is advance
    :param lookin: 'formulas','values' or 'comments' #default formulas
    :param after_cell: 'next','previous' or 'comments' #default next
    :param delim: for iteratables #default :
    :return: dict['row'] and dict['col']
    """
    search_order_dict = {'by_row': xlwings.constants.SearchOrder.xlByRows,
                         'by_col': xlwings.constants.SearchOrder.xlByColumns}
    lookin_dict = {'values': xlwings.constants.FindLookIn.xlValues, 'formulas': xlwings.constants.FindLookIn.xlFormulas,
                   'comments': xlwings.constants.FindLookIn.xlComments}
    after_dict = {'next': xlwings.constants.SearchDirection.xlNext,
                  'prev': xlwings.constants.SearchDirection.xlPrevious}
    lookat_dict = {'part': xlwings.constants.LookAt.xlPart, 'whole': xlwings.constants.LookAt.xlWhole}

    after_rng = scsht.Cells(srow, scol)

    if search_sheet == 'advance':
        cell = scsht.Range(scsht.Cells(srow, scol), scsht.Cells(lrow, lcol)).Find(What=tofind.strip(),
                                                                                  LookAt=lookat_dict['whole'],
                                                                                  LookIn=lookin_dict[lookin],
                                                                                  SearchOrder=search_order_dict[
                                                                                      search_order], MatchCase=False)
        if cell is None:
            cell = scsht.Range(scsht.Cells(srow, scol), scsht.Cells(lrow, lcol)).Find(What=tofind.strip(),
                                                                                      LookAt=lookat_dict['part'],
                                                                                      LookIn=lookin_dict[lookin],
                                                                                      SearchOrder=search_order_dict[
                                                                                          search_order],
                                                                                      MatchCase=False)
        print('{} at , row = {}, col = {}'.format(tofind, cell.Row, cell.Column))
    elif search_sheet == 'basic':
        cell = scsht.Cells.Find(What="*", After=scsht.Cells(1, 1),SearchOrder=search_order_dict[search_order],SearchDirection=after_dict['prev'])
        
    elif search_sheet == 'iter':
        cell = after_rng
        for temp in tofind.split(delim):
            if cell is not None:
                if lrow > 1:
                    cell = scsht.Range(str(srow) + ":" + str(lrow)).Find(What=temp.strip().replace('(-)', ''),
                                                                         LookAt=lookat_dict[lookat], After=cell,
                                                                         SearchOrder=search_order_dict[search_order],
                                                                         SearchDirection=after_dict[after_cell])
                    if '(-)' in temp.lower():
                        cell = rng_offset(scsht, cell)
                else:
                    cell = scsht.Cells.Find(What=temp.strip(), LookAt=lookat_dict[lookat], After=cell,
                                            SearchOrder=search_order_dict[search_order],
                                            SearchDirection=after_dict[after_cell])

    cellpos = {'row': cell.Row, 'col': cell.Column}
    return cellpos

def find_via_columnset(sh, tofind, st_row=1, st_col=1, s_row= 1, s_col=1, ret ='col'):
    e_row = st_row + s_row
    e_col = st_col + s_col
    try:
        rngFind = sh.Range(sh.Cells(st_row,st_col),sh.Cells(e_row,e_col))
        if ret=='col':
            res_col = find_mapping(sh,tofind,srch_order='by_col', rngFind=rngFind ,rngFindBool=True).Column
        else:
            res_col = find_mapping(sh, tofind, srch_order='by_col', rngFind=rngFind, rngFindBool=True).Row
        return  res_col
    except:
        return  0
        
def remove_non_digits(input):
    output = ''.join(c for c in input if c.isdigit())
    return output
