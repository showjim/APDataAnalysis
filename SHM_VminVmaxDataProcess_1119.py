# -*- coding:utf-8 -*-

import xlwt
import xlrd
from xlutils.copy import copy
import time
import datetime
from xlwt import *
import os
from xlrd import open_workbook

ticks = time.localtime()
print(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
itime = datetime.datetime.now()
CurrTime = "_M" + str(itime.month) + "_D" + str(itime.day) + "_hr" + str(itime.hour) + "_Min" + str(
    itime.minute) + "_Sec" + str(itime.second)
print(str(itime.day) + str(itime.hour) + str(itime.minute))
print(CurrTime)
print(itime.minute)

"""
DLG=win32ui.CreateFileDialog(1)
DLG.SetOFNInitialDir("C:\Py_Test\SHMOO")
DLG.DoModal()
filename=DLG.GetPathName()
print((filename))
"""
xflag = 0


def GetFileName(file_dir, *args):
    i = 0
    for root, dirs, files in os.walk(file_dir):
        # print (root)
        #  print (dirs)
        # print (files[0])
        i = i + 1
    # print (files)
    return (files)


# return (root)


print("1-->SHM\n" + "2-->VminVmax\n")
KeyWordInFileName = input("LogType(SHM / VminVmax): ")
# InputPath=os.getcwd()+'\\'+KeyWordInFileName+CurrTime
# folder = os.path.exists(InputPath)
# if not folder:
#    os.makedirs(InputPath)
if KeyWordInFileName == "SHM":
    InputPath = os.getcwd() + '\SHMOO'
    print(InputPath)
    tmpPath = InputPath
    ReportPath = os.getcwd() + "\SHMOO\DatalogProcess.xls"
elif KeyWordInFileName == "VminVmax":
    InputPath = os.getcwd() + '\VminVmax'
    tmpPath = InputPath
    print(InputPath)
    ReportPath = tmpPath + "\DatalogProcess.xls"

inputfiles = GetFileName(InputPath)
# print(inputfiles)
filecnt = len(inputfiles)
print("Total Files Count =  " + str(filecnt))

is_ChooseDiffSiteNumFromDiffFiles = input("Process Different Site Num( True / False ) : ")
if is_ChooseDiffSiteNumFromDiffFiles != "True":
    site = input("site num: ")
    sites = site.split(',')
# print(len(sites))

# GenMultiFiles=input(" 0:GenToMultiFile,1: GenToOneExcelFile:  ") # 1: means get all input txt files into one excel file
GenMultiFiles = 1
is_TER_Log = False
is_93K_Log = False
Is_StartPrint = False
ValidFileCnt = -1

shtName = "Sht_" + KeyWordInFileName

shtsCnt = 0
ExistedShts = ""
shmXpoints = 11 + 8
n = 70000
InsertBlankLines = 0
SiteSpace = 25
Activefilecnt = 0
fileSites = [8, 8, 8, 8, 8, 8, 8, 8, 8, 8]
TotalTestItemsPerSite = 0

for FileIdx in range(filecnt):
    result = (KeyWordInFileName in inputfiles[FileIdx]) and (
            ".txt" in inputfiles[FileIdx] or ".TXT" in inputfiles[FileIdx])
    if result == True:
        Activefilecnt = Activefilecnt + 1
        if is_ChooseDiffSiteNumFromDiffFiles == "True":
            fileSites[Activefilecnt - 1] = input(
                "(" + inputfiles[FileIdx] + ")" + "File" + str(Activefilecnt - 1) + "__Sites: ")
# ------------------------------
if os.path.exists(ReportPath):
    os.remove(ReportPath)
"""
if not(os.path.exists(ReportPath)):
    print("++++++++++  NOP ++++++++++")
    
     bk0 = xlwt.Workbook(ReportPath)
     if GenMultiFiles==0:
         for siteIdx in range(len(sites)):
            shtName=KeyWordInFileName+"_Site"+str(sites[siteIdx])
            sht0=bk0.add_sheet(shtName,True)
         sht0=bk0.add_sheet("SHM_AllSites",True)
         sht0 = bk0.add_sheet("SHM_AllSites_Summary", True)
         bk0.save(ReportPath)
     elif GenMultiFiles==1:
         for i in range(Activefilecnt):
             for siteIdx in range(len(sites)):
                 shtName = KeyWordInFileName +"_File"+str(i)+ "_Site" + str(sites[siteIdx])
                 sht0 = bk0.add_sheet(shtName, True)
         sht0 = bk0.add_sheet("SHM_AllSites", True)
         sht0 = bk0.add_sheet("SHM_AllSites_Summary", True)
         bk0.save(ReportPath)
    

else:
    bk0=xlrd.open_workbook(ReportPath)
    shtsCnt=len(bk0.sheets())
    print (shtsCnt)
    #bk_copy = copy(bk0)
    for i in range(shtsCnt):
        #print(bk0.sheet_by_index(i).name)
        ExistedShts=ExistedShts+','+bk0.sheet_by_index(i).name
    print (ExistedShts)
    if shtName in ExistedShts:
        os.remove(ReportPath)
        bk_copy=xlwt.Workbook()
        shtName = shtName+"_Site"+str(site)
        newSht = bk_copy.add_sheet(shtName, False)

    else:
        bk_copy=copy(bk0)
        print (shtName)
        shtName = shtName +"_Site"+str(site)
        newSht=bk_copy.add_sheet(shtName,False)
    bk_copy.save(ReportPath)
    print ("Done")
print(ReportPath)
"""


def GenerateWorksheet(ReportPath, Activefilecnt):
    bk0 = xlwt.Workbook(ReportPath)
    if GenMultiFiles == 0:
        for siteIdx in range(len(sites)):
            shtName = KeyWordInFileName + "_Site" + str(sites[siteIdx])
            sht0 = bk0.add_sheet(shtName, True)
        sht0 = bk0.add_sheet(KeyWordInFileName + "_AllSites", True)
        sht0 = bk0.add_sheet(KeyWordInFileName + "_AllSites_Summary", True)
        bk0.save(ReportPath)
    elif GenMultiFiles == 1:
        for i in range(Activefilecnt):
            for siteIdx in range(len(sites)):
                shtName = KeyWordInFileName + "_File" + str(i) + "_Site" + str(sites[siteIdx])
                sht0 = bk0.add_sheet(shtName, True)
        sht0 = bk0.add_sheet(KeyWordInFileName + "_AllSites", True)
        sht0 = bk0.add_sheet(KeyWordInFileName + "_AllSites_Summary", True)
        bk0.save(ReportPath)


def GenerateCurrentWorksheet(ReportPath, filenum):
    if filenum != 0:
        bk0 = xlrd.open_workbook(ReportPath, formatting_info=True)
        bk_copy0 = copy(bk0)
    elif filenum == 0:
        bk0 = xlwt.Workbook(ReportPath)
    # print("okay")
    if GenMultiFiles == 0:
        for siteIdx in range(len(sites)):
            shtName = KeyWordInFileName + "_Site" + str(sites[siteIdx])
            sht0 = bk0.add_sheet(shtName, True)
        sht0 = bk0.add_sheet(KeyWordInFileName + "_AllSites", True)
        sht0 = bk0.add_sheet(KeyWordInFileName + "_AllSites_Summary", True)
        bk0.save(ReportPath)
    elif GenMultiFiles == 1:
        if filenum != 0:
            for siteIdx in range(len(sites)):
                shtName = KeyWordInFileName + "_File" + str(filenum) + "_Site" + str(sites[siteIdx])
                sht0 = bk_copy0.add_sheet(shtName, True)
            bk_copy0.save(ReportPath)
        else:
            sht0 = bk0.add_sheet(KeyWordInFileName + "_AllSites", True)
            sht0 = bk0.add_sheet(KeyWordInFileName + "_AllSites_Summary", True)
            for siteIdx in range(len(sites)):
                shtName = KeyWordInFileName + "_File" + str(filenum) + "_Site" + str(sites[siteIdx])
                sht0 = bk0.add_sheet(shtName, True)
            bk0.save(ReportPath)


if is_ChooseDiffSiteNumFromDiffFiles != "True":
    GenerateWorksheet(ReportPath, Activefilecnt)

rowIdx = 1
Idx = 0
columnIdx = 1
AllSites = -1
ValidFileCnt = 0
TitleWritten = ""
style = XFStyle()
Key = 0
MarkFileName = 0
VminVmaxStartlineNum = 0
xlUsedRowsCnt = 0

iws = ["a", "b", "c", "d", "e", "f", "g", "h", "j"]

if KeyWordInFileName.upper() == "SHM":
    for FileIdx in range(filecnt):
        is_LogToExcel = 0
        is_SplitStr = 0
        if GenMultiFiles == 0:
            ValidFileCnt = 0
        rowIdx = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        startPrint = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        columnIdx = 1
        ExecutedSiteCnt = 0
        linescnt = 0
        tmpLineNum0 = 0
        tmpLineNum1 = 0
        TestItemName = ""

        """
        ReportOutputPath = os.getcwd() + "\SHMOO\DatalogProcess.xls"
        if GenMultiFiles == 1 and  ValidFileCnt!=0:
            ReportOutputPath = InputPath + "\SHMOO_" + KeyWordInFileName + "_" + CurrTime + "_ALL" + ".xls"
        
        bk0 = xlrd.open_workbook(ReportOutputPath,formatting_info=True)
        print ("-----"+inputfiles[FileIdx]+"----  "+ str(FileIdx)+ "  --------")
        bk1 = copy(bk0)
        """

        if "ADV" in inputfiles[FileIdx]:
            Platform = "ADV"
        elif "TER" in inputfiles[FileIdx]:
            Platform = "TER"
        result = (KeyWordInFileName in inputfiles[FileIdx]) and (
                    ".txt" in inputfiles[FileIdx] or ".TXT" in inputfiles[FileIdx])
        # print (result)
        if result == True:
            is_LogToExcel = 1
            ValidFileCnt = ValidFileCnt + 1
        # print ("      *Need Process This File ? *  "+ str(result)+"\n")
        if is_LogToExcel == 1:
            print("*   *   *   *   *   *   *   *   *   *")
            print(
                " File Num:  " + str(ValidFileCnt - 1) + "   ---" + inputfiles[FileIdx] + " ---Site Num: " + fileSites[
                    ValidFileCnt - 1])
            print("*   *   *   *   *   *   *   *   *   *")

            if GenMultiFiles == 1 and (ValidFileCnt - 1) != 0:
                ReportOutputPath = tmpPath + "\SHMOO_" + KeyWordInFileName + "_" + CurrTime + "_ALL" + ".xls"
                if is_ChooseDiffSiteNumFromDiffFiles == "True":
                    # site = input("site num: ")  # add 20181019
                    site = fileSites[ValidFileCnt - 1]
                    sites = site.split(',')  # add 20181019
                    GenerateCurrentWorksheet(ReportOutputPath, ValidFileCnt - 1)
            elif GenMultiFiles == 1 and (ValidFileCnt - 1) == 0:
                ReportOutputPath = ReportPath
                if is_ChooseDiffSiteNumFromDiffFiles == "True":
                    # site = input("site num: ")  # add 20181019
                    site = fileSites[ValidFileCnt - 1]
                    sites = site.split(',')  # add 20181019
                    GenerateCurrentWorksheet(ReportOutputPath, ValidFileCnt - 1)
                    ReportOutputPath = tmpPath + "\DatalogProcess.xls"

            bk0 = xlrd.open_workbook(ReportOutputPath, formatting_info=True)
            # print("-----" + inputfiles[FileIdx] + "----  " + str(FileIdx) + "  --------")
            bk1 = copy(bk0)
            # print("*   *   *   *   *   *   *   *   *   *")
            # print(" File Num:  "+ str(ValidFileCnt-1)+"   ---"+inputfiles[FileIdx])
            # print("*   *   *   *   *   *   *   *   *   *")

            for siteIdx in range(len(sites)):
                AllSites = AllSites + 1
                path = InputPath + '\\' + inputfiles[FileIdx]
                file = open(path, 'r', encoding='utf-8')

                if GenMultiFiles == 0:
                    iws[siteIdx] = bk1.get_sheet(KeyWordInFileName + "_Site" + str(sites[siteIdx]))
                elif GenMultiFiles == 1:
                    xx = KeyWordInFileName + "_File" + str(FileIdx) + "_Site" + str(sites[siteIdx])
                    iws[siteIdx] = bk1.get_sheet(
                        KeyWordInFileName + "_File" + str(ValidFileCnt - 1) + "_Site" + str(sites[siteIdx]))

                # sht_AllSites=bk1.get_sheet("SHM_AllSites")
                sht_AllSites = bk1.get_sheet(KeyWordInFileName + "_AllSites")
                rowIdx[siteIdx] = 0
                i = 0
                TotalTestItemsPerSite = 0

                for Strline in file.readlines()[0:n]:
                    # print (lineIdx)
                    tmpStr = Strline.strip()
                    linescnt = linescnt + 1
                    # if "Site Failed tests/Executed tests" in Strline:
                    # print ("total lines in current file:"+str(linescnt))
                    # break
                    if Platform == "TER":
                        # if tmpStr.endswith("(Site"+str(sites[siteIdx])+"):") and "SHM_" in tmpStr:  # for old version
                        if "_SHM:" in tmpStr:
                            Key = 1
                            TestItemName = tmpStr
                            tmpLineNum0 = linescnt

                            # iws[siteIdx].write(rowIdx[siteIdx], columnIdx, Strline)
                            # rowIdx[siteIdx] = rowIdx[siteIdx] + 1
                        if "Site: " + str(sites[siteIdx]) in tmpStr and Key == 1:
                            startPrint[siteIdx] = 1
                            TotalTestItemsPerSite = TotalTestItemsPerSite + 1
                            iws[siteIdx].write(0, columnIdx, TotalTestItemsPerSite)
                            iws[siteIdx].write(0, columnIdx + 1, inputfiles[FileIdx])
                            rowIdx[siteIdx] = rowIdx[siteIdx] + 1

                            if linescnt - tmpLineNum0 < 5:
                                iws[siteIdx].write(rowIdx[siteIdx], columnIdx, TestItemName)
                                if GenMultiFiles == 0:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + (siteIdx) * SiteSpace,
                                                       TestItemName)
                                elif GenMultiFiles == 1:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace,
                                                       TestItemName)
                                rowIdx[siteIdx] = rowIdx[siteIdx] + 1

                        if "Tcoef(%)" in Strline:
                            is_SplitStr = 1
                        if startPrint[siteIdx] == 1:
                            # print(Strline)
                            # if MarkFileName==0 and rowIdx[siteIdx]==0:
                            # iws[siteIdx].write(0, columnIdx, inputfiles[FileIdx])
                            if is_SplitStr == 0:
                                iws[siteIdx].write(rowIdx[siteIdx], columnIdx, Strline)
                                if GenMultiFiles == 0:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + (siteIdx) * SiteSpace, Strline)
                                elif GenMultiFiles == 1:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace, Strline)
                                    # print ("#############################################################")
                                    # print (columnIdx + i + AllSites * SiteSpace)
                                    # sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + siteIdx * SiteSpace, listStrline[i])
                            elif is_SplitStr == 1:
                                # print(Strline)
                                listStrline = Strline.split("\t")

                                # print(listStrline)
                                for i in range(len(listStrline)):
                                    # print(listStrline[i]+"\n")
                                    """
                                    if GenMultiFiles==0:
                                        sht_AllSites.write(rowIdx[siteIdx], columnIdx +i+ siteIdx * SiteSpace, listStrline[i])
                                    elif GenMultiFiles==1:
                                        sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace,listStrline[i])
                                """
                                    if listStrline[i].strip() == "P":
                                        pattern = Pattern()
                                        pattern.pattern = Pattern.SOLID_PATTERN
                                        pattern.pattern_fore_colour = 3
                                        style.pattern = pattern
                                        iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i, listStrline[i], style)
                                        if GenMultiFiles == 0:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + siteIdx * SiteSpace,
                                                               listStrline[i], style)
                                        elif GenMultiFiles == 1:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace,
                                                               listStrline[i], style)
                                    elif listStrline[i].strip() == ".":
                                        pattern = Pattern()
                                        pattern.pattern = Pattern.SOLID_PATTERN
                                        pattern.pattern_fore_colour = 2
                                        style.pattern = pattern
                                        iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i, listStrline[i], style)
                                        if GenMultiFiles == 0:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + siteIdx * SiteSpace,
                                                               listStrline[i], style)
                                        elif GenMultiFiles == 1:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace,
                                                               listStrline[i], style)
                                    elif listStrline[i].strip() == "*":
                                        pattern = Pattern()
                                        pattern.pattern = Pattern.SOLID_PATTERN
                                        pattern.pattern_fore_colour = 5
                                        style.pattern = pattern
                                        iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i, listStrline[i], style)
                                        if GenMultiFiles == 0:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + siteIdx * SiteSpace,
                                                               listStrline[i], style)
                                        elif GenMultiFiles == 1:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace,
                                                               listStrline[i], style)
                                    else:
                                        iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i, listStrline[i])
                                        if GenMultiFiles == 0:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + siteIdx * SiteSpace,
                                                               listStrline[i])
                                        elif GenMultiFiles == 1:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace,
                                                               listStrline[i])
                            rowIdx[siteIdx] = rowIdx[siteIdx] + 1
                            xlUsedRowsCnt = rowIdx[siteIdx] + 100


                    # $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
                    elif Platform == "ADV":
                        InsertBlankLines = 0
                        if "SHMOO_RESULT:" in tmpStr:
                            Key = 1
                            TestItemName = tmpStr
                            tmpLineNum0 = linescnt

                        if "CURRENT SITE NUMBER" in Strline and tmpStr[len(tmpStr) - 1] == sites[siteIdx] and Key == 1:
                            startPrint[siteIdx] = 1
                            rowIdx[siteIdx] = rowIdx[siteIdx] + 2
                            TotalTestItemsPerSite = TotalTestItemsPerSite + 1
                            iws[siteIdx].write(0, columnIdx, TotalTestItemsPerSite)
                            rowIdx[siteIdx] = rowIdx[siteIdx] + 1

                            if linescnt - tmpLineNum0 < 5:
                                iws[siteIdx].write(rowIdx[siteIdx], columnIdx, TestItemName)
                                if GenMultiFiles == 0:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + (siteIdx) * SiteSpace,
                                                       TestItemName)
                                elif GenMultiFiles == 1:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace,
                                                       TestItemName)
                                rowIdx[siteIdx] = rowIdx[siteIdx] + 1
                        if "Tcoef(nS)" in Strline or "Port_Period@pCLK_19p2M(nS)" in Strline:
                            is_SplitStr = 1
                        if startPrint[siteIdx] == 1:
                            # print(Strline)
                            if is_SplitStr == 0:
                                iws[siteIdx].write(rowIdx[siteIdx], columnIdx, Strline)
                                if GenMultiFiles == 0:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + (siteIdx) * SiteSpace, Strline)
                                elif GenMultiFiles == 1:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace, Strline)
                                    # print ("#############################################################")
                                    # print (columnIdx + i + AllSites * SiteSpace)
                            elif is_SplitStr == 1:
                                # print(Strline)
                                listStrline = Strline.split("\t")
                                # print(listStrline)
                                for i in range(len(listStrline)):
                                    """
                                    if GenMultiFiles==0:
                                        sht_AllSites.write(rowIdx[siteIdx], columnIdx +i+1+ siteIdx * SiteSpace, listStrline[i])
                                    elif GenMultiFiles==1:
                                        sht_AllSites.write(rowIdx[siteIdx],columnIdx + i +1+ AllSites * SiteSpace,listStrline[i])
                                  """
                                    if listStrline[i].strip() == "P":
                                        pattern = Pattern()
                                        pattern.pattern = Pattern.SOLID_PATTERN
                                        pattern.pattern_fore_colour = 3
                                        style.pattern = pattern
                                        iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i + 1, listStrline[i], style)
                                        if GenMultiFiles == 0:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + 1 + siteIdx * SiteSpace,
                                                               listStrline[i], style)
                                        elif GenMultiFiles == 1:
                                            sht_AllSites.write(rowIdx[siteIdx],
                                                               columnIdx + i + 1 + AllSites * SiteSpace,
                                                               listStrline[i], style)
                                    elif listStrline[i].strip() == ".":
                                        pattern = Pattern()
                                        pattern.pattern = Pattern.SOLID_PATTERN
                                        pattern.pattern_fore_colour = 2
                                        style.pattern = pattern
                                        iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i + 1, listStrline[i], style)
                                        if GenMultiFiles == 0:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + 1 + siteIdx * SiteSpace,
                                                               listStrline[i], style)
                                        elif GenMultiFiles == 1:
                                            sht_AllSites.write(rowIdx[siteIdx],
                                                               columnIdx + i + 1 + AllSites * SiteSpace,
                                                               listStrline[i], style)
                                    elif listStrline[i].strip() == "*":
                                        pattern = Pattern()
                                        pattern.pattern = Pattern.SOLID_PATTERN
                                        pattern.pattern_fore_colour = 5
                                        style.pattern = pattern
                                        iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i + 1, listStrline[i], style)
                                        if GenMultiFiles == 0:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + 1 + siteIdx * SiteSpace,
                                                               listStrline[i], style)
                                        elif GenMultiFiles == 1:
                                            sht_AllSites.write(rowIdx[siteIdx],
                                                               columnIdx + i + 1 + AllSites * SiteSpace,
                                                               listStrline[i], style)
                                    else:
                                        iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i + 1, listStrline[i])
                                        if GenMultiFiles == 0:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + 1 + siteIdx * SiteSpace,
                                                               listStrline[i])
                                        elif GenMultiFiles == 1:
                                            sht_AllSites.write(rowIdx[siteIdx],
                                                               columnIdx + i + 1 + AllSites * SiteSpace,
                                                               listStrline[i])
                            rowIdx[siteIdx] = rowIdx[siteIdx] + 1
                            xlUsedRowsCnt = rowIdx[siteIdx] + 100
                        # break
                    if Strline.strip() == "":
                        if startPrint[siteIdx] == 1:
                            rowIdx[siteIdx] = rowIdx[siteIdx] + InsertBlankLines
                        startPrint[siteIdx] = 0
                        is_SplitStr = 0
                        Key = 0
            # for  i in range(255):
            sht_AllSites.write(xlUsedRowsCnt, 255, "Done")

            if GenMultiFiles == 0:
                tmp = inputfiles[FileIdx][0:len(inputfiles[FileIdx]) - 4] + "_Site"
                for siteIdx in range(len(sites)):
                    tmp = tmp + "_" + str(sites[siteIdx])
                ReportOutPath = InputPath + "\SHMOO_" + KeyWordInFileName + "_" + CurrTime + "_" + tmp + ".xls"
                bk1.save(ReportPath)
            elif GenMultiFiles == 1:
                ReportOutputPath = InputPath + "\SHMOO_" + KeyWordInFileName + "_" + CurrTime + "_ALL" + ".xls"
                bk1.save(ReportOutputPath)

            # to combine data to one table
            # ReportPath==InputPath+"\DatalogProcess"+ ".xls"
            workbook = open_workbook(ReportOutputPath, formatting_info=True)
            # worksheet=workbook.sheet_by_name("SHM_AllSites")
            worksheet = workbook.sheet_by_name(KeyWordInFileName + "_AllSites")
            tmpStr = ""
            # for rowIdx in range(12,worksheet.nrows):
            print("test")
            print(str(worksheet.ncols))

            bk = copy(workbook)
            # SHM_AllSites_Summary=bk.get_sheet("SHM_AllSites_Summary")
            AllSites_Summary = bk.get_sheet(KeyWordInFileName + "_AllSites_Summary")

            for rowIdx in range(0, xlUsedRowsCnt):
                for columnIdx in range(0, 25):
                    # tmpStr = worksheet.cell_value(rowIdx, columnIdx)
                    if worksheet.cell_value(rowIdx, 1) != "" and columnIdx == 1:
                        tmpStr = worksheet.cell_value(rowIdx, 1)
                        AllSites_Summary.write(rowIdx, 1, tmpStr)
                    if worksheet.cell_value(rowIdx, 2) != "" and columnIdx == 2:
                        tmpStr = worksheet.cell_value(rowIdx, 2)
                        AllSites_Summary.write(rowIdx, 2, tmpStr)

                    if str(worksheet.cell_value(rowIdx, columnIdx)).strip() == "P" or str(
                            worksheet.cell_value(rowIdx, columnIdx)).strip() == "." or str(
                            worksheet.cell_value(rowIdx, columnIdx)).strip() == "*" or str(
                            worksheet.cell_value(rowIdx, columnIdx)).strip() == "":
                        tmpStr = worksheet.cell_value(rowIdx, columnIdx)
                        for k in range(len(sites) + (Activefilecnt - 1) * GenMultiFiles * len(sites)):
                            # print(worksheet.cell_value(rowIdx,columnIdx))
                            tmpStr = tmpStr + str(worksheet.cell_value(rowIdx, columnIdx + (k + 1) * 25))
                            if ("." in tmpStr and "P" in tmpStr) or ("." in tmpStr and "*" in tmpStr) or (
                                    "#" in tmpStr and "*" in tmpStr) or ("P" in tmpStr and "*" in tmpStr) or (
                                    "#" in tmpStr and "*" in tmpStr) or ("#" in tmpStr and "P" in tmpStr):

                                pattern = Pattern()
                                pattern.pattern = Pattern.SOLID_PATTERN
                                pattern.pattern_fore_colour = 2
                                style.pattern = pattern
                                AllSites_Summary.write(rowIdx, columnIdx, tmpStr, style)

                            elif "." in tmpStr and not ("P" in tmpStr) and not ("#" in tmpStr) and not ("*" in tmpStr):
                                pattern = Pattern()
                                pattern.pattern = Pattern.SOLID_PATTERN
                                pattern.pattern_fore_colour = 7
                                style.pattern = pattern
                                AllSites_Summary.write(rowIdx, columnIdx, tmpStr, style)
                            else:
                                AllSites_Summary.write(rowIdx, columnIdx, tmpStr)
                            # print(tmpStr)
                    elif columnIdx != 1 or columnIdx != 2:
                        tmpStr = worksheet.cell_value(rowIdx, columnIdx)
                        AllSites_Summary.write(rowIdx, columnIdx, tmpStr)

            print("Done")

            if GenMultiFiles == 0:
                tmp = inputfiles[FileIdx][0:len(inputfiles[FileIdx]) - 4] + "_Site"
                for siteIdx in range(len(sites)):
                    tmp = tmp + "_" + str(sites[siteIdx])
                ReportOutPath = InputPath + "\SHMOO_" + KeyWordInFileName + "_" + CurrTime + "_" + tmp + ".xls"
                bk.save(ReportOutPath)
            elif GenMultiFiles == 1:
                ReportOutputPath = InputPath + "\SHMOO_" + KeyWordInFileName + "_" + CurrTime + "_ALL" + ".xls"
                bk.save(ReportOutputPath)

    # """
elif KeyWordInFileName.upper() == "VMINVMAX":
    for FileIdx in range(filecnt):
        is_LogToExcel = 0
        is_SplitStr = 0
        if GenMultiFiles == 0:
            ValidFileCnt = 0
        rowIdx = [0, 0, 0, 0, 0, 0, 0, 0]
        startPrint = [0, 0, 0, 0, 0, 0, 0, 0]
        columnIdx = 1
        ExecutedSiteCnt = 0
        linescnt = 0
        tmpLineNum0 = 0
        tmpLineNum1 = 0
        TestItemName = ""

        """
        ReportOutputPath = os.getcwd() + "\SHMOO\DatalogProcess.xls"
        if GenMultiFiles == 1 and  ValidFileCnt!=0:
            ReportOutputPath = InputPath + "\SHMOO_" + KeyWordInFileName + "_" + CurrTime + "_ALL" + ".xls"
    
        bk0 = xlrd.open_workbook(ReportOutputPath,formatting_info=True)
        print ("-----"+inputfiles[FileIdx]+"----  "+ str(FileIdx)+ "  --------")
        bk1 = copy(bk0)
        """

        if "ADV" in inputfiles[FileIdx]:
            Platform = "ADV"
        elif "TER" in inputfiles[FileIdx]:
            Platform = "TER"
        result = (KeyWordInFileName in inputfiles[FileIdx]) and (
                ".txt" in inputfiles[FileIdx] or ".TXT" in inputfiles[FileIdx])
        # print (result)
        if result == True:
            is_LogToExcel = 1
            ValidFileCnt = ValidFileCnt + 1
        # print ("      *Need Process This File ? *  "+ str(result)+"\n")
        if is_LogToExcel == 1:
            print("*   *   *   *   *   *   *   *   *   *")
            print(" File Num:  " + str(ValidFileCnt - 1) + "   ---" + inputfiles[FileIdx])
            print("*   *   *   *   *   *   *   *   *   *")

            if GenMultiFiles == 1 and (ValidFileCnt - 1) != 0:
                ReportOutputPath = tmpPath + "\VminVmax_" + KeyWordInFileName + "_" + CurrTime + "_ALL" + ".xls"
                if is_ChooseDiffSiteNumFromDiffFiles == "True":
                    # site = input("site num: ")  # add 20181019
                    site = fileSites[ValidFileCnt - 1]
                    sites = site.split(',')  # add 20181019
                    GenerateCurrentWorksheet(ReportOutputPath, ValidFileCnt - 1)
            elif GenMultiFiles == 1 and (ValidFileCnt - 1) == 0:
                ReportOutputPath = ReportPath
                if is_ChooseDiffSiteNumFromDiffFiles == "True":
                    # site = input("site num: ")  # add 20181019
                    site = fileSites[ValidFileCnt - 1]
                    sites = site.split(',')  # add 20181019
                    GenerateCurrentWorksheet(ReportOutputPath, ValidFileCnt - 1)
                    ReportOutputPath = tmpPath + "\DatalogProcess.xls"

            bk0 = xlrd.open_workbook(ReportOutputPath, formatting_info=True)
            bk1 = copy(bk0)

            for siteIdx in range(len(sites)):
                AllSites = AllSites + 1
                path = InputPath + '\\' + inputfiles[FileIdx]
                file = open(path, 'r', encoding='utf-8')

                if GenMultiFiles == 0:
                    iws[siteIdx] = bk1.get_sheet(KeyWordInFileName + "_Site" + str(sites[siteIdx]))
                elif GenMultiFiles == 1:
                    xx = KeyWordInFileName + "_File" + str(FileIdx) + "_Site" + str(sites[siteIdx])
                    iws[siteIdx] = bk1.get_sheet(
                        KeyWordInFileName + "_File" + str(ValidFileCnt - 1) + "_Site" + str(sites[siteIdx]))

                sht_AllSites = bk1.get_sheet(KeyWordInFileName + "_AllSites")
                sht_AllSites_Summary_tmp = bk1.get_sheet(KeyWordInFileName + "_AllSites_Summary")
                rowIdx[siteIdx] = 0
                i = 0
                TotalTestItemsPerSite = 0

                for Strline in file.readlines()[0:n]:
                    # print (lineIdx)
                    tmpStr = Strline.strip()
                    linescnt = linescnt + 1
                    # if "Site Failed tests/Executed tests" in Strline:
                    # iws[siteIdx].write(rowIdx[siteIdx]+10, 1, TotalTestItemsPerSite)

                    # break
                    if Platform == "TER":
                        # if tmpStr.endswith("(Site"+str(sites[siteIdx])+"):") and "SHM_" in tmpStr:  # for old version
                        if "_Vmin_Vmax:" in tmpStr:
                            Key = 1
                            TestItemName = tmpStr
                            tmpLineNum0 = linescnt

                            # iws[siteIdx].write(rowIdx[siteIdx], columnIdx, Strline)
                            # rowIdx[siteIdx] = rowIdx[siteIdx] + 1
                        if "Site: " + str(sites[siteIdx]) in tmpStr and Key == 1:
                            startPrint[siteIdx] = 1
                            TotalTestItemsPerSite = TotalTestItemsPerSite + 1
                            iws[siteIdx].write(0, columnIdx, TotalTestItemsPerSite)
                            rowIdx[siteIdx] = rowIdx[siteIdx] + 1
                            if linescnt - tmpLineNum0 < 5:
                                iws[siteIdx].write(rowIdx[siteIdx], columnIdx, TestItemName)
                                if GenMultiFiles == 0:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + (siteIdx) * SiteSpace,
                                                       TestItemName)
                                elif GenMultiFiles == 1:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace,
                                                       TestItemName)
                                rowIdx[siteIdx] = rowIdx[siteIdx] + 1

                        if "Cur_V_SPEC:" in Strline:
                            is_SplitStr = 1
                            VminVmaxStartlineNum = rowIdx[siteIdx]
                        if startPrint[siteIdx] == 1:
                            # print(Strline)
                            # if MarkFileName==0 and rowIdx[siteIdx]==0:
                            # iws[siteIdx].write(0, columnIdx, inputfiles[FileIdx])
                            if is_SplitStr == 0:
                                iws[siteIdx].write(rowIdx[siteIdx], columnIdx, Strline)
                                if GenMultiFiles == 0:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + (siteIdx) * SiteSpace, Strline)
                                elif GenMultiFiles == 1:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace, Strline)
                                    # print ("#############################################################")
                                    # print (columnIdx + i + AllSites * SiteSpace)
                                    # sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + siteIdx * SiteSpace, listStrline[i])
                            elif is_SplitStr == 1:
                                # print(Strline)
                                listStrline = Strline.split("\t")

                                # print(listStrline)
                                for i in range(len(listStrline)):
                                    # print(listStrline[i]+"\n")
                                    """
                                    if GenMultiFiles==0:
                                        sht_AllSites.write(rowIdx[siteIdx], columnIdx +i+ siteIdx * SiteSpace, listStrline[i])
                                    elif GenMultiFiles==1:
                                        sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace,listStrline[i])
                                """
                                    if listStrline[i].strip() == "P":
                                        pattern = Pattern()
                                        pattern.pattern = Pattern.SOLID_PATTERN
                                        pattern.pattern_fore_colour = 3
                                        style.pattern = pattern
                                        iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i, listStrline[i], style)
                                        if GenMultiFiles == 0:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + siteIdx * SiteSpace,
                                                               listStrline[i], style)
                                        elif GenMultiFiles == 1:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace,
                                                               listStrline[i], style)
                                    elif listStrline[i].strip() == ".":
                                        pattern = Pattern()
                                        pattern.pattern = Pattern.SOLID_PATTERN
                                        pattern.pattern_fore_colour = 2
                                        style.pattern = pattern
                                        iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i, listStrline[i], style)
                                        if GenMultiFiles == 0:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + siteIdx * SiteSpace,
                                                               listStrline[i], style)
                                        elif GenMultiFiles == 1:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace,
                                                               listStrline[i], style)
                                    elif listStrline[i].strip() == "*":
                                        pattern = Pattern()
                                        pattern.pattern = Pattern.SOLID_PATTERN
                                        pattern.pattern_fore_colour = 5
                                        style.pattern = pattern
                                        iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i, listStrline[i], style)
                                        if GenMultiFiles == 0:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + siteIdx * SiteSpace,
                                                               listStrline[i], style)
                                        elif GenMultiFiles == 1:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace,
                                                               listStrline[i], style)
                                    else:
                                        if ("Vmin=" in Strline):
                                            rowIdx[siteIdx] = VminVmaxStartlineNum + 7
                                            iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i, listStrline[i])
                                            rowIdx[siteIdx] = rowIdx[siteIdx] + 1
                                        else:
                                            iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i, listStrline[i])
                                        if GenMultiFiles == 0:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + siteIdx * SiteSpace,
                                                               listStrline[i])
                                        elif GenMultiFiles == 1:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace,
                                                               listStrline[i])
                            rowIdx[siteIdx] = rowIdx[siteIdx] + 1
                            xlUsedRowsCnt = rowIdx[siteIdx] + 100


                    # $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
                    elif Platform == "ADV":
                        InsertBlankLines = 0
                        if "CURRENT SITE NUMBER =" in tmpStr:
                            Key = 1
                            TestItemName = tmpStr
                            tmpLineNum0 = linescnt

                        if "CURRENT SITE NUMBER =" + str(sites[siteIdx]) in Strline and Key == 1:
                            startPrint[siteIdx] = 1
                            rowIdx[siteIdx] = rowIdx[siteIdx] + 2
                            TotalTestItemsPerSite = TotalTestItemsPerSite + 1
                            iws[siteIdx].write(0, columnIdx, TotalTestItemsPerSite)
                            rowIdx[siteIdx] = rowIdx[siteIdx] + 1
                            """
                            if linescnt - tmpLineNum0 < 5:
                                iws[siteIdx].write(rowIdx[siteIdx], columnIdx, TestItemName)
                                if GenMultiFiles == 0:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + (siteIdx) * SiteSpace, TestItemName)
                                elif GenMultiFiles == 1:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace, TestItemName)
                                rowIdx[siteIdx] = rowIdx[siteIdx] + 1
                                """
                        if "Cur_V_SPEC_value:" in Strline:
                            is_SplitStr = 1
                            VminVmaxStartlineNum = rowIdx[siteIdx]
                        if startPrint[siteIdx] == 1:
                            # print(Strline)
                            if is_SplitStr == 0:
                                iws[siteIdx].write(rowIdx[siteIdx], columnIdx, Strline)
                                if GenMultiFiles == 0:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + (siteIdx) * SiteSpace, Strline)
                                elif GenMultiFiles == 1:
                                    sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + AllSites * SiteSpace, Strline)
                                    # print ("#############################################################")
                                    # print (columnIdx + i + AllSites * SiteSpace)
                            elif is_SplitStr == 1:
                                # print(Strline)
                                listStrline = Strline.split("\t")
                                # print(listStrline)
                                for i in range(len(listStrline)):
                                    """
                                    if GenMultiFiles==0:
                                        sht_AllSites.write(rowIdx[siteIdx], columnIdx +i+1+ siteIdx * SiteSpace, listStrline[i])
                                    elif GenMultiFiles==1:
                                        sht_AllSites.write(rowIdx[siteIdx],columnIdx + i +1+ AllSites * SiteSpace,listStrline[i])
                                  """
                                    if listStrline[i].strip() == "P":
                                        pattern = Pattern()
                                        pattern.pattern = Pattern.SOLID_PATTERN
                                        pattern.pattern_fore_colour = 3
                                        style.pattern = pattern
                                        iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i + 1, listStrline[i], style)
                                        if GenMultiFiles == 0:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + 1 + siteIdx * SiteSpace,
                                                               listStrline[i], style)
                                        elif GenMultiFiles == 1:
                                            sht_AllSites.write(rowIdx[siteIdx],
                                                               columnIdx + i + 1 + AllSites * SiteSpace,
                                                               listStrline[i], style)
                                    elif listStrline[i].strip() == ".":
                                        pattern = Pattern()
                                        pattern.pattern = Pattern.SOLID_PATTERN
                                        pattern.pattern_fore_colour = 2
                                        style.pattern = pattern
                                        iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i + 1, listStrline[i], style)
                                        if GenMultiFiles == 0:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + 1 + siteIdx * SiteSpace,
                                                               listStrline[i], style)
                                        elif GenMultiFiles == 1:
                                            sht_AllSites.write(rowIdx[siteIdx],
                                                               columnIdx + i + 1 + AllSites * SiteSpace,
                                                               listStrline[i], style)
                                    elif listStrline[i].strip() == "*":
                                        pattern = Pattern()
                                        pattern.pattern = Pattern.SOLID_PATTERN
                                        pattern.pattern_fore_colour = 5
                                        style.pattern = pattern
                                        iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i + 1, listStrline[i], style)
                                        if GenMultiFiles == 0:
                                            sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + 1 + siteIdx * SiteSpace,
                                                               listStrline[i], style)
                                        elif GenMultiFiles == 1:
                                            sht_AllSites.write(rowIdx[siteIdx],
                                                               columnIdx + i + 1 + AllSites * SiteSpace,
                                                               listStrline[i], style)
                                    else:
                                        if not ("Vmin=" in Strline or "Vmax=" in Strline):
                                            iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i + 1, listStrline[i])

                                            if GenMultiFiles == 0:
                                                sht_AllSites.write(rowIdx[siteIdx],
                                                                   columnIdx + i + 1 + siteIdx * SiteSpace,
                                                                   listStrline[i])
                                            elif GenMultiFiles == 1:
                                                sht_AllSites.write(rowIdx[siteIdx],
                                                                   columnIdx + i + 1 + AllSites * SiteSpace,
                                                                   listStrline[i])
                                        else:
                                            if "Vmin=" in Strline:
                                                rowIdx[siteIdx] = VminVmaxStartlineNum + 7
                                                iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i, listStrline[i])
                                                rowIdx[siteIdx] = rowIdx[siteIdx] + 1
                                            elif "Vmax=" in Strline:
                                                iws[siteIdx].write(rowIdx[siteIdx], columnIdx + i, listStrline[i])
                                            if GenMultiFiles == 0:
                                                sht_AllSites.write(rowIdx[siteIdx], columnIdx + i + siteIdx * SiteSpace,
                                                                   listStrline[i])
                                            elif GenMultiFiles == 1:
                                                sht_AllSites.write(rowIdx[siteIdx],
                                                                   columnIdx + i + AllSites * SiteSpace,
                                                                   listStrline[i])
                            rowIdx[siteIdx] = rowIdx[siteIdx] + 1
                            xlUsedRowsCnt = rowIdx[siteIdx] + 100
                        # break
                    # if Strline.strip() == "":
                    if "Vmax=" in Strline:
                        if startPrint[siteIdx] == 1:
                            rowIdx[siteIdx] = rowIdx[siteIdx] + InsertBlankLines
                        startPrint[siteIdx] = 0
                        is_SplitStr = 0
                        Key = 0
            # for  i in range(255):
            sht_AllSites.write(xlUsedRowsCnt, 255, "Done")
            sht_AllSites_Summary_tmp.write(xlUsedRowsCnt, 255, "Done")
            if GenMultiFiles == 0:
                tmp = inputfiles[FileIdx][0:len(inputfiles[FileIdx]) - 4] + "_Site"
                for siteIdx in range(len(sites)):
                    tmp = tmp + "_" + str(sites[siteIdx])
                ReportOutPath = InputPath + "\VminVmax_" + KeyWordInFileName + "_" + CurrTime + "_" + tmp + ".xls"
                bk1.save(ReportPath)
            elif GenMultiFiles == 1:
                ReportOutputPath = tmpPath + "\VminVmax_" + KeyWordInFileName + "_" + CurrTime + "_ALL" + ".xls"
                bk1.save(ReportOutputPath)

            # to combine data to one table
            # ReportPath==InputPath+"\DatalogProcess"+ ".xls"
            workbook = open_workbook(ReportOutputPath, formatting_info=True)
            worksheet = workbook.sheet_by_name(KeyWordInFileName + "_AllSites")
            tmpStr = ""
            # for rowIdx in range(12,worksheet.nrows):
            print("test")
            print(str(worksheet.ncols))

            bk = copy(workbook)
            AllSites_Summary = bk.get_sheet(KeyWordInFileName + "_AllSites_Summary")

            for rowIdx in range(0, xlUsedRowsCnt):

                for columnIdx in range(0, 25):

                    if worksheet.cell_value(rowIdx, 1) != "" and columnIdx == 1:
                        tmpStr = str(worksheet.cell_value(rowIdx, 1)).strip()
                        # AllSites_Summary.write(rowIdx, 1, tmpStr)
                        if "Vmin=" in worksheet.cell_value(rowIdx, 1) or "Vmax=" in worksheet.cell_value(rowIdx, 1):
                            AllSites_Summary.write(rowIdx, 1, tmpStr[0:4])
                            tmpStr = tmpStr[5:len(tmpStr)]
                            # AllSites_Summary.write(rowIdx, 2, tmpStr[5:len(tmpStr)])
                            for j in range(len(sites) + (Activefilecnt - 1) * GenMultiFiles * len(sites)):
                                tmpStr = tmpStr + "," + str(worksheet.cell_value(rowIdx, 1 + (j + 1) * 25))[5:len(
                                    str(worksheet.cell_value(rowIdx, 1 + (j + 1) * 25)))].strip()
                                # AllSites_Summary.write(rowIdx, 1, tmpStr[0:4])
                                tmpStr = str(worksheet.cell_value(rowIdx, 1 + (j) * 25))[
                                         5:len(str(worksheet.cell_value(rowIdx, 1 + (j) * 25)))].strip()
                                if tmpStr != "":
                                    AllSites_Summary.write(rowIdx, 2 + j, (tmpStr.replace("v", "")))
                                    # AllSites_Summary.write(rowIdx, 2 + j, tmpStr)

                        else:
                            AllSites_Summary.write(rowIdx, 1, tmpStr)
                    if worksheet.cell_value(rowIdx, 2) != "" and columnIdx == 2 and not (
                            "Vmin" in worksheet.cell_value(rowIdx, 1)) and not (
                            "Vmax" in worksheet.cell_value(rowIdx, 1)):
                        tmpStr = worksheet.cell_value(rowIdx, 2)
                        AllSites_Summary.write(rowIdx, 2, tmpStr)

                    if columnIdx > 2:
                        if str(worksheet.cell_value(rowIdx, columnIdx)).strip() == "P" or str(
                                worksheet.cell_value(rowIdx, columnIdx)).strip() == "." or str(
                            worksheet.cell_value(rowIdx, columnIdx)).strip() == "*" or str(
                            worksheet.cell_value(rowIdx, columnIdx)).strip() == "":

                            tmpStr = worksheet.cell_value(rowIdx, columnIdx)

                            for k in range(len(sites) + (Activefilecnt - 1) * GenMultiFiles * len(sites)):
                                # print(worksheet.cell_value(rowIdx,columnIdx))
                                tmpStr = tmpStr + str(worksheet.cell_value(rowIdx, columnIdx + (k + 1) * 25))
                                if ("." in tmpStr and "P" in tmpStr) or ("." in tmpStr and "*" in tmpStr) or (
                                        "#" in tmpStr and "*" in tmpStr) or ("P" in tmpStr and "*" in tmpStr) or (
                                        "#" in tmpStr and "*" in tmpStr) or ("#" in tmpStr and "P" in tmpStr):

                                    pattern = Pattern()
                                    pattern.pattern = Pattern.SOLID_PATTERN
                                    pattern.pattern_fore_colour = 2
                                    style.pattern = pattern
                                    AllSites_Summary.write(rowIdx, columnIdx, tmpStr, style)

                                elif not ("Vmin" in worksheet.cell_value(rowIdx, 1) or "Vmax" in worksheet.cell_value(
                                        rowIdx, 1)):
                                    AllSites_Summary.write(rowIdx, columnIdx, tmpStr)
                                # print(tmpStr)

                        elif columnIdx != 1 or columnIdx != 2 and not (
                                "Vmin" in worksheet.cell_value(rowIdx, 1)) and not (
                                "Vmax" in worksheet.cell_value(rowIdx, 1)):
                            tmpStr = worksheet.cell_value(rowIdx, columnIdx)
                            AllSites_Summary.write(rowIdx, columnIdx, tmpStr)

            print("Done")

            if GenMultiFiles == 0:
                tmp = inputfiles[FileIdx][0:len(inputfiles[FileIdx]) - 4] + "_Site"
                for siteIdx in range(len(sites)):
                    tmp = tmp + "_" + str(sites[siteIdx])
                ReportOutPath = InputPath + "\VminVmax_" + KeyWordInFileName + "_" + CurrTime + "_" + tmp + ".xls"
                bk.save(ReportOutPath)
            elif GenMultiFiles == 1:
                ReportOutputPath = tmpPath + "\VminVmax_" + KeyWordInFileName + "_" + CurrTime + "_ALL" + ".xls"
                bk.save(ReportOutputPath)
