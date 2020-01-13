# -*- coding: utf-8 -*-
import ntpath
from os import listdir, makedirs
import os.path as osPath
from multiprocessing import Pool
import logging
import sys
from re import match
import pandas as pd
from tkinter import Tk, filedialog, messagebox
from PyQt5 import QtCore, QtGui, QtWidgets
import pyautogui as pyag
from random import randint
import time
import win32clipboard

logging.basicConfig(filename=("CrashBenchDataEntryer.log"), level=logging.DEBUG, format='%(asctime)s: %(message)s')

class CrashBenchDataEntryer(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(640, 176)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("imgTpl/starLogo24x24.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.lcdNumber = QtWidgets.QLCDNumber(self.centralwidget)
        self.lcdNumber.setEnabled(True)
        self.lcdNumber.setGeometry(QtCore.QRect(589, 129, 41, 24))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.lcdNumber.setFont(font)
        self.lcdNumber.setAutoFillBackground(True)
        self.lcdNumber.setStyleSheet("color:rgb(0, 0, 255)")
        self.lcdNumber.setDigitCount(3)
        self.lcdNumber.setSegmentStyle(QtWidgets.QLCDNumber.Flat)
        self.lcdNumber.setProperty("value", 0.0)
        self.lcdNumber.setProperty("intValue", 0)
        self.lcdNumber.setObjectName("lcdNumber")
        self.btnBrowse = QtWidgets.QPushButton(self.centralwidget)
        self.btnBrowse.setGeometry(QtCore.QRect(10, 10, 75, 22))
        self.btnBrowse.setObjectName("btnBrowse")
        self.btnBrowse.clicked.connect(self.browse4ExcelFile)
        self.lblExcelFilePath = QtWidgets.QLabel(self.centralwidget)
        self.lblExcelFilePath.setGeometry(QtCore.QRect(90, 10, 511, 22))
        self.lblExcelFilePath.clear()
        self.lblExcelFilePath.setObjectName("lblExcelFilePath")
        self.lblExcelFilePath.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.lnEdStartRow = QtWidgets.QLineEdit(self.centralwidget)
        self.lnEdStartRow.setGeometry(QtCore.QRect(115, 50, 33, 20))
        self.lnEdStartRow.setAlignment(QtCore.Qt.AlignCenter)
        self.lnEdStartRow.setObjectName("lnEdStartRow")
        self.lblStep0 = QtWidgets.QLabel(self.centralwidget)
        self.lblStep0.setGeometry(QtCore.QRect(608, 10, 22, 22))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.lblStep0.setFont(font)
        self.lblStep0.setStyleSheet("background-color:rgb(255, 0, 0);\n" "color:rgb(255, 255, 0)")
        self.lblStep0.setAlignment(QtCore.Qt.AlignCenter)
        self.lblStep0.setObjectName("lblStep0")
        self.lblStep1 = QtWidgets.QLabel(self.centralwidget)
        self.lblStep1.setGeometry(QtCore.QRect(608, 50, 22, 22))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.lblStep1.setFont(font)
        self.lblStep1.setStyleSheet("background-color:rgb(255, 0, 0);\n" "color:rgb(255, 255, 0)")
        self.lblStep1.setAlignment(QtCore.Qt.AlignCenter)
        self.lblStep1.setObjectName("lblStep1")
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(10, 130, 501, 22))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setGeometry(QtCore.QRect(0, 40, 640, 2))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.lblFromRow = QtWidgets.QLabel(self.centralwidget)
        self.lblFromRow.setGeometry(QtCore.QRect(10, 50, 105, 22))
        self.lblFromRow.setObjectName("lblFromRow")
        self.lblToRow = QtWidgets.QLabel(self.centralwidget)
        self.lblToRow.setGeometry(QtCore.QRect(152, 50, 35, 22))
        self.lblToRow.setObjectName("lblToRow")
        self.lnEdEndRow = QtWidgets.QLineEdit(self.centralwidget)
        self.lnEdEndRow.setGeometry(QtCore.QRect(188, 50, 33, 20))
        self.lnEdEndRow.setAlignment(QtCore.Qt.AlignCenter)
        self.lnEdEndRow.setObjectName("lnEdEndRow")
        self.lblWithinSheet = QtWidgets.QLabel(self.centralwidget)
        self.lblWithinSheet.setGeometry(QtCore.QRect(224, 50, 61, 22))
        self.lblWithinSheet.setObjectName("lblWithinSheet")
        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(289, 50, 231, 22))
        self.comboBox.setObjectName("comboBox")
        self.btnLoad = QtWidgets.QPushButton(self.centralwidget)
        self.btnLoad.setGeometry(QtCore.QRect(527, 50, 75, 22))
        self.btnLoad.setObjectName("btnLoad")
        self.btnLoad.clicked.connect(self.loadCrashData)
        self.line_2 = QtWidgets.QFrame(self.centralwidget)
        self.line_2.setGeometry(QtCore.QRect(0, 80, 640, 2))
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.lblVehicleType = QtWidgets.QLabel(self.centralwidget)
        self.lblVehicleType.setGeometry(QtCore.QRect(355, 90, 61, 22))
        self.lblVehicleType.setObjectName("lblVehicleType")
        self.lnEdVehicleType = QtWidgets.QLineEdit(self.centralwidget)
        self.lnEdVehicleType.setGeometry(QtCore.QRect(418, 90, 104, 20))
        self.lnEdVehicleType.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.lnEdVehicleType.setObjectName("lnEdVehicleType")
        self.lnEdSpeedUnit = QtWidgets.QLineEdit(self.centralwidget)
        self.lnEdSpeedUnit.setGeometry(QtCore.QRect(595, 90, 35, 20))
        self.lnEdSpeedUnit.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.lnEdSpeedUnit.setObjectName("lnEdSpeedUnit")
        self.lblSpeedUnit = QtWidgets.QLabel(self.centralwidget)
        self.lblSpeedUnit.setGeometry(QtCore.QRect(540, 90, 54, 22))
        self.lblSpeedUnit.setObjectName("lblSpeedUnit")
        self.lnEdManufacturer = QtWidgets.QLineEdit(self.centralwidget)
        self.lnEdManufacturer.setGeometry(QtCore.QRect(265, 90, 72, 20))
        self.lnEdManufacturer.clear()
        self.lnEdManufacturer.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.lnEdManufacturer.setObjectName("lnEdManufacturer")
        self.lblManufacturer = QtWidgets.QLabel(self.centralwidget)
        self.lblManufacturer.setGeometry(QtCore.QRect(240, 90, 21, 22))
        self.lblManufacturer.setObjectName("lblManufacturer")
        self.lnEdSamplingTime = QtWidgets.QLineEdit(self.centralwidget)
        self.lnEdSamplingTime.setGeometry(QtCore.QRect(85, 90, 31, 20))
        self.lnEdSamplingTime.clear()
        self.lnEdSamplingTime.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.lnEdSamplingTime.setObjectName("lnEdSamplingTime")
        self.lblSamplingTime = QtWidgets.QLabel(self.centralwidget)
        self.lblSamplingTime.setGeometry(QtCore.QRect(10, 90, 71, 22))
        self.lblSamplingTime.setObjectName("lblSamplingTime")
        self.lnEdUnit = QtWidgets.QLineEdit(self.centralwidget)
        self.lnEdUnit.setGeometry(QtCore.QRect(182, 90, 40, 20))
        self.lnEdUnit.clear()
        self.lnEdUnit.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.lnEdUnit.setObjectName("lnEdUnit")
        self.lblUnit = QtWidgets.QLabel(self.centralwidget)
        self.lblUnit.setGeometry(QtCore.QRect(154, 90, 31, 22))
        self.lblUnit.setObjectName("lblUnit")
        self.lblSamplingTimeUnit = QtWidgets.QLabel(self.centralwidget)
        self.lblSamplingTimeUnit.setGeometry(QtCore.QRect(120, 90, 16, 22))
        self.lblSamplingTimeUnit.setObjectName("lblSamplingTimeUnit")
        self.line_3 = QtWidgets.QFrame(self.centralwidget)
        self.line_3.setGeometry(QtCore.QRect(144, 80, 2, 40))
        self.line_3.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName("line_3")
        self.line_4 = QtWidgets.QFrame(self.centralwidget)
        self.line_4.setGeometry(QtCore.QRect(230, 80, 2, 40))
        self.line_4.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_4.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_4.setObjectName("line_4")
        self.line_5 = QtWidgets.QFrame(self.centralwidget)
        self.line_5.setGeometry(QtCore.QRect(345, 80, 2, 40))
        self.line_5.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_5.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_5.setObjectName("line_5")
        self.line_6 = QtWidgets.QFrame(self.centralwidget)
        self.line_6.setGeometry(QtCore.QRect(10, 120, 620, 2))
        self.line_6.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_6.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_6.setObjectName("line_6")
        self.line_7 = QtWidgets.QFrame(self.centralwidget)
        self.line_7.setGeometry(QtCore.QRect(530, 80, 2, 40))
        self.line_7.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_7.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_7.setObjectName("line_7")
        self.btnGo = QtWidgets.QPushButton(self.centralwidget)
        self.btnGo.setGeometry(QtCore.QRect(510, 130, 75, 22))
        self.btnGo.setText("")
        self.btnGo.clicked.connect(self.importDataToCrashBench)
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("imgTpl/playButton32x32.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btnGo.setIcon(icon1)
        self.btnGo.setObjectName("btnGo")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        self.state0()
    infoTpl4ExcelData = dict(CRASHCODE = "CRASH[\s_]*CODE",\
                                CRASHTYPE = "CRASH[\s_]*TYPE",\
                                VELOCITY = "VELOCITY|SPEED",\
                                DATALOCATION = "DATA[\s_]*LOCATION")
    constFobInfoTpl = dict(CRASHCODE = "_ne_",\
                              VEHICLETYPE = "_ne_",\
                              VELOCITY = "_ne_",\
                              SPEEDUNIT = "_ne_",\
                              DIRECTION = "_ne_",\
                              LOCATION = "_ne_",\
                              MANUFACTURER = "_ne_",\
                              CRASHTYPE = "_ne_",\
                              SIGNALTYPE = "_ne_",\
                              NUMBER_OF_SAMPLES = "_ne_",\
                              SAMPLING_TIME = "_ne_",\
                              WEIGHT = "_ne_",\
                              UNIT = "_ne_",\
                              DATA = "_ne_")
    fobInfoTpl = constFobInfoTpl.copy()
    constFobDataKeyWorkTpl = dict(DIRECTION = ["DIRECTION"],\
                                  LOCATION = ["NAME OF THE CHANNEL", "LOCATION"],\
                                  NUMBER_OF_SAMPLES = ["NUMBER OF SAMPLES"],\
                                  SAMPLING_TIME = ["SAMPLING INTERVAL"],\
                                  UNIT = ["UNIT"])
    fobDataKeyWorkTpl = constFobDataKeyWorkTpl.copy()
    def isNumeric(numStr):
        return numStr.replace('.','',1).replace('e-','',1).lstrip('-')\
               .replace('e+','',1).lstrip('+').replace('e','',1).isdigit()
    
    def uniqueList(list1): 
        # intilize a null list 
        unique_list = [] 
        # traverse for all elements 
        for x in list1: 
            # check if exists in unique_list or not 
            if x not in unique_list: 
                unique_list.append(x) 
        return unique_list
    
    def path_leaf(path):
        head, tail = ntpath.split(path)
        return tail or ntpath.basename(head)
    
    def makeFobFile(fullFilePath,fobInfoDict):
        f = open(fullFilePath, "w")
        f.write("***************************************************\n")
        f.write("* Unique Fob Dedicated For 1 Crash, Auto Generated*\n")
        f.write("***************************************************\n")
        f.write("*        10        20        30        40        50\n")
        f.write("*2345678901234567890123456789012345678901234567890\n")
        f.write("*\n")
        f.write("FILEFORMAT:\t\tsingle\n")
        f.write("DATAFORMAT:\t\tASCII\n")
        f.write("ORIENTATION:\t\tposition\n")
        f.write("PATH:\t\t\t_ne_\n")
        f.write("*\n")
        for key in fobInfoDict:
            f.write("{0}:{1}{2}\n".format(key.replace("_"," ").upper(),' '*(23-len(key)),fobInfoDict[key]))
        f.write("*\n")
        f.write("NUMBER OF SENSORS:\t\"1\"\n")
        f.write("SENSORTYPE:\t\t_ne_\n")
        f.write("DATE:\t\t\t_ne_\n")
        f.write("COMMENT 1:\t\t_ne_\n")
        f.write("COMMENT 2:\t\t_ne_\n")
        f.write("COMMENT 3:\t\t_ne_\n")
        f.write("COMMENT 4:\t\t_ne_\n")
        f.write("COMMENT 5:\t\t_ne_\n")
        f.write("*\n")
        f.write("NUMBER OF COLUMNS:\t1\n")
        f.write("SEPARATOR:\t\t_ne_\n")
        f.write("STARTCHANNEL:\t\t1\n")
        f.write("READ DIRECTION:\t\tdown\n")
        f.write("*\n")
        f.write("SPECIAL FORMAT:\t\tnormal\n")
        f.write("*\n")
        f.close()
        
    def readSingleChannelDataFile(fullFilePath):
        fileBaseName = CrashBenchDataEntryer.path_leaf(fullFilePath)
        fileBaseNameSplit = fileBaseName.split(".")
        if len(fileBaseNameSplit) == 2:
            if match("CHN|DEF|EVA|PRO|MME|LOG|FOB",fileBaseNameSplit[1].strip().upper()):
                return ""
        if match("PROTOCOL",fileBaseNameSplit[0].strip().upper()):
            return ""
        f = open(fullFilePath, "r")
        fobInfoData = CrashBenchDataEntryer.fobInfoTpl.copy()
        fobPendingInfoKeyWord = CrashBenchDataEntryer.fobDataKeyWorkTpl.copy()
        if f.mode == "r":
            line = f.readline()
            lineSplit = []
            tmpStr1 = ""
            tmpStr2 = ""
            linecount = 1
            dataRowInd = -1
            while dataRowInd == -1 and line:
                lineSplit = line.split(":")
                if len(lineSplit) == 2:
                    tmpStr1 = lineSplit[0].strip().upper()
                    tmpStr2 = lineSplit[1].strip()
                    for key in fobPendingInfoKeyWord:
                        if tmpStr1 in fobPendingInfoKeyWord[key]:
                            if tmpStr1 == "SAMPLING INTERVAL":
                                fobInfoData["SAMPLING_TIME"] = "\"" + str(float(tmpStr2)*1000) + "\""
                            elif tmpStr1 == "UNIT":
                                fobInfoData["UNIT"] = "\"" + tmpStr2 + "\""
                            else:
                                fobInfoData[key] = str(linecount) + "," + str(line.find(tmpStr2)+1) + ",50"
                            fobPendingInfoKeyWord.pop(key)
                            break
                elif CrashBenchDataEntryer.isNumeric(line.strip()):
                    dataRowInd = linecount
                else:
                    logging.info("Failed To Parse " + fullFilePath + " !!!")
                line = f.readline()
                linecount+=1
            fobInfoData["DATA"] = str(dataRowInd) + ",1,50"
        else:
            logging.info("Failed To Read " + fullFilePath + " !!!")
        f.close()
        return {fileBaseName:fobInfoData} if len(fobPendingInfoKeyWord) < len(CrashBenchDataEntryer.fobDataKeyWorkTpl) else ""
    
    def readAllChannelDataInACrashFolder(crashFolder):
        fileList = []
        try:
            fileList = [f for f in listdir(crashFolder) if osPath.isfile(osPath.join(crashFolder, f))]
        except:
            logging.info("Fatal Error With Data Location " + crashFolder)
            messagebox.showerror("Fatal Error","Encounter Error Reading Data Location:\n" + crashFolder + "\nPlease Make Sure Existence of All Data Location Within Selected Rows!!!")
            sys.exit()
        pool=Pool()
        results = pool.map(CrashBenchDataEntryer.readSingleChannelDataFile, [osPath.join(crashFolder, fileName) for fileName in fileList])
        pool.terminate()
        dataDict = {}
        for item in results:
           dataDict.update(item)
        uniqueFobInfoDataList = CrashBenchDataEntryer.uniqueList(list(dataDict.values()))
        classifiedChannelGroupList = [] 
        for i in range(len(uniqueFobInfoDataList)):
            if (i+1) > len(classifiedChannelGroupList):
                classifiedChannelGroupList.append([])
            for key in dataDict:
                for j in range(i):
                    if key in classifiedChannelGroupList[j]:
                        break
                if dataDict[key] == uniqueFobInfoDataList[i]:
                    classifiedChannelGroupList[i].append(key)
        if not classifiedChannelGroupList or not uniqueFobInfoDataList:
            logging.info("Fatal Error With Data Files In:\n" + crashFolder)
            messagebox.showerror("Fatal Error","Encounter Error Reading Data Files In:\n" + crashFolder + "\nPlease Make Sure Qualified Data Files Within Selected Rows!!!")
            sys.exit()
        return classifiedChannelGroupList,uniqueFobInfoDataList
    
    def makeBatchFobFilesInACrashFolder(crashInfoInput):
        chList,fobInfoDataList = CrashBenchDataEntryer.readAllChannelDataInACrashFolder(crashInfoInput[0])
        tmpFobFileName = ""
        fobFileFolderPath = osPath.join(crashInfoInput[0],"BatchFobFiles")
        makedirs(fobFileFolderPath,exist_ok = True)
        f = open(osPath.join(fobFileFolderPath,"FobFileToDataFileSelectionMapping.config"), "w")
        fobFileDict = {}
        tmpStr1 = ""
        tmpStr2 = ""
        for i in range(len(chList)):
            tmpFobFileName = "Part" + str(i) + ".fob"
            tmpStr1 = osPath.join(fobFileFolderPath,tmpFobFileName)
            tmpStr2 = "\"" + "\"\"".join(chList[i]) + "\""
            f.write(tmpFobFileName + "\t" + tmpStr2 + "\n")
            fobInfoDataList[i]["CRASHCODE"] = "\"" + crashInfoInput[1] + "\""
            fobInfoDataList[i]["CRASHTYPE"] = "\"" + crashInfoInput[2] + "\""
            fobInfoDataList[i]["VELOCITY"] = "\"" + crashInfoInput[3] + "\""
            CrashBenchDataEntryer.makeFobFile(tmpStr1,fobInfoDataList[i])
            fobFileDict[tmpStr1] = tmpStr2
        f.close()
        return fobFileDict
        
    def setClipboardData(txt):
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(txt)
        win32clipboard.CloseClipboard()
    
    def getClipboardData():
        win32clipboard.OpenClipboard()
        data = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        return data
    
    def splitChannelStr(chStr):
        startInd = 0
        step = 256
        endInd = startInd + step
        resultList = []
        chStrLen = len(chStr)
        while endInd < chStrLen:
            endInd = chStr.rfind("\"",startInd,endInd) + 1
            resultList.append("\"" + chStr[startInd:endInd])
            startInd = endInd
            endInd = startInd + step
        resultList.append("\"" + chStr[startInd:] + "\"")
        return resultList
    
    def locateOnScreen(tplImageFilePath,confidence):
        try:
            x, y, w, h = pyag.locateOnScreen(tplImageFilePath,confidence)
            return True, x, y, w, h
        except:
            return False,
    
    def locateCenterOnScreen(tplImageFilePath,confidence):
        try:
            loc = pyag.locateOnScreen(tplImageFilePath,confidence)
            x, y = pyag.center(loc)
            return True, x, y
        except:
            return False,
    
    def importDataElementToCrashBenchNG332(fullFobFilePath,crashChSelList):
        tmpLoc = CrashBenchDataEntryer.locateCenterOnScreen("imgTpl/btnLoadFob.png",confidence = 1)
        while not tmpLoc[0]:
            time.sleep(0.1)
            tmpLoc = CrashBenchDataEntryer.locateCenterOnScreen("imgTpl/btnLoadFob.png",confidence = 1)
        pyag.click(tmpLoc[1:3]) 
        CrashBenchDataEntryer.setClipboardData(fullFobFilePath)
        time.sleep(0.1)
        tmpLoc = CrashBenchDataEntryer.locateCenterOnScreen("imgTpl/dialogFobFileOpen.png",confidence = 1)
        while not tmpLoc[0]:
            pyag.click()
            time.sleep(0.1)
            tmpLoc = CrashBenchDataEntryer.locateCenterOnScreen("imgTpl/dialogFobFileOpen.png",confidence = 1)
        pyag.click(tmpLoc[1] + 70,tmpLoc[2] - 28)
        pyag.hotkey('ctrl', 'v')
        while CrashBenchDataEntryer.locateCenterOnScreen("imgTpl/stsFobFilePathNotEnterred.png",confidence = 1)[0]:
            time.sleep(0.1)
        pyag.press('enter')
        while CrashBenchDataEntryer.locateCenterOnScreen("imgTpl/dialogFobFileOpen.png",confidence = 1)[0]:
            time.sleep(0.1)
        pyag.press('f4')
        CrashBenchDataEntryer.setClipboardData(osPath.join(osPath.dirname(fullFobFilePath),"..\\"))
        time.sleep(0.1)
        tmpLoc = CrashBenchDataEntryer.locateCenterOnScreen("imgTpl/dialogDataFileOpen.png",confidence = 1)
        while not tmpLoc[0]:
            time.sleep(0.1)
            tmpLoc = CrashBenchDataEntryer.locateCenterOnScreen("imgTpl/dialogDataFileOpen.png",confidence = 1)
        pyag.click(tmpLoc[1] + 70,tmpLoc[2] - 28)
        pyag.hotkey('ctrl', 'v')
        while CrashBenchDataEntryer.locateCenterOnScreen("imgTpl/stsDataFilePathNotEnterred.png",confidence = 1)[0]:
            time.sleep(0.1)
        pyag.press('enter')
        for crashChSelStrInd in range(len(crashChSelList)):
            if crashChSelStrInd:
                pyag.press('f4')   
                tmpLoc = CrashBenchDataEntryer.locateCenterOnScreen("imgTpl/dialogDataFileOpen.png",confidence = 1)            
                while not tmpLoc[0]:
                    time.sleep(0.1)
                    tmpLoc = CrashBenchDataEntryer.locateCenterOnScreen("imgTpl/dialogDataFileOpen.png",confidence = 1)            
                pyag.click(tmpLoc[1] + 70,tmpLoc[2] - 28)
            CrashBenchDataEntryer.setClipboardData(crashChSelList[crashChSelStrInd])        
            time.sleep(0.1)
            pyag.hotkey('ctrl', 'v')
            while CrashBenchDataEntryer.locateCenterOnScreen("imgTpl/stsDataFilePathNotEnterred.png",confidence = 1)[0]:
                time.sleep(0.1)
            pyag.press('enter')
            for i in range(3):
                time.sleep(0.7)
                tmpLoc = CrashBenchDataEntryer.locateCenterOnScreen("imgTpl/stsDataProcessing.png",confidence = 1)[0]
                if not tmpLoc:
                    break
            if tmpLoc:
                while CrashBenchDataEntryer.locateCenterOnScreen("imgTpl/stsDataReading.png",confidence = 1)[0]:
                    time.sleep(0.2)
                while CrashBenchDataEntryer.locateCenterOnScreen("imgTpl/stsDataReadFailed.png",confidence = 1)[0]:
                    logging.info("Failed To Read Data With Fob " + fullFobFilePath)
                    pyag.press('enter')
                    time.sleep(0.2)
        return True
        
    def importSingleCrashDataToCrashBenchNG332(fobFileDict):
        for key in fobFileDict:
            CrashBenchDataEntryer.importDataElementToCrashBenchNG332(key,CrashBenchDataEntryer.splitChannelStr(fobFileDict[key]))
    
    def browse4ExcelFile(self):
        self.state0()
        excel_file_path = filedialog.askopenfilename(title="Select Channel Overview Excel File",
                                                     filetypes=[("Excel Workbook", ("*.xls","*.xlsx"))])
        fileExtension = excel_file_path.split(".")[-1]
        if fileExtension == "xlsx" or fileExtension == "xls":
            self.lblExcelFilePath.setText(excel_file_path)
            self.excelFile = pd.ExcelFile(excel_file_path)
            self.comboBox.addItems(self.excelFile.sheet_names)
            self.state1()
        else:
            messagebox.showerror("Invalid Input", "Excel File Type Is Expected !!!")
    def loadCrashData(self):
        self.state1()
        self.excelFile = pd.ExcelFile(self.lblExcelFilePath.text())   
        CrashBenchDataEntryer.fobDataKeyWorkTpl = CrashBenchDataEntryer.constFobDataKeyWorkTpl.copy()
        CrashBenchDataEntryer.fobInfoTpl = CrashBenchDataEntryer.constFobInfoTpl.copy()
        tmpSelectedTxt = self.comboBox.currentText()
        self.comboBox.clear()
        self.comboBox.addItems(self.excelFile.sheet_names)
        self.comboBox.setCurrentText(tmpSelectedTxt)
        if self.lnEdStartRow.text().isdigit() and self.lnEdEndRow.text().isdigit():
            tmpInt1 = int(self.lnEdStartRow.text())
            tmpInt2 = int(self.lnEdEndRow.text())
            startRow = (tmpInt1 if tmpInt1 < tmpInt2 else tmpInt2) - 2
            endRow = (tmpInt2 if tmpInt1 < tmpInt2 else tmpInt1) - 2
            activeSheet = self.excelFile.parse(sheet_name = self.comboBox.currentText(),header=0,skiprows=0,skip_blank_lines=False)
            infoColName = {}
            pendingInfoColName = CrashBenchDataEntryer.infoTpl4ExcelData.copy()
            headerRow = 0
            for row in range(activeSheet.shape[0]):
                if headerRow:
                    if pendingInfoColName == {}:
                        break
                    else:                   
                        pendingInfoColName = CrashBenchDataEntryer.infoTpl4ExcelData.copy()
                        headerRow = 0
                for col in range(activeSheet.shape[1]):
                    orgCellStr = str(activeSheet.iloc[row,col])
                    tmpCellStr = orgCellStr.strip().upper()
                    for key,value in list(pendingInfoColName.items()):
                        if match(value,tmpCellStr):                
                            if key == "CRASHCODE":
                                headerRow = row
                            infoColName[key] = orgCellStr
                            pendingInfoColName.pop(key,None)
                            break                    
            if headerRow and not pendingInfoColName:
                activeSheet.columns = activeSheet.iloc[headerRow]
                activeSheet = activeSheet.loc[startRow:endRow,list(infoColName.values())]
                activeSheet.rename(columns=dict((v,k) for k,v in infoColName.items()),inplace=True)
                #activeSheet.reset_index(drop=True, inplace=True)
                pendingInfoColName = {}
                for infoTag in infoColName:
                    tmpNACol = activeSheet[infoTag][activeSheet[infoTag].isna()]
                    if not tmpNACol.empty:
                        pendingInfoColName[infoTag] = list(tmpNACol.index+2)     
                if pendingInfoColName:
                    tmpMsg = ""
                    for k,v in pendingInfoColName.items():
                        tmpMsg += "\nColumn <" + infoColName[k] + "> at row " + ', '.join(map(str,v))
                    messagebox.showerror("Missing Required Information", "Please check and fill in necessary information as follows," + tmpMsg)
                else:
                    activeSheet['VELOCITY'] = activeSheet['VELOCITY'].astype(str).str.extract(r'(\d+\.?\d*)\D*(\d+\.?\d*)?').astype('float64').max(axis=1,skipna=True).apply(lambda x: str(int(x)) if x == int(x) else str(x))
                    activeSheet['DATALOCATION'] = activeSheet['DATALOCATION'].str.replace('/', '\\', regex=False)
                    activeSheet = activeSheet[['DATALOCATION','CRASHCODE','CRASHTYPE','VELOCITY']]
                    self.CrashInfoList = activeSheet.values.tolist()
                    randomCrashInfoIndex = randint(0,len(self.CrashInfoList)-1)
                    refData = CrashBenchDataEntryer.readAllChannelDataInACrashFolder(self.CrashInfoList[randomCrashInfoIndex][0])[1]
                    if not refData:
                        logging.info("Fatal Error With Data Files In:\n" + self.CrashInfoList[randomCrashInfoIndex][0])
                        messagebox.showerror("Fatal Error","Encounter Error Reading Data Files In:\n" + self.CrashInfoList[randomCrashInfoIndex][0] + "\nPlease Make Sure Qualified Data Files Within Selected Rows!!!")
                        sys.exit()
                    refData = refData[randint(0,len(refData)-1)]
                    self.lnEdSamplingTime.setText(refData["SAMPLING_TIME"].replace("\"",""))
                    self.lnEdUnit.setText(refData["UNIT"].replace("\"",""))
                    self.lnEdManufacturer.setText(refData["MANUFACTURER"].replace("\"",""))
                    self.lnEdVehicleType.setText(refData["VEHICLETYPE"].replace("\"",""))
                    self.lnEdSpeedUnit.setText(refData["SPEEDUNIT"].replace("\"",""))
                    self.state2()
            else:
                messagebox.showerror("Missing Required Column", "Missing Data For " + ", ".join(list([x for x in pendingInfoColName if not x in infoColName])))
        else:
            messagebox.showerror("Invalid Input", "Please Enter Positive Integer !!!")
            
    def adaptInfoFromLnEd(infoTxt):
        tmpInpStr = infoTxt.strip().replace("\"","").lower()
        if len(tmpInpStr) == 0 or match("\_*n[ea]\_*",tmpInpStr):
            return "_ne_"
        elif match("\d+\,\d+\,\d+",tmpInpStr):
            return tmpInpStr
        else:
            return "\"" + tmpInpStr + "\""
    def importDataToCrashBench(self):
        self.state3()
        CrashBenchDataEntryer.fobDataKeyWorkTpl.pop("SAMPLING_TIME", None)
        CrashBenchDataEntryer.fobDataKeyWorkTpl.pop("UNIT", None)
        CrashBenchDataEntryer.fobInfoTpl["SAMPLING_TIME"] = CrashBenchDataEntryer.adaptInfoFromLnEd(self.lnEdSamplingTime.text())
        CrashBenchDataEntryer.fobInfoTpl["UNIT"] = CrashBenchDataEntryer.adaptInfoFromLnEd(self.lnEdUnit.text())
        CrashBenchDataEntryer.fobInfoTpl["MANUFACTURER"] = CrashBenchDataEntryer.adaptInfoFromLnEd(self.lnEdManufacturer.text())
        CrashBenchDataEntryer.fobInfoTpl["VEHICLETYPE"] = CrashBenchDataEntryer.adaptInfoFromLnEd(self.lnEdVehicleType.text())
        CrashBenchDataEntryer.fobInfoTpl["SPEEDUNIT"] = CrashBenchDataEntryer.adaptInfoFromLnEd(self.lnEdSpeedUnit.text())
        crashCount = len(self.CrashInfoList)
        self.progressBar.setRange(0,crashCount)
        for crashInd in range(crashCount):            
            CrashBenchDataEntryer.importSingleCrashDataToCrashBenchNG332(CrashBenchDataEntryer.makeBatchFobFilesInACrashFolder(self.CrashInfoList[crashInd]))
            self.progressBar.setValue(crashInd+1)
            self.lcdNumber.display(crashInd+1)
            QtWidgets.QApplication.processEvents()
        messagebox.showinfo("Progress Completion","All Data Is Successfully Imported Into CrashBench !")
        self.state4()
        
    def state0(self):
        self.state = 0
        self.lblExcelFilePath.clear()
        self.comboBox.clear()
        self.lnEdStartRow.clear()
        self.lnEdEndRow.clear()
        self.lnEdSamplingTime.clear()
        self.lnEdUnit.setText("g")
        self.lnEdManufacturer.clear()
        self.lnEdVehicleType.clear()
        self.lnEdSpeedUnit.setText("km/h")      
        self.progressBar.reset()
        self.lcdNumber.display(0)
        self.lblStep0.setText("-")
        self.lblStep0.setStyleSheet("background-color:rgb(255, 0, 0);\n" "color:rgb(255, 255, 0)")
        self.lblStep1.setText("-")
        self.lblStep1.setStyleSheet("background-color:rgb(255, 0, 0);\n" "color:rgb(255, 255, 0)")
        self.btnLoad.setEnabled(False)
        self.btnGo.setEnabled(False)
    def state1(self):
        self.state = 1
        self.lnEdSamplingTime.clear()
        self.lnEdUnit.setText("g")
        self.lnEdManufacturer.clear()
        self.lnEdVehicleType.clear()
        self.lnEdSpeedUnit.setText("km/h")
        self.progressBar.reset()
        self.lcdNumber.display(0)
        self.lblStep0.setText("ok")
        self.lblStep0.setStyleSheet("background-color:rgb(0, 255, 0);\n" "color:rgb(0, 0, 255)")
        self.lblStep1.setText("-")
        self.lblStep1.setStyleSheet("background-color:rgb(255, 0, 0);\n" "color:rgb(255, 255, 0)")
        self.btnLoad.setEnabled(True)
        self.btnGo.setEnabled(False)

    def state2(self):
        self.state = 2
        self.lblStep1.setText("ok")
        self.lblStep1.setStyleSheet("background-color:rgb(0, 255, 0);\n" "color:rgb(0, 0, 255)")
        self.btnGo.setEnabled(True)

    def state3(self):
        self.lnEdStartRow.setEnabled(False)
        self.lnEdEndRow.setEnabled(False)
        self.comboBox.setEnabled(False)
        self.lnEdSamplingTime.setEnabled(False)
        self.lnEdUnit.setEnabled(False)
        self.lnEdManufacturer.setEnabled(False)
        self.lnEdVehicleType.setEnabled(False)
        self.lnEdSpeedUnit.setEnabled(False)
        self.btnBrowse.setEnabled(False)
        self.btnLoad.setEnabled(False)
        self.btnGo.setEnabled(False)
        
    def state4(self):
        self.lnEdStartRow.setEnabled(True)
        self.lnEdEndRow.setEnabled(True)
        self.comboBox.setEnabled(True)
        self.lnEdSamplingTime.setEnabled(True)
        self.lnEdUnit.setEnabled(True)
        self.lnEdManufacturer.setEnabled(True)
        self.lnEdVehicleType.setEnabled(True)
        self.lnEdSpeedUnit.setEnabled(True)
        self.btnBrowse.setEnabled(True)
        self.btnLoad.setEnabled(True)
        self.btnGo.setEnabled(True)
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "CrashBench Data Entryer"))
        self.btnBrowse.setText(_translate("MainWindow", "Browse"))
        self.lblStep0.setText(_translate("MainWindow", "-"))
        self.lblStep1.setText(_translate("MainWindow", "-"))
        self.lblFromRow.setText(_translate("MainWindow", "Select data from row"))
        self.lblToRow.setText(_translate("MainWindow", "to row"))
        self.lblWithinSheet.setText(_translate("MainWindow", "within sheet"))
        self.btnLoad.setText(_translate("MainWindow", "Load"))
        self.lblVehicleType.setText(_translate("MainWindow", "VehicleType"))
        self.lnEdSpeedUnit.setText(_translate("MainWindow", "km/h"))
        self.lblSpeedUnit.setText(_translate("MainWindow", "SpeedUnit"))
        self.lblManufacturer.setText(_translate("MainWindow", "Mfr."))
        self.lblSamplingTime.setText(_translate("MainWindow", "SamplingTime"))
        self.lblUnit.setText(_translate("MainWindow", "Unit"))
        self.lblSamplingTimeUnit.setText(_translate("MainWindow", "ms"))

def main():
    tmpTopLevelWin = Tk()
    tmpTopLevelWin.attributes("-topmost",True)
    tmpTopLevelWin.withdraw()
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    myprog = CrashBenchDataEntryer()
    myprog.setupUi(MainWindow)    
    MainWindow.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
