# PS-Cal Corrector
# By Micah Hurd
programName = "PS-Cal Corrector"
version = 2

# Dependent on Installation of Excel Wings (see data to Excel function)
# Use "pip install xlwings"
import re
import math
# import scipy.interpolate # Non-native
import os
import shutil
import os.path
from os import path
from pathlib import Path
from distutils.dir_util import copy_tree

from tkinter import *
from tkinter import filedialog
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import time
import tempfile

def readConfigFile(filename, searchTag, sFunc=""):
    searchTag = searchTag.lower()
    # print("Search Tag: ",searchTag)

    # Open the file
    with open(filename, "r") as filestream:
        # Loop through each line in the file

        for line in filestream:

            if line[0] != "#":

                currentLine = line
                equalIndex = currentLine.find('=')
                if equalIndex != -1:

                    tempLength = len(currentLine)
                    # print("{} {}".format(equalIndex,tempLength))
                    tempIndex = equalIndex
                    configTag = currentLine[0:(equalIndex)]
                    configTag = configTag.lower()
                    configTag = configTag.strip()
                    # print(configTag)

                    configField = currentLine[(equalIndex + 1):]
                    configField = configField.strip()
                    # print(configField)

                    # print("{} {}".format(configTag,searchTag))
                    if configTag == searchTag:

                        # Split each line into separated elements based upon comma delimiter
                        # configField = configField.split(",")

                        # Remove the newline symbol from the list, if present
                        lineLength = len(configField)
                        lastElement = lineLength - 1
                        if configField[lastElement] == "\n":
                            configField.remove("\n")
                        # Remove the final comma in the list, if present
                        lineLength = len(configField)
                        lastElement = lineLength - 1

                        if configField[lastElement] == ",":
                            configField = configField[0:lastElement]

                        lineLength = len(configField)
                        lastElement = lineLength - 1

                        # Apply string manipulation functions, if requested (optional argument)
                        if sFunc != "":
                            sFunc = sFunc.lower()

                            if sFunc == "listout":
                                configField = configField.split(",")

                            if sFunc == "stringout":
                                configField = configField.strip("\"")

                            if sFunc == "int":
                                configField = int(configField)

                            if sFunc == "float":
                                configField = float(configField)

                        filestream.close()
                        return configField

        filestream.close()
        return "Searched term could not be found"

def create_log(log_file):
    f = open(log_file, "w+")

    f.close()
    return 0

def writeLog(analysis, logFile):
    import datetime
    import csv
    write_mode = "a"

    currentDT = datetime.datetime.now()

    date_time = currentDT.strftime("%Y-%m-%d %H:%M:%S")

    with open(logFile, mode=write_mode, newline='') as result_file:
        result_writer = csv.writer(result_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)

        result_writer.writerow([date_time, analysis])

    result_file.close()

    return 0

def userInterfaceHeader(program, version, cwd, logFile, msg=""):
    print(program + ", Version " + str(version))
    print("Current Working Directory: " + str(cwd))
    print("Log file located at working directory: " + str(logFile))
    print("=======================================================================")
    if msg != "":
        print(msg)
        print("_______________________________________________________________________")
    return 0

def clear():  # Clears the console
    # for windows
    from os import system, name
    if name == 'nt':
        _ = system('cls')

        # for mac and linux(here, os.name is 'posix')
    else:
        _ = system('clear')

def yesNoGUI(questionStr, windowName=""):
    root = Tk()

    canvas1 = tk.Canvas(root, width=1, height=1)
    canvas1.pack()

    MsgBox = tk.messagebox.askquestion(windowName, questionStr, icon='warning')
    if MsgBox == 'yes':
        # root.destroy()
        # print(1)
        response = True
    else:
        # tk.messagebox.showinfo('Return', 'You will now return to the application screen')
        response = False

    # ExitApplication()
    root.destroy()
    return response

def getFilePath(extensionType, initialDir="", extensionDescription="", multi=False):
    root = Tk()

    canvas1 = tk.Canvas(root, width=1, height=1)
    canvas1.pack()

    if multi == True:
        root.filenames = filedialog.askopenfilenames(initialdir=initialDir, title="Select file",
                                                     filetypes=(
                                                     (extensionDescription, extensionType), ("all files", "*.*")))
        list = root.filenames
    else:
        root.filename = filedialog.askopenfilename(initialdir=initialDir, title="Select file",
                                                   filetypes=(
                                                   (extensionDescription, extensionType), ("all files", "*.*")))
        list = root.filename
    root.destroy()
    return list

def getDirectoryPath(initialDir=""):
    root = Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory(initialdir=initialDir)
    root.destroy()
    return folder_selected

def popupMsg(msg, popTitle=""):
    popup = tk.Tk()
    popup.wm_title(popTitle)
    label = ttk.Label(popup, text=msg)
    label.pack(side="top", fill="x", pady=10)
    B1 = ttk.Button(popup, text="Okay", command=popup.destroy)
    B1.pack()
    popup.mainloop()

def getTextEntry(buttonText="", labelText="", titleText=""):
    root = Tk()
    root.title(titleText)

    mystring = StringVar()

    def getvalue():
        global output

        output = mystring.get()

        root.destroy()

    Label(root, text=labelText).grid(row=0, sticky=W)  # label
    Entry(root, textvariable=mystring).grid(row=0, column=1, sticky=E)  # entry textbox

    WSignUp = Button(root, text=buttonText, command=getvalue).grid(row=3, column=0, sticky=W)  # button

    root.mainloop()
    return output

def yesNoPrompt(message, titleText=""):
    root = tk.Tk()  # create window

    canvas1 = tk.Canvas(root, width=0, height=0)
    canvas1.pack()

    MsgBox = tk.messagebox.askquestion(titleText, message,
                                       icon='warning')
    if MsgBox == 'yes':
        response = True
        root.destroy()
    else:
        response = False
        root.destroy()

    root.mainloop()
    return response

def readTxtFile(filename):
    # Place contents of XML files into variable
    f = open(filename, 'r')
    x = f.readlines()
    f.close()
    return x

def interpolation(yVal,x,y):

    # print(x)
    # print(y)
    # print(yVal)
    y_interp = scipy.interpolate.interp1d(x, y)
    return(y_interp(yVal))

def performInterpolation():

    for index, freq in enumerate(rhoFreqList):

        cfFreqListNewFreq = cfFreqListNew[index]
        if cfFreqListNewFreq == 0:

            cfFreqListNew[index] = freq

            x = rhoFreqList.copy()

            newCFVal = float(interpolation(freq, xVals, yValsCF))
            newUncVal = interpolation(freq, xVals, yValsUNC)
            newDbVal = float(interpolation(freq, xVals, yValsDb))

            # Take the interpolated uncertainty value and RSS double it
            unc1 = newUncVal ** 2
            unc2 = newUncVal ** 2
            unc3 = unc1+unc2
            newUncVal = math.sqrt(unc3)

            cfListNew[index] = newCFVal
            uncListNew[index] = newUncVal
            dbListNew[index] = newDbVal

def insertCalFactor(index,newCFBlock):
    # print(newCFBlock)
    # print("IndexLoc {}".format(xmlDataNew[index]))
    #
    #
    # print(newCFBlock)
    tempLength = len(newCFBlock)
    tempCounter = tempLength
    while tempCounter >= 1:
        tempCounter -= 1
        data = newCFBlock[tempCounter]

        xmlDataNew.insert(index, data)

def editCFblock(cfBlockList,frequency,calFactor,uncertainty,dB):

    newCFBlock = cfBlockList.copy()

    for index, element in enumerate(newCFBlock):

        filterList1 = ['<OnLabel>', '<DUT_Power_Avg>', '<DeviationError>', '<RFOnStdDev>', '<DUT_Power_1>',
                       '<MisMatchFactor>']
        filterList2 = ['</OnLabel>', '</DUT_Power_Avg>', '</DeviationError>', '</RFOnStdDev>',
                       '</DUT_Power_1>', '</MisMatchFactor>']

        if ("<Frequency>" in element) and ("</Frequency>" in element):
            # print(element)
            line = element.split(">")
            element1 = line[0] + ">"
            line = element.split("<")
            element2 = "<" + line[2]
            # print(element1)
            # print(element2)

            if frequency > 1:
                frequency = "{:,.0f}".format(frequency)
            else:
                frequency = "{:,.6f}".format(frequency)

            line = "{}{}{}".format(element1,frequency,element2)
            # print(line)
            newCFBlock[index] = line

        elif ("<CalFactor>" in element) and ("</CalFactor>" in element):
            # print(element)
            line = element.split(">")
            element1 = line[0] + ">"
            line = element.split("<")
            element2 = "<" + line[2]
            # print(element1)
            # print(element2)

            calFactor = "{:.4f}".format(calFactor)

            line = "{}{}{}".format(element1,calFactor,element2)
            # print(line)
            newCFBlock[index] = line

        elif ("<Uncertainty>" in element) and ("</Uncertainty>" in element):
            # print(element)
            line = element.split(">")
            element1 = line[0] + ">"
            line = element.split("<")
            element2 = "<" + line[2]
            # print(element1)
            # print(element2)

            uncertainty = "{:.4f}".format(uncertainty)

            line = "{}{}{}".format(element1,uncertainty,element2)
            # print(line)
            newCFBlock[index] = line

        elif ("<dB>" in element) and ("</dB>" in element):
            # print(element)
            line = element.split(">")
            element1 = line[0] + ">"
            line = element.split("<")
            element2 = "<" + line[2]
            # print(element1)
            # print(element2)

            # print("dB: {}".format(dB))
            dB = "{:.4f}".format(dB)
            # print("dB: {}".format(dB))

            line = "{}{}{}".format(element1,dB,element2)
            # print(line)
            newCFBlock[index] = line

        else:
            for index2, i in enumerate(filterList1):

                filter1 = i
                filter2 = filterList2[index2]

                if (filter1 in element) and (filter2 in element):
                    # print(element)
                    line = element.split(">")
                    element1 = line[0] + ">"
                    line = element.split("<")
                    element2 = "<" + line[2]
                    # print(element1)
                    # print(element2)

                    insert = "- -"

                    line = "{}{}{}".format(element1, insert, element2)
                    # print(line)
                    # print("--------------------------------------------")
                    # print("filter1 {}".format(filter1))
                    # print("filter2 {}".format(filter2))
                    # print("element {}".format(element))
                    # print("line {}".format(line))
                    newCFBlock[index] = line
                    # print("newCFBlock {}".format(newCFBlock))
                    # input("Press Any Key To Continue...")



    return newCFBlock

def extractValueFromXML(firstWrapper, secondWrapper, lineData):

    lengthLineData = len(lineData)
    lengthFirstWrapper = len(firstWrapper)
    lengthSecondWrapper = len(secondWrapper) + 1

    # Strip the second wrapper off of the line of text
    firstSlice = lineData[:-lengthSecondWrapper]
    # print("FirstSlice: {}".format(firstSlice))
    firstSliceLength = len(firstSlice)

    # Find the string length to the end of the first wrapper
    counter = firstSliceLength - 1
    while counter > 0:

        if firstSlice[counter] == ">":
            firstWrapperEndIndex = counter + 1
            break
        counter -= 1

    # Obtain the measurement value from the XML String
    value = firstSlice[firstWrapperEndIndex:]
    valueLength = len(value)

    # Obtain the XML string up to the line
    stringWithoutValue = firstSlice[:-valueLength]

    # Create a XML string line which is prepped for having a new value inserted
    outputXMLstring = stringWithoutValue + "val" + secondWrapper

    return (value, outputXMLstring)

def setSigDigits(value,qtySigDigitsRequired):

    # Convert value to a string, if it is not one already
    value = str(value)

    filterVals = ["1","2","3","4","5","6","7","8","9"]
    firstSDFlag = False
    qtySD = 0
    stringLength = 0
    for index,character in enumerate(value):
        if (character in filterVals) and (firstSDFlag == False):
            qtySD += 1
            firstSDFlag = True
        elif (firstSDFlag == True) and (character != "."):
            qtySD += 1

        if qtySD > qtySigDigitsRequired:
            stringLength = index + 1                # Load an additional digit at the end, for rounding
            # print("Index: {}".format(index))
            break

    # If the value already meets the significant digits requirements, then return the same value.
    if stringLength == 0:
        # Convert value from string to float
        value = float(value)

        # Check to see if the resulting number is an integer; if yes then return it as one
        checkValue = int(value)
        difference = value - checkValue
        if difference == 0:
            value = int(value)
        else:
            value = float(value)

        return value

    newValue = value[:stringLength]

    # Find out how many places past the decimal point the value is
    decimalFlag = False
    qtyCount = 0
    filterVals = ["."]
    for index, character in enumerate(newValue):
        #print(character)
        if (character in filterVals) and (decimalFlag == False):
            decimalFlag = True
        elif decimalFlag == True:
            qtyCount += 1

    # print("qtyCounter: {}".format(qtyCount))


    newValue = float(newValue)
    #print(newValue)
    newValue = round(newValue, (qtyCount - 1))

    # Check to see if the resulting number is an integer; if yes then return it as one
    checkValue = int(newValue)
    difference = newValue - checkValue
    if difference == 0:
        newValue = int(newValue)

    return newValue

def checkUncBudget(budgetTxtFile,uncVal,uncFreq,uncMsmt=0.0):
    import datetime
    from datetime import date

    # Suppress scientific notation (this code ended up being unnecessary)
    # uncVal = f'{uncVal:.20f}'                       # 20 digits, because why not?
    # uncFreq = f'{uncFreq:.0f}'                      # Frequency in Hz should have no resolution after the decimal

    # Place contents of XML files into variable
    f = open(budgetTxtFile, 'r')
    budgetData = f.readlines()
    f.close()

    # Strip any existing whitespace out of the list elements
    for index, element in enumerate(budgetData):
        if element == "\n":
            del budgetData[index]
        else:
            newElement = element.strip()
            budgetData[index] = newElement



    # Placed current date into variable
    today = date.today()
    currentDate = today.strftime("%Y-%m-%d")

    # Obtain the line containing the expiration date of the budget file
    tempList = budgetData[0]
    tempList = tempList.strip()
    tempList = tempList.split(",")
    fileDate = tempList[1]
    fileDate = fileDate.split("-")                  # Store date elements into a list
    currentDate = currentDate.split("-")            # Also store today's date elements into list

    # Convert dates into formats which can be compared against
    fileDateInFormat = datetime.datetime(int(fileDate[0]), int(fileDate[1]), int(fileDate[2]))
    currentDateInFormat = datetime.datetime(int(currentDate[0]), int(currentDate[1]), int(currentDate[2]))

    # Check to see if the file date has been exceeded
    if currentDateInFormat > fileDateInFormat:
        return "File > {} < expiration date is exceeded!".format(budgetTxtFile)

    del budgetData[0]                            # Delete the list element containing the date

    # Check to see if this budget format includes Power ranges
    tempList = budgetData[0].split(",")
    tempQty = len(tempList)
    if tempQty > 2:
        pRangePresent = True
    else:
        pRangePresent = False

    # Break out the elements of the uncertainty budget list
    freqList = []
    rangeList = []
    uncList = []
    for index, element in enumerate(budgetData):
        tempList = element.split(",")
        if pRangePresent == True:
            freqList.append(tempList[0])
            rangeList.append(tempList[1])
            uncList.append(float(tempList[2]))
        else:
            freqList.append(tempList[0])
            uncList.append(float(tempList[1]))



    # print(freqList)
    # print(rangeList)
    # print(uncList)

    # Find the current uncertainty frequency within the budget frequency list
    for index, element in enumerate(freqList):
        tempList = element.split(">")
        startFreq = float(tempList[0])
        stopFreq = float(tempList[1])
        uncFreq = float(uncFreq)


        if (uncFreq >= startFreq) and (uncFreq <= stopFreq):

            if pRangePresent == True:
                tempList = rangeList[index].split(">")
                startPow = float(tempList[0])
                stopPow = float(tempList[1])

                if (uncMsmt >= startPow) and (uncMsmt <= stopPow):
                    tempUncListVal = uncList[index]

                    if uncVal < tempUncListVal:
                        uncVal = tempUncListVal
                        return uncVal
            else:
                tempUncListVal = uncList[index]
                if uncVal < tempUncListVal:
                    uncVal = tempUncListVal
                    return uncVal


    return uncVal

def imporStdList(filename):
    standardList = []
    for i in range(1, 21):
        searchTag = "standard{}".format(i)
        result = readConfigFile(filename, searchTag, sFunc="")
        if result != "Searched term could not be found":
            standardList.append(result)
    return standardList

def findAndExtractValueFromXML(header, wrapper, xmlDataList, returnLine=False):

    header = header.lower()
    wrapper = wrapper.lower()

    headerFound = False

    searchHeaderStart = header
    searchHeaderEnd = "</" + header + ">"

    searchWrapperStart = "<" + wrapper
    searchWrapperEnd = "</" + wrapper + ">"

    # Go through the XML data to find the required value
    for index, line in enumerate(xmlDataList):

        currentLine = line.lower()
        #print(currentLine)
        #input("Press enter to continue: ")

        # If the header is found, then mark flag as true until the end of the header occurs
        if searchHeaderStart in currentLine:
            headerFound = True
            # print("HeaderFound: {}".format(headerFound))
        elif searchHeaderEnd in currentLine:
            headerFound = False
            # print("HeaderFound: {}".format(headerFound))

        # Random comment found here which had no comment associated with it
        if (headerFound == True) and (searchWrapperStart in currentLine):
            # print("Found data: {}".format(currentLine))
            lineData = line
            break
        else:
            lineData = ""

    # If no result were found, return empty string
    if lineData ==  "":
        return""

    # Check to see if the XML element contains any data
    # An empty line will contain "/>" only, once the start wrapper is removed
    currentLine = lineData.lower()

    currentLine = currentLine.replace(searchWrapperStart, "")
    currentLine = currentLine.strip()

    if currentLine == "/>":
        emptyXmlElement = True
    else:
        emptyXmlElement = False

    if emptyXmlElement == False:
        # Extract the data from the XML element
        lengthLineData = len(lineData)
        lengthFirstWrapper = len(searchWrapperStart)
        lengthSecondWrapper = len(searchWrapperEnd) + 1

        # Strip the second wrapper off of the line of text
        firstSlice = lineData[:-lengthSecondWrapper]
        # print("FirstSlice: {}".format(firstSlice))
        firstSliceLength = len(firstSlice)

        # Find the string length to the end of the first wrapper
        counter = firstSliceLength - 1
        while counter > 0:

            if firstSlice[counter] == ">":
                firstWrapperEndIndex = counter + 1
                break
            counter -= 1

        # Obtain the measurement value from the XML String
        value = firstSlice[firstWrapperEndIndex:]
        valueLength = len(value)

        # Obtain the XML string up to the line
        stringWithoutValue = firstSlice[:-valueLength]

        # Only return the formatted empty line if instructed to do so
        if returnLine == False:
            return (value)
        else:
            # Create a XML string line which is prepped for having a new value inserted
            outputXMLstring = stringWithoutValue + "val" + searchWrapperEnd
            return (value, outputXMLstring)


    else:
        if returnLine == False:
            return (" ")
        else:
            # Create a XML string line which is prepped for having a new value inserted
            outputXMLstring = searchWrapperStart + "val/>"
            return (value, outputXMLstring)

def extractXmlData(xmlDataLine, wrapper):
    xmlDataLine = xmlDataLine.lower()
    xmlDataLine = xmlDataLine.strip()
    wrapper = wrapper.lower()

    startWrapper = "<" + wrapper + ">"
    endWrapper = "</" + wrapper + ">"
    nullStartWrapper = "<" + wrapper
    nullEndWrapper = "/>"

    cleanData = xmlDataLine.replace(startWrapper,"")
    cleanData = cleanData.replace(endWrapper, "")
    cleanData = cleanData.replace(nullStartWrapper, "")
    cleanData = cleanData.replace(nullEndWrapper, "")

    return str(cleanData)

def exportXmlToExcel(xmlList,cfgFilename,excelSpreadsheet,standardsList):
    import xlwings as xw
    from xlwings.constants import InsertShiftDirection
    # xmlList = []
    # xmlList = readTxtFile("C:/Users/Micah/PycharmProjects/LinearityCal_Fluke96270A/8481A.XML")

    # =========================== Open and Read Configuration File ===============================

    # cfgFilename = "C:/Users/Micah/PycharmProjects/LinearityCal_Fluke96270A/PS-Cal-Corrector.cfg"

    powerRefUnc = readConfigFile(cfgFilename, "powerRefUnc", sFunc="")
    excelTemplateFile = excelSpreadsheet

    print(excelTemplateFile)
    # input("Pause:")

    # =========================== Load in cert header information =================================
    header = "UUTHeader"
    wrapper = "ModelNumber"
    # print(xmlList)

    wrapper = "manufacturer"
    manufacturer = findAndExtractValueFromXML(header, wrapper, xmlList)
    print("Returned >{}<".format(manufacturer))

    wrapper = "modelNumber"
    modelNumber = findAndExtractValueFromXML(header, wrapper, xmlList)
    print("Returned >{}<".format(modelNumber))

    wrapper = "Description"
    description = findAndExtractValueFromXML(header, wrapper, xmlList)
    print("Returned >{}<".format(description))

    wrapper = "SerialNumber"
    serialNumber = findAndExtractValueFromXML(header, wrapper, xmlList)
    print("Returned >{}<".format(serialNumber))

    wrapper = "AssetNumber"
    assetNumber = findAndExtractValueFromXML(header, wrapper, xmlList)
    print("Returned >{}<".format(assetNumber))

    header = "ProcedureHeader"
    wrapper = "JobOrderNumber"
    jobOrderNumber = findAndExtractValueFromXML(header, wrapper, xmlList)
    print("Returned >{}<".format(jobOrderNumber))

    header = "CalibrationHeader"
    wrapper = "ProcedureName"
    procedureName = findAndExtractValueFromXML(header, wrapper, xmlList)
    print("Returned >{}<".format(procedureName))

    wrapper = "CalibrationDate"
    calibrationDate = findAndExtractValueFromXML(header, wrapper, xmlList)
    print("Returned >{}<".format(calibrationDate))

    wrapper = "CalibrationType"
    calibrationType = findAndExtractValueFromXML(header, wrapper, xmlList)
    print("Returned >{}<".format(calibrationType))

    wrapper = "CalibrationTechnician"
    technician = findAndExtractValueFromXML(header, wrapper, xmlList)
    print("Returned >{}<".format(technician))

    wrapper = "PinDepth>"
    pinDepth = findAndExtractValueFromXML(header, wrapper, xmlList)
    print("Returned >{}<".format(pinDepth))

    # =========================== Insert Header Information =======================================
    wb = xw.Book(excelTemplateFile)
    wb.sheets["Table 1"].activate()

    sht = wb.sheets["Table 1"]

    xw.Range("C12").value = manufacturer
    xw.Range("C13").value = modelNumber
    xw.Range("C14").value = description
    xw.Range("C15").value = serialNumber
    xw.Range("C16").value = assetNumber
    xw.Range("M12").value = jobOrderNumber
    xw.Range("M13").value = calibrationDate
    xw.Range("M14").value = technician
    xw.Range("M15").value = calibrationType
    xw.Range("M16").value = pinDepth

    # input("Pause:")

    # =========================== Get Standards List =============================================

    standardsList = standardsList

    # =========================== Insert data into Excel =========================================

    wb = xw.Book(excelTemplateFile)
    wb.sheets["Table 1"].activate()

    sht = wb.sheets["Table 1"]

    searchColumn = "B"
    searchString = "Standards Start"
    for i in range(1, 100, 1):
        startColumn = i
        cell = searchColumn + str(i)
        cellData = xw.Range(cell).value

        if cellData == searchString:
            startRow = i

            print("Found Standard Start Column: {}".format(cell))
            break

    for i, currentItem in enumerate(standardsList):
        writeRow = int(startRow) + int(i)
        inputList = currentItem.split(",")
        print(inputList)

        print("writeRow: {}".format(writeRow))
        print("index: {}".format(i))

        cell = "B{}".format(writeRow)
        print("Cell: {}".format(cell))
        print("I: {}".format(i))
        xw.Range(cell).value = inputList[0]

        cell = "C" + str(writeRow)
        xw.Range(cell).value = inputList[1]

        cell = "G" + str(writeRow)
        xw.Range(cell).value = inputList[2]

        cell = "K" + str(writeRow)
        xw.Range(cell).value = inputList[3]

        writeRow += 1

        rangeVal = "A" + str(writeRow) + ":M" + str(writeRow)
        print("rangeVal: {} ".format(rangeVal))
        # input()
        sht.range(rangeVal).api.Insert(InsertShiftDirection.xlShiftDown)

    # =========================== Load in Rho information =========================================

    rhoFreq = []
    rhoMsd = []
    rhoLimit = []
    rhoUnc = []
    rhoMag = []
    rhoPhase = []
    rhoPF = []
    searchHeader = "RhoData"
    for index, line in enumerate(xmlList):

        searchHeaderEnd = "</" + searchHeader + ">"
        searchHeaderEnd = searchHeaderEnd.lower()
        searchHeader = searchHeader.lower()
        searchHeaderStart = "<" + searchHeader
        currentLine = line.lower()

        if searchHeaderStart in currentLine:
            # print(currentLine)
            for i in range(1, 20, 1):
                rhoFound = False

                print(i)
                innerIndex = index + i
                data = xmlList[innerIndex].lower()
                print(data)

                if searchHeaderEnd in data:
                    # print("break")
                    break

                # print("Before: {}".format(data))
                searchTerm = "frequency"
                searchTerm = searchTerm.lower()
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    rhoFreq.append(strippedData)
                # print("After: {}".format(strippedData))

                searchTerm = "Rho_Limit"
                searchTerm = searchTerm.lower()
                # print(data)
                if searchTerm in data:
                    print(2)
                    strippedData = extractXmlData(data, searchTerm)
                    rhoLimit.append(strippedData)
                    rhoFound = True

                searchTerm = "Rho_Uncertainty"
                searchTerm = searchTerm.lower()
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    rhoUnc.append(strippedData)
                    rhoFound = True

                searchTerm = "Rho"
                searchTerm = searchTerm.lower()
                if (searchTerm in data) and (rhoFound == False):
                    data = extractXmlData(data, "rho")
                    rhoMsd.append(data)

                searchTerm = "Magnitude"
                searchTerm = searchTerm.lower()
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    rhoMag.append(strippedData)

                searchTerm = "Phase"
                searchTerm = searchTerm.lower()
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    rhoPhase.append(strippedData)

                searchTerm = "Pass_Fail"
                searchTerm = searchTerm.lower()
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    rhoPF.append(strippedData)

    print("Freq: {}".format(rhoFreq))
    print("msd: {}".format(rhoMsd))
    print("Limit: {}".format(rhoLimit))
    print("Unc: {}".format(rhoUnc))
    print("Magnitude: {}".format(rhoMag))
    print("Phase: {}".format(rhoPhase))
    print("Pass Fail: {}".format(rhoPF))

    # =========================== Insert data into Excel =========================================

    wb = xw.Book(excelTemplateFile)
    wb.sheets["Table 1"].activate()

    sht = wb.sheets["Table 1"]

    searchColumn = "B"
    searchString = "Rho Start"
    for i in range(1, 100, 1):
        startColumn = i
        cell = searchColumn + str(i)
        cellData = xw.Range(cell).value

        if cellData == searchString:
            startRow = i

            print("Found Rho Start Column: {}".format(cell))
            break

    for i, index in enumerate(rhoFreq):
        writeRow = int(startRow) + int(i)

        print("writeRow: {}".format(writeRow))
        print("index: {}".format(i))

        cell = "B{}".format(writeRow)
        print("Cell: {}".format(cell))
        print("I: {}".format(i))
        xw.Range(cell).value = index

        cell = "C" + str(writeRow)
        xw.Range(cell).value = rhoMsd[i]

        cell = "E" + str(writeRow)
        xw.Range(cell).value = rhoLimit[i]

        cell = "G" + str(writeRow)
        xw.Range(cell).value = rhoUnc[i]

        cell = "I" + str(writeRow)
        xw.Range(cell).value = rhoMag[i]

        cell = "K" + str(writeRow)
        xw.Range(cell).value = rhoPhase[i]

        cell = "M" + str(writeRow)
        xw.Range(cell).value = rhoPF[i]

        writeRow += 1

        rangeVal = "A" + str(writeRow) + ":M" + str(writeRow)
        print("rangeVal: {} ".format(rangeVal))
        # input()
        sht.range(rangeVal).api.Insert(InsertShiftDirection.xlShiftDown)

    # =========================== Load in Cal Factor information =========================================

    cfFreq = []
    cfMsd = []
    cfUnc = []
    cfDb = []
    searchHeader = "CalFactor"
    for index, line in enumerate(xmlList):

        searchHeaderEnd = "</" + searchHeader + ">"
        searchHeaderEnd = searchHeaderEnd.lower()
        searchHeader = searchHeader.lower()
        searchHeaderStart = "<" + searchHeader
        currentLine = line.lower()
        extraSearchFilter = "msdata"

        if (searchHeaderStart in currentLine) and (extraSearchFilter in currentLine):
            # print(currentLine)
            cfFoundCounter = 0
            for i in range(1, 20, 1):
                searchFound = False

                print(i)
                innerIndex = index + i
                data = xmlList[innerIndex].lower()
                print(data)

                if (searchHeaderEnd in data) and (cfFoundCounter > 0):
                    # print("break")
                    break

                # print("Before: {}".format(data))
                searchTerm = "frequency"
                searchTerm = searchTerm.lower()
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    cfFreq.append(strippedData)
                # print("After: {}".format(strippedData))

                searchTerm = "CalFactor"
                searchTerm = searchTerm.lower()
                # print(data)
                if searchTerm in data:
                    print("Found CF")
                    cfFoundCounter += 1
                    strippedData = extractXmlData(data, searchTerm)
                    cfMsd.append(strippedData)

                searchTerm = "Uncertainty"
                searchTerm = searchTerm.lower()
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    cfUnc.append(strippedData)

                searchTerm = "dB"
                searchTerm = searchTerm.lower()
                if (searchTerm in data):
                    print("dB Search Found: {}".format(data))
                    data = extractXmlData(data, searchTerm)
                    print("dB data: {}".format(data))
                    cfDb.append(data)

    print("cfFreq: {}".format(cfFreq))
    print("cfMsd: {}".format(cfMsd))
    print("cfUnc: {}".format(cfUnc))
    print("cfDb: {}".format(cfDb))

    # input("Pause:")

    # =========================== Insert data into Excel =========================================

    wb = xw.Book(excelTemplateFile)
    wb.sheets["Table 1"].activate()

    sht = wb.sheets["Table 1"]

    searchColumn = "B"
    searchString = "CF Start"
    for i in range(1, 100, 1):
        startColumn = i
        cell = searchColumn + str(i)
        cellData = xw.Range(cell).value

        if cellData == searchString:
            startRow = i

            print("Found CF Start Column: {}".format(cell))
            break

    for i, index in enumerate(cfFreq):
        writeRow = int(startRow) + int(i)

        print("writeRow: {}".format(writeRow))
        print("index: {}".format(i))

        cell = "B{}".format(writeRow)
        print("Cell: {}".format(cell))
        print("I: {}".format(i))
        xw.Range(cell).value = index

        cell = "C" + str(writeRow)
        xw.Range(cell).value = cfMsd[i]

        cell = "E" + str(writeRow)
        xw.Range(cell).value = cfUnc[i]

        cell = "G" + str(writeRow)
        xw.Range(cell).value = cfDb[i]

        writeRow += 1

        rangeVal = "A" + str(writeRow) + ":M" + str(writeRow)
        print("rangeVal: {} ".format(rangeVal))
        # input()
        sht.range(rangeVal).api.Insert(InsertShiftDirection.xlShiftDown)

    # =========================== Load-in Linearity information =========================================

    linFound = False
    linNominal = []
    linMsd = []
    linLimits = []
    linUnc = []
    linPF = []
    searchHeader = "Linearity"
    for index, line in enumerate(xmlList):

        searchHeaderEnd = "</" + searchHeader + ">"
        searchHeaderEnd = searchHeaderEnd.lower()
        searchHeader = searchHeader.lower()
        searchHeaderStart = "<" + searchHeader
        currentLine = line.lower()
        extraSearchFilter = "msdata"

        if (searchHeaderStart in currentLine) and (extraSearchFilter in currentLine):
            # print(currentLine)
            cfFoundCounter = 0
            linFound = True
            for i in range(1, 20, 1):
                searchFound = False

                print(i)
                innerIndex = index + i
                data = xmlList[innerIndex].lower()
                print(data)

                if (searchHeaderEnd in data) and (cfFoundCounter > 0):
                    # print("break")
                    break

                # print("Before: {}".format(data))
                searchTerm = "Nominal_Power"
                searchTerm = searchTerm.lower()
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    linNominal.append(strippedData)
                # print("After: {}".format(strippedData))

                searchTerm = "Measured_Power"
                searchTerm = searchTerm.lower()
                # print(data)
                if searchTerm in data:
                    print("Found CF")
                    cfFoundCounter += 1
                    strippedData = extractXmlData(data, searchTerm)
                    linMsd.append(strippedData)

                searchTerm = "Limits"
                searchTerm = searchTerm.lower()
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    linLimits.append(strippedData)

                searchTerm = "Uncertainty"
                searchTerm = searchTerm.lower()
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    linUnc.append(strippedData)

                searchTerm = "Pass_Fail"
                searchTerm = searchTerm.lower()
                if (searchTerm in data):
                    print("dB Search Found: {}".format(data))
                    data = extractXmlData(data, searchTerm)
                    print("dB data: {}".format(data))
                    linPF.append(data)

    print("linNominal: {}".format(linNominal))
    print("linMsd: {}".format(linMsd))
    print("linLimits: {}".format(linLimits))
    print("linUnc: {}".format(linUnc))
    print("linPF: {}".format(linPF))

    # input("Pause:")

    if linFound == True:

        # =========================== Insert data into Excel =========================================

        wb = xw.Book(excelTemplateFile)
        wb.sheets["Table 1"].activate()

        sht = wb.sheets["Table 1"]

        searchColumn = "B"
        searchString = "Linearity Start"
        for i in range(1, 10000, 1):
            startColumn = i
            cell = searchColumn + str(i)
            cellData = xw.Range(cell).value

            if cellData == searchString:
                startRow = i

                print("Found Linearity Start Column: {}".format(cell))
                break

        for i, index in enumerate(linNominal):
            writeRow = int(startRow) + int(i)

            print("writeRow: {}".format(writeRow))
            print("index: {}".format(i))

            cell = "B{}".format(writeRow)
            print("Cell: {}".format(cell))
            print("I: {}".format(i))
            xw.Range(cell).value = index

            cell = "C" + str(writeRow)
            xw.Range(cell).value = linMsd[i]

            cell = "E" + str(writeRow)
            xw.Range(cell).value = linLimits[i]

            cell = "G" + str(writeRow)
            xw.Range(cell).value = linUnc[i]

            cell = "M" + str(writeRow)
            xw.Range(cell).value = linPF[i]

            writeRow += 1

            rangeVal = "A" + str(writeRow) + ":M" + str(writeRow)
            print("rangeVal: {} ".format(rangeVal))
            # input()
            sht.range(rangeVal).api.Insert(InsertShiftDirection.xlShiftDown)
    else:
        # Delete the linearity section out of the excel spreadsheet
        wb = xw.Book(excelTemplateFile)
        wb.sheets["Table 1"].activate()

        sht = wb.sheets["Table 1"]

        searchColumn = "B"
        searchString = "Linearity"
        for i in range(1, 10000, 1):
            startColumn = i
            cell = searchColumn + str(i)
            cellData = xw.Range(cell).value

            if cellData == searchString:
                startRow = i

                print("Found Linearity Section Start Column: {}".format(cell))
                break

        endRow = startRow + 6

        rangeVal = "A" + str(startRow) + ":M" + str(endRow)
        print("rangeVal: {}".format(rangeVal))
        try:
            sht.range(rangeVal).api.delete()
        except:
            print("Known error; move on.")

    # =========================== Load in Absolute Power Reference Information ==========================

    aprFound = False
    aprFreq = []
    aprRef = []
    aprMsd = []
    aprUL = []
    aprLL = []
    aprUnc = []
    aprPF = []
    searchHeader = "PowerRef"
    for index, line in enumerate(xmlList):

        searchHeaderEnd = "</" + searchHeader + ">"
        searchHeaderEnd = searchHeaderEnd.lower()
        searchHeader = searchHeader.lower()
        searchHeaderStart = "<" + searchHeader
        currentLine = line.lower()
        extraSearchFilter = "msdata"

        if (searchHeaderStart in currentLine) and (extraSearchFilter in currentLine):
            # print(currentLine)
            cfFoundCounter = 0
            aprFound = True
            for i in range(1, 20, 1):
                searchFound = False

                print(i)
                innerIndex = index + i
                data = xmlList[innerIndex].lower()
                print(data)

                if (searchHeaderEnd in data):
                    # print("break")
                    break

                # print("Before: {}".format(data))
                searchTerm = "Frequency"
                searchTerm = searchTerm.lower()
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    aprFreq.append(strippedData)
                # print("After: {}".format(strippedData))

                searchTerm = "RefPower"
                searchTerm = searchTerm.lower()
                # print(data)
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    aprRef.append(strippedData)

                searchTerm = "MeasurePower"
                searchTerm = searchTerm.lower()
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    aprMsd.append(strippedData)

                searchTerm = "UpperLimit"
                searchTerm = searchTerm.lower()
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    aprUL.append(strippedData)

                searchTerm = "LowerLimit"
                searchTerm = searchTerm.lower()
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    aprLL.append(strippedData)

                searchTerm = "RefPower_Unc"
                searchTerm = searchTerm.lower()
                if searchTerm in data:
                    strippedData = extractXmlData(data, searchTerm)
                    # Insert Default Uncertainty Value From Configuration File
                    if strippedData == "- -":
                        strippedData = powerRefUnc
                    aprUnc.append(strippedData)

                searchTerm = "Pass_Fail"
                searchTerm = searchTerm.lower()
                if (searchTerm in data):
                    print("dB Search Found: {}".format(data))
                    data = extractXmlData(data, searchTerm)
                    print("dB data: {}".format(data))
                    aprPF.append(data)

    print("aprFreq: {}".format(aprFreq))
    print("aprRef: {}".format(aprRef))
    print("aprMsd: {}".format(aprMsd))
    print("aprUL: {}".format(aprUL))
    print("aprLL: {}".format(aprLL))
    print("aprUnc: {}".format(aprUnc))
    print("aprPF: {}".format(aprPF))

    # input("Pause:")

    if aprFound == True:

        # =========================== Insert data into Excel =========================================

        wb = xw.Book(excelTemplateFile)
        wb.sheets["Table 1"].activate()

        sht = wb.sheets["Table 1"]

        searchColumn = "B"
        searchString = "Power Ref Start"
        for i in range(1, 10000, 1):
            startColumn = i
            cell = searchColumn + str(i)
            cellData = xw.Range(cell).value

            if cellData == searchString:
                startRow = i

                print("Found Absolute Power Reference Start Cell: {}".format(cell))
                break

        for i, index in enumerate(aprFreq):
            writeRow = int(startRow) + int(i)

            print("writeRow: {}".format(writeRow))
            print("index: {}".format(i))

            cell = "B{}".format(writeRow)
            print("Cell: {}".format(cell))
            print("I: {}".format(i))
            xw.Range(cell).value = index

            cell = "C" + str(writeRow)
            xw.Range(cell).value = aprRef[i]

            cell = "E" + str(writeRow)
            xw.Range(cell).value = aprMsd[i]

            cell = "G" + str(writeRow)
            xw.Range(cell).value = aprLL[i]

            cell = "I" + str(writeRow)
            xw.Range(cell).value = aprUL[i]

            cell = "K" + str(writeRow)
            xw.Range(cell).value = aprUnc[i]

            cell = "M" + str(writeRow)
            xw.Range(cell).value = aprPF[i]

            writeRow += 1

            rangeVal = "A" + str(writeRow) + ":M" + str(writeRow)
            print("rangeVal: {} ".format(rangeVal))
            # input()
            sht.range(rangeVal).api.Insert(InsertShiftDirection.xlShiftDown)
    else:
        # Delete the linearity section out of the excel spreadsheet
        wb = xw.Book(excelTemplateFile)
        wb.sheets["Table 1"].activate()

        sht = wb.sheets["Table 1"]

        searchColumn = "B"
        searchString = "Absolute Power Reference"
        for i in range(1, 10000, 1):
            startColumn = i
            cell = searchColumn + str(i)
            cellData = xw.Range(cell).value

            if cellData == searchString:
                startRow = i

                print("Found Absolute Power Reference Section Start Cell: {}".format(cell))
                break

        endRow = startRow + 6

        rangeVal = "A" + str(startRow) + ":M" + str(endRow)
        print("rangeVal: {}".format(rangeVal))
        try:
            sht.range(rangeVal).api.delete()
        except:
            print("Known error; move on.")

def setCalibrationStandardList(listOfStandards):
    # Requires: from datetime import *
    from datetime import datetime
    from datetime import date

    standardsList = listOfStandards

    # Remove expired standards
    currentDate = datetime.date(datetime.now())

    expiredStandards = []
    for index, i in enumerate(standardsList):
        listItem = i.split(",")

        checkLastItem = listItem[-1]
        checkLastItem = checkLastItem.lower()
        checkLastItem = checkLastItem[0]

        if checkLastItem == "d":
            del listItem[-1]

        currentItemDate = listItem[-1]
        currentItemDateList = currentItemDate.split("-")
        year = int(currentItemDateList[0])
        month = int(currentItemDateList[1])
        day = int(currentItemDateList[2])

        dateFromList = date(year, month, day)
        if currentDate > dateFromList:
            expiredStandards.append(i)
            del standardsList[index]

    if len(expiredStandards) > 0:
        print("true")
        print("==============================================================================")
        print("|                                                                            |")
        print("|                          EXPIRED SYSTEM STANDARDS                          |")
        print("|                                                                            |")
        print("==============================================================================")
        print("{:2}    {:15}{:30}{:15}{:10}".format(" #", "Model", "Description", "Asset Number", "Cal Due"))
        print("==============================================================================")
        for index, i in enumerate(expiredStandards):
            listItem = i.split(",")
            # print("{}\t\t{}\t\t\t{}\t\t\t{}".format(listItem[0],listItem[1],listItem[2],listItem[3]))
            print("{:2}:   {:15}{:30}{:15}{:10}".format(index, listItem[0], listItem[1], listItem[2], listItem[3]))
        print("==============================================================================")
        print("The standards noted above are EXPIRED and cannot be selected until the configuration file is updated!")
        input("> Press Enter to continue...")

    # Populate default standards
    selectedStandards = []
    for index, i in enumerate(standardsList):
        listItem = i.split(",")
        defaultFlag = listItem[-1]
        defaultFlag = defaultFlag.lower()
        defaultFlag = defaultFlag[0]
        if defaultFlag == "d":
            selectedStandards.append(i)

    print("==============================================================================")
    print("|                                                                            |")
    print("|                        CONFIGURED SYSTEM STANDARDS                         |")
    print("|                                                                            |")
    print("==============================================================================")
    print("{:2}    {:15}{:30}{:15}{:10}".format(" #", "Model", "Description", "Asset Number", "Cal Due"))
    print("==============================================================================")
    for index, i in enumerate(standardsList):
        listItem = i.split(",")
        # print("{}\t\t{}\t\t\t{}\t\t\t{}".format(listItem[0],listItem[1],listItem[2],listItem[3]))
        print("{:2}:   {:15}{:30}{:15}{:10}".format(index, listItem[0], listItem[1], listItem[2], listItem[3]))
    print("")

    selection = ""
    while selection != "c":
        print("==============================================================================")
        print("|                                                                            |")
        print("|               SELECTED STANDARDS (used for this calibration)               |")
        print("|                                                                            |")
        print("==============================================================================")
        print("{:2}    {:15}{:30}{:15}{:10}".format(" #", "Model", "Description", "Asset Number", "Cal Due"))
        print("==============================================================================")
        for index, i in enumerate(selectedStandards):
            listItem = i.split(",")
            # print("{}\t\t{}\t\t\t{}\t\t\t{}".format(listItem[0],listItem[1],listItem[2],listItem[3]))
            print("{:2}:   {:15}{:30}{:15}{:10}".format(index, listItem[0], listItem[1], listItem[2], listItem[3]))

        print("==============================================================================")
        selection = input("Please Select (A)dd, (R)emove, or (C)onfirm Selection: ")
        selection = selection.lower()

        if selection != "a" and selection != "r" and selection != "c":
            print("")
            print("Valid options are (A)dd, (R)emove), or (C)onfirm Selection only!")
            print("i.e. you must enter {}, {}, or {}. So try again...".format("a", "r", "c"))
            print("")
            print(
                "See, these are the things which I, the programmer, must anticipate; otherwise who knows what crazy nonsense would go on!")
            print(
                "Seriously, this simple \"standard selection\" section of code is like, 100 lines long! A lot of it exists")
            print("just to handle user input entry errors.")
            print("I really need to get my act together and finish learning how to program GUIs in Python.")

        if selection == "a":
            print("==============================================================================")
            print("|                                                                            |")
            print("|                          SYSTEM STANDARDS LIST                             |")
            print("|                                                                            |")
            print("==============================================================================")
            print("{:2}    {:15}{:30}{:15}{:10}".format(" #", "Model", "Description", "Asset Number", "Cal Due"))
            print("==============================================================================")
            for index, i in enumerate(standardsList):
                listItem = i.split(",")
                # print("{}\t\t{}\t\t\t{}\t\t\t{}".format(listItem[0],listItem[1],listItem[2],listItem[3]))
                print("{:2}:   {:15}{:30}{:15}{:10}".format(index, listItem[0], listItem[1], listItem[2], listItem[3]))
            print("==============================================================================")

            selection = -1
            while selection < 0 or selection > index:
                selection = int(input("Enter the item # of the standard to add (0 through {}): ".format(index)))
                if selection < 0 or selection > index:
                    print("You can only choose item number 0 through {}!".format(index))
                    print("Try again...")

            currentSelectedItem = standardsList[selection]
            print(currentSelectedItem)
            selectedStandards.append(currentSelectedItem)

        if selection == "r":
            print("==============================================================================")
            print("|                                                                            |")
            print("|                        DELETE A SELECTED STANDARDS                         |")
            print("|                                                                            |")
            print("==============================================================================")
            print("{:2}    {:15}{:30}{:15}{:10}".format(" #", "Model", "Description", "Asset Number", "Cal Due"))
            print("==============================================================================")
            for index, i in enumerate(selectedStandards):
                listItem = i.split(",")
                # print("{}\t\t{}\t\t\t{}\t\t\t{}".format(listItem[0],listItem[1],listItem[2],listItem[3]))
                print("{:2}:   {:15}{:30}{:15}{:10}".format(index, listItem[0], listItem[1], listItem[2], listItem[3]))
            print("==============================================================================")

            selection = -1
            while selection < 0 or selection > index:
                selection = int(input("Enter the item # of the standard to remove (0 through {}): ".format(index)))
                if selection < 0 or selection > index:
                    print("You can only choose item number 0 through {}!".format(index))
                    print("Try again...")

            del selectedStandards[selection]

    return selectedStandards

# Start Program =================================================================
# Set initial program variables --------------------
logFile = "PS-Cal_Corrector_Log.txt"
if os.path.isfile(logFile):
    # print ("Log File Exists")
    logFile = logFile
else:
    create_log(logFile)

cwd = os.getcwd() + "\\"
clear()
userInterfaceHeader(programName, version, cwd, logFile)

print(" ____  ____         ____      _              ")
print("|  _ \/ ___|       / ___|__ _| |             ")
print("| |_) \___ \ _____| |   / _` | |             ")
print("|  __/ ___) |_____| |__| (_| | |             ")
print("|_|   |____/       \____\__,_|_|")
print("  ____                         _")
print(" / ___|___  _ __ _ __ ___  ___| |_ ___  _ __")
print("| |   / _ \| '__| '__/ _ \/ __| __/ _ \| '__|")
print("| |__| (_) | |  | | |  __/ (__| || (_) | |")
print(" \____\___/|_|  |_|  \___|\___|\__\___/|_| ")
print("=======================================================================")
time.sleep(2)

# Pull in settings from the config file ------------
configFile = "PS-Cal-Corrector.cfg"

try:
    debug = readConfigFile(configFile, "debug", "int")
    writeLog("Debug on or off: {}.".format(debug), logFile)
    PS_CalResultsFolder = readConfigFile(configFile, "PS_CalResultsFolder")
    writeLog("PS_CalResultsFolder: {}.".format(PS_CalResultsFolder), logFile)
    archivePath = readConfigFile(configFile, "archivePath")
    writeLog("archivePath: {}.".format(archivePath), logFile)
    standardsDataFolder = readConfigFile(configFile, "standardsDataFolder")
    writeLog("standardsDataFolder: {}.".format(standardsDataFolder), logFile)
    interpReferenceMethod = readConfigFile(configFile, "interpReferenceMethod", "int")
    writeLog("Interpolation method: {}.".format(interpReferenceMethod), logFile)
    numberSigDigits = readConfigFile(configFile, "numberSigDigits", "int")
    writeLog("Number of significant digits to correct to: {}.".format(numberSigDigits), logFile)
    rhoBudgetTxtFile = readConfigFile(configFile, "rhoBudgetTxtFile")
    writeLog("Location of Rho budget: {}.".format(rhoBudgetTxtFile), logFile)
    cfBudgetTxtFile = readConfigFile(configFile, "cfBudgetTxtFile")
    writeLog("Location of Cal Factor budget: {}.".format(cfBudgetTxtFile), logFile)
    linBudgetTxtFile = readConfigFile(configFile, "linBudgetTxtFile")
    writeLog("Location of linearity budget: {}.".format(linBudgetTxtFile), logFile)
    excelTemplateFile = readConfigFile(configFile, "excelTemplateFile", sFunc="")
    writeLog("Location of excel template file: {}.".format(excelTemplateFile), logFile)
except:
    writeLog("Failed to load configuration file variables. Check that the file is present.", logFile)
    exit()


# Set debug flag
if debug == 1:
    debugBool = True
else:
    debugBool = False
writeLog("Debug flag set to {}.".format(debugBool), logFile)

# Load XML file for interpolation ---------------------------------------------
# if debugBool == True:
#     xmlFile = "test3.XML"
#     xmlFilePath = cwd + xmlFile
# else:
print("")
print("Use the file dialogue window to select the PS-Cal XML to be corrected...")
if debugBool == True:
    xmlFile = "debugData.xml"
    xmlFilePath = cwd + xmlFile

    # Split out the xmlFilePath to obtain the xmlFile name itself
    tempList = xmlFilePath.split("\\")
    xmlFile = tempList[-1]

else:
    time.sleep(2)
    extensionType = "*.XML"
    xmlFilePath = getFilePath(extensionType,initialDir=PS_CalResultsFolder,extensionDescription="PSCAL XML")
    writeLog("User selected the following file for correction: {}.".format(xmlFilePath), logFile)

    # Split out the xmlFilePath to obtain the xmlFile name itself
    tempList = xmlFilePath.split("/")
    xmlFile = tempList[-1]

# Read-in the XML file data to a list
xmlData = readTxtFile(xmlFilePath)


# Setup to allow for interpolation to occur via the alternate reference method (using the Standard's data as the ref)
# The alternate method was introduced later, so the code below only massages the xmlFile data into a state where
# it can be interpolated by the normal method. Thereafter the normal method is used.
if interpReferenceMethod == 2:
    if debugBool == True:
        standardDataFile = "3538.XML"
        xmlFilePath = cwd + xmlFile
    else:
        time.sleep(2)
        print("Use the file dialogue window to select the XML data file of the standard used for the sensor cal...")
        time.sleep(2)
        extensionType = "*.XML"
        standardDataFile = getFilePath(extensionType, initialDir=standardsDataFolder, extensionDescription="PSCAL XML")

    stdXMLData = readTxtFile(standardDataFile)
    writeLog("User selected the following standard cal data file: {}.".format(standardDataFile), logFile)

    # Pull the frequency of all available cal points from the standard data
    stdFreqList = []
    for index, line in enumerate(stdXMLData):

        if ("Data diffgr:id" in line):
            rFreq = stdXMLData[index + 1]
            rFreq = re.sub("[^0-9.]", "", rFreq)
            rFreq = float(rFreq)
            stdFreqList.append(rFreq)

    # print(stdFreqList)
    # input("Press Enter To Continue...")

    # Build a list of CF frequencies present in the XML calibration data
    cfFreqList = []
    for index, line in enumerate(xmlData):

        if ("CalFactor diffgr" in line):
            lineList = line.split(" ")
            cFreq = xmlData[index + 1]
            cFreq = re.sub("[^0-9.]", "", cFreq)
            cFreq = float(cFreq)
            cfFreqList.append(cFreq)

    # Compare the cfFreqList to the freqs contained in the standard's data to list which require interpolation
    requiredInterpList = []
    for index, i in enumerate(cfFreqList):

        # Check if the cf Freq is in the standard data list; if not then set required interp to 1
        try:
            tempIndex = stdFreqList.index(i)
            requiredInterpList.append(0)
        except:
            requiredInterpList.append(1)

    # Delete CF blocks from existing XML data for all frequencies that must be interpolated
    # Missing CF blocks is what triggers the normal method to know that interpolation must occur
    for index, i in enumerate(requiredInterpList):
        if i == 1:
            tempFreq = cfFreqList[index]

            for index2, line in enumerate(xmlData):
                if ("CalFactor diffgr" in line):
                    checkFreq = xmlData[index2 + 1]
                    checkFreq = re.sub("[^0-9.]", "", checkFreq)
                    checkFreq = float(checkFreq)
                    if checkFreq == tempFreq:
                        del xmlData[index2]
                        tempBool = False
                        while tempBool == False:
                            if ("CalFactor diffgr" in xmlData[index2]):
                                tempBool = True
                            else:
                                del xmlData[index2]

# Create backup copy of existing XML file
tempBool = path.exists(archivePath)
if tempBool == True:
    tempBool = tempBool
else:
    os.mkdir(archivePath)


# Check if the file has already been backed up; if yes then create a unique name that will not overwrite the existing
# file
archiveFilePath = archivePath + xmlFile
tempBool = path.exists(archiveFilePath)
tempCounter = 1
while tempBool == True:

    archiveFilePath = archivePath + xmlFile

    filename = Path(archiveFilePath)
    filename_wo_ext = filename.with_suffix('')
    archiveFilePath = str(filename_wo_ext)
    archiveFilePath+=" - Copy(" + str(tempCounter) +")"

    filename = Path(archiveFilePath)
    filename_replace_ext = filename.with_suffix('.xml')
    archiveFilePath = filename_replace_ext

    tempBool = path.exists(archiveFilePath)
    tempCounter+=1

# Backup of the original file
print("xmlFilePath: {}".format(xmlFilePath))
print("archiveFilePath: {}".format(archiveFilePath))
dest = shutil.copyfile(xmlFilePath, archiveFilePath)
writeLog("Created backup of the original DUT calibration file at: {}.".format(archiveFilePath), logFile)

# Create a working copy of the Excel Template file for use by this routine
winTempFolder = tempfile.gettempdir()                       # Gets the windows temporary folder
tempDirectory = winTempFolder + "/temporary_sensorDataTemplate.xls".format()

dest = shutil.copyfile(excelTemplateFile, tempDirectory)
writeLog("Created backup of the original Excel Template File at: {}.".format(tempDirectory), logFile)

excelTemplateFile = tempDirectory

# Import Standards List and instruct user to select standards
standardsList = imporStdList(configFile)
selectedStandards = setCalibrationStandardList(standardsList)

writeLog("Technician confirmed the following standards list: {}.".format(selectedStandards), logFile)

# ===================================================================================================================
#                                    Check Against Uncertainty Budget Files
# ===================================================================================================================
writeLog("Started uncertainty budget check.", logFile)

for index, line in enumerate(xmlData):

    # Do uncertainty lookup for all Rho data
    searchHeaderStart = "<RhoData"
    searchHeaderEnd = "</RhoData>"
    if searchHeaderStart in line:

        # Obtain all the parameters necessary to perform the uncertainty lookup
        counter = 0
        while counter <= 8:
            counter += 1
            index2 = index + counter
            line2 = xmlData[index2]
            if debugBool == True:
                writeLog("Uncertainy Lookup Line2: {}".format(line2), logFile)

            searchTerm = "Frequency"
            firstWrapper = "<" + searchTerm + ">"
            secondWrapper = "</" + searchTerm + ">"
            if (firstWrapper in line2) and (secondWrapper in line2):
                line2 = line2.replace(",","")
                # print("line2: {}".format(line2))
                value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line2)
                freqValue = float(value.replace(",",""))

            searchTerm = "Rho"
            firstWrapper = "<" + searchTerm + ">"
            secondWrapper = "</" + searchTerm + ">"
            try:
                if (firstWrapper in line2) and (secondWrapper in line2):
                    value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line2)
                    rhoValue = float(value)
            except:
                writeLog("Could note parse Rho Measured Value: value at frequency: {} Hz".format(freqValue), logFile)
                break

            try:
                searchTerm = "Rho_Uncertainty"
                firstWrapper = "<" + searchTerm + ">"
                secondWrapper = "</" + searchTerm + ">"
                if (firstWrapper in line2) and (secondWrapper in line2):
                    value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line2)
                    rhoUnc = float(value)
            except:
                writeLog("Could note parse Rho Uncertainty value at frequency: {} Hz".format(freqValue), logFile)
                break

            if (searchHeaderEnd in line2) and not (searchHeaderStart in line2):
                break


        try:
            # Perform the uncertainty value lookup
            newUncValue = checkUncBudget(rhoBudgetTxtFile, rhoUnc, freqValue, rhoValue)

            # Apply the uncertainty value returned by the lookup
            counter = 0
            while counter <= 8:
                counter += 1
                index2 = index + counter
                line2 = xmlData[index2]

                searchTerm = "Rho_Uncertainty"
                firstWrapper = "<" + searchTerm + ">"
                secondWrapper = "</" + searchTerm + ">"
                if (firstWrapper in line2) and (secondWrapper in line2):
                    value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line2)


                    value = str(newUncValue)

                    outputXMLstring = outputXMLstring.replace("val", value)
                    line = outputXMLstring + "\n"
                    xmlData[index2] = line
                    break
        except:
            writeLog("Rho budget lookup not performed at: {} Hz".format(freqValue), logFile)


    # Do uncertainty lookup for all CF data
    searchHeaderStart = "<CalFactor"
    searchHeaderEnd = "</CalFactor>"
    if searchHeaderStart in line:

        # Obtain all the parameters necessary to perform the uncertainty lookup
        counter = 0
        while counter <= 8:
            counter += 1
            index2 = index + counter
            line2 = xmlData[index2]

            searchTerm = "Frequency"
            firstWrapper = "<" + searchTerm + ">"
            secondWrapper = "</" + searchTerm + ">"
            if (firstWrapper in line2) and (secondWrapper in line2):
                line2 = line2.replace(",", "")
                # print("line2: {}".format(line2))
                value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line2)
                freqValue = float(value.replace(",", ""))

            try:
                searchTerm = "Uncertainty"
                firstWrapper = "<" + searchTerm + ">"
                secondWrapper = "</" + searchTerm + ">"
                if (firstWrapper in line2) and (secondWrapper in line2):
                    value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line2)
                    cfUnc = float(value)
            except:
                writeLog("Could note parse CF Uncertainty value at frequency: {} Hz".format(freqValue), logFile)
                break

            if (searchHeaderEnd in line2) and not (searchHeaderStart in line2):
                break

        try:
            # Perform the uncertainty value lookup
            newUncValue = checkUncBudget(cfBudgetTxtFile, cfUnc, freqValue)

            # Apply the uncertainty value returned by the lookup
            counter = 0
            while counter <= 8:
                counter += 1
                index2 = index + counter
                line2 = xmlData[index2]

                searchTerm = "Uncertainty"
                firstWrapper = "<" + searchTerm + ">"
                secondWrapper = "</" + searchTerm + ">"
                if (firstWrapper in line2) and (secondWrapper in line2):
                    value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line2)

                    value = str(newUncValue)

                    outputXMLstring = outputXMLstring.replace("val", value)
                    line = outputXMLstring + "\n"
                    xmlData[index2] = line
                    break
        except:
            writeLog("CF budget lookup not performed at: {} Hz".format(freqValue), logFile)

    # Do uncertainty lookup for all Linearity data
    searchHeaderStart = "<Linearity"
    searchHeaderEnd = "</Linearity>"
    if searchHeaderStart in line:

        # Obtain all the parameters necessary to perform the uncertainty lookup
        counter = 0
        while counter <= 8:
            counter += 1
            index2 = index + counter
            line2 = xmlData[index2]

            try:
                searchTerm = "Measured_Power"
                firstWrapper = "<" + searchTerm + ">"
                secondWrapper = "</" + searchTerm + ">"
                if (firstWrapper in line2) and (secondWrapper in line2):
                    line2 = line2.replace(",", "")
                    # print("line2: {}".format(line2))
                    value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line2)
                    linValue = float(value.replace(",", ""))
            except:
                writeLog("Could note parse Linearity Measured Power value at frequency: {} Hz".format(freqValue), logFile)
                break

            try:
                searchTerm = "Uncertainty"
                firstWrapper = "<" + searchTerm + ">"
                secondWrapper = "</" + searchTerm + ">"
                if (firstWrapper in line2) and (secondWrapper in line2):
                    value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line2)
                    linUnc = float(value)
            except:
                writeLog("Could note parse Linearity Uncertainty value at frequency: {} Hz".format(freqValue),
                         logFile)
                break

            if (searchHeaderEnd in line2) and not (searchHeaderStart in line2):
                break

        try:
            # Perform the uncertainty value lookup
            freqValue = 50_000_000                          # Hard-coded because linearity is always at 50 MHz
            newUncValue = checkUncBudget(linBudgetTxtFile, linUnc, freqValue, linValue)

            # Apply the uncertainty value returned by the lookup
            counter = 0
            while counter <= 8:
                counter += 1
                index2 = index + counter
                line2 = xmlData[index2]

                searchTerm = "Uncertainty"
                firstWrapper = "<" + searchTerm + ">"
                secondWrapper = "</" + searchTerm + ">"
                if (firstWrapper in line2) and (secondWrapper in line2):
                    value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line2)

                    value = str(newUncValue)

                    outputXMLstring = outputXMLstring.replace("val", value)
                    line = outputXMLstring + "\n"
                    xmlData[index2] = line
                    break
        except:
            writeLog("Linearity budget lookup not performed at: {} Hz".format(freqValue), logFile)



# Write the XML data to a file
with open(xmlFilePath, 'w') as filehandle:
    for listItem in xmlData:
        # print(listItem)
        filehandle.write(listItem)

writeLog("Wrote results of the uncertainty budget lookup to the instrument file at: {}".format(xmlFilePath), logFile)
print("> Uncertainty lookup completed...")
print("")



# ===================================================================================================================
#                                         Check Significant Figures
# ===================================================================================================================
writeLog("Starting check of significant figures", logFile)

# Read-in the XML file data to a list
xmlDataNew = []
xmlData = readTxtFile(xmlFilePath)

# Correct Significant Figures in all Uncertainty
for index, line in enumerate(xmlData):

    searchTerm = "ProcedureName"
    firstWrapper = "<" + searchTerm + ">"
    secondWrapper = "</" + searchTerm + ">"
    if (firstWrapper in line) and (secondWrapper in line):
        value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line)

        value = "PS-Cal Corrected"

        outputXMLstring = outputXMLstring.replace("val", value)
        # print(outputXMLstring)
        line = outputXMLstring + "\n"

    try:
        searchTerm = "Rho_Uncertainty"
        firstWrapper = "<" + searchTerm + ">"
        secondWrapper = "</" + searchTerm + ">"
        if (firstWrapper in line) and (secondWrapper in line):
            value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line)

            value = str(setSigDigits(value, numberSigDigits))

            outputXMLstring = outputXMLstring.replace("val", value)
            line = outputXMLstring + "\n"
    except:
        writeLog("Rho Uncertainty Significant Digits not corrected at XML file line: ".format(index), logFile)

    try:
        searchTerm = "Uncertainty"
        firstWrapper = "<" + searchTerm + ">"
        secondWrapper = "</" + searchTerm + ">"
        if (firstWrapper in line) and (secondWrapper in line):
            value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line)

            value = str(setSigDigits(value, numberSigDigits))

            outputXMLstring = outputXMLstring.replace("val", value)
            line = outputXMLstring + "\n"
    except:
        writeLog("Cal Factor or Linearity Uncertainty Significant Digits not corrected at XML file line: ".format(index), logFile)

    xmlDataNew.append(line)

# Write the XML data to a file
with open(xmlFilePath, 'w') as filehandle:
    for listItem in xmlDataNew:
        # print(listItem)
        filehandle.write(listItem)

writeLog("Wrote results of the significant figures correction to the instrument file at: {}".format(xmlFilePath), logFile)
print("> Quantity of significant digits set to {}".format(numberSigDigits))
print("")


# ===================================================================================================================
#                                  Perform Interpolation of Data (if required)
# ===================================================================================================================
writeLog("Starting interpolation check of cal factor values, if required", logFile)

# Read-in the XML file data to a list
xmlDataNew = []
xmlData = readTxtFile(xmlFilePath)

rhoFreqList = []
cfFreqList = []
cfList = []
uncList = []
dbList = []

writeLog("Read in the existing Rho and CF points", logFile)
# Read in the existing Rho and CF data
for index, line in enumerate(xmlData):

    if ("RhoData diffgr" in line):
        lineList = line.split(" ")
        rFreq = xmlData[index + 1]
        rFreq = re.sub("[^0-9.]", "", rFreq)
        rFreq = float(rFreq)
        rhoFreqList.append(rFreq)

    if ("CalFactor diffgr" in line):
        lineList = line.split(" ")
        cFreq = xmlData[index + 1]
        cFreq = re.sub("[^0-9.]", "", cFreq)
        cFreq = float(cFreq)
        cfFreqList.append(cFreq)

        try:
            cf = xmlData[index + 2]
            cf = re.sub("[^0-9.]", "", cf)
            # print("cf: {}".format(cf))
            cf = float(cf)
            cfList.append(cf)
        except:
            writeLog("Interpolation failed at: {} Hz".format(rFreq), logFile)
            writeLog("This is most likely because the XML template is not fully run.", logFile)
            writeLog("However, all operations up to this point have been completed and are saved to the XML file.", logFile)
            print("Interpolation failed at: {} Hz".format(rFreq))
            print("This is most likely because the XML template is not fully run.")
            print("However, all operations up to this point have been completed and are saved to the XML file.")
            print("")
            print("XML File saved at: {}".format(xmlFilePath))
            print("")
            print("This program will close automatically in 5 seconds...")
            writeLog("Program ended prematurely!", logFile)
            time.sleep(5)
            exit()

        unc = xmlData[index + 3]
        unc = re.sub("[^0-9.]", "", unc)
        unc = float(unc)
        uncList.append(unc)

        db = xmlData[index + 5]
        db = re.sub("[^0-9.-]", "", db)
        db = float(db)
        dbList.append(db)



# Create new lists for each fields requiring interp which are equal length to the rhoFreq list
cfFreqListNew = []
cfListNew = []
uncListNew = []
dbListNew = []
requiredInterpList = []
for index, freq in enumerate(rhoFreqList):

    # If the point does not require interp then insert the already existing value
    if (freq in cfFreqList):
        cfFreqListNew.append(freq)

        tempIndex = cfFreqList.index(freq)
        tempValue = cfList[tempIndex]
        cfListNew.append(tempValue)

        tempValue = uncList[tempIndex]
        uncListNew.append(tempValue)

        tempValue = dbList[tempIndex]
        dbListNew.append(tempValue)

        # Tracks which points require interp (1 means required, 0 means not required)
        requiredInterpList.append(0)

    # If the point requires interp then fill it with 0 for now
    else:
        cfFreqListNew.append(float(0))
        cfListNew.append(float(0))
        uncListNew.append(float(0))
        dbListNew.append(float(0))

        # Tracks which points require interp (1 means required, 0 means not required)
        requiredInterpList.append(1)


xVals = []
yValsCF = []
yValsUNC = []
yValsDb = []
for index, freq in enumerate(rhoFreqList):

    if (freq in cfFreqList):
        xVals.append(freq)

        tempIndex = cfFreqListNew.index(freq)
        tempValue = cfListNew[tempIndex]
        yValsCF.append(tempValue)

        tempValue = uncListNew[tempIndex]
        yValsUNC.append(tempValue)

        tempValue = dbListNew[tempIndex]
        yValsDb.append(tempValue)

writeLog("Performing interpolation for any CF points which require it", logFile)
# Perform interpolation process for all missing points
performInterpolation()

writeLog("Copying existing CF XML data block for use in creating interpolated CF data points", logFile)
# Create Cal Factor Points XML Template by copying existing CF XML block into list
tempBool = False
cfBlockList = []
for index, line in enumerate(xmlData):

    if tempBool == True:
        break

    if ("CalFactor diffgr" in line):
        cfBlockList.append(line)

        tempCounter = 1
        while tempBool == False:
            line2 = xmlData[index + tempCounter]
            cfBlockList.append(line2)
            tempCounter += 1
            if ("</CalFactor>" in line2) and (not "<CalFactor>" in line2):
                tempBool = True

cfBlockLength = len(cfBlockList)


writeLog("Applying XML CF data chunk to the existing XML data as necessary", logFile)
# Determine where an interpolated CF block needs to be entered into the existing data
xmlDataNew = xmlData.copy()
for index, element in enumerate(requiredInterpList):

    if element == 1:
        freqToAdd = cfFreqListNew[index]
        CFToAdd = cfListNew[index]
        uncertaintyToAdd = uncListNew[index]
        dbToAdd = dbListNew[index]

        # print("freqToAdd {}".format(freqToAdd))

        for index2, line in enumerate(xmlData):

            if ("CalFactor diffgr" in line):
                # Get the current frequency of the CF block
                lineList = xmlData[index2 + 1]
                # print(lineList)
                freq = re.sub("[^0-9.]", "", lineList)
                freq = float(freq)
                #print("freq {}".format(freq))

                if freq > freqToAdd:

                    newCFBlock = editCFblock(cfBlockList, freqToAdd, CFToAdd, uncertaintyToAdd,dbToAdd)
                    tempIndex = index2
                    # print("newCFBlock {}".format(newCFBlock))
                    # print("cfBlockLength {}".format(cfBlockLength))
                    #
                    # print("tempIndex {}".format(tempIndex))

                    insertCalFactor(tempIndex,newCFBlock)
                    break

    # The original xmlData needs to be updated each time new data is added
    xmlData = xmlDataNew.copy()

writeLog("Going through the XML data and updating the XML row order to accomodate the new CF data chunks", logFile)
# Go through the XML data in the variable and update the row order of the CF data
tempBool = False
cfBlockList = []
rowOrderCounter = 0
calFactorCounter = 1
hiddenIndexCounter = 0
for index, line in enumerate(xmlDataNew):

    if tempBool == True:
        break

    if ("CalFactor diffgr" in line):

        # Split out the line so it can be searched and updated easily
        splitLine = line.split("\"")
        # print(splitLine)

        # Find the list element for the row number
        for index2, element in enumerate(splitLine):
            loweredElemet = element.lower()
            if ("roworder" in loweredElemet):
                splitLine[index2+1] = str(rowOrderCounter)

        # Iterate the newly updated splitLine list back into a string
        newLine = ""
        for index2, element in enumerate(splitLine):
            newLine+=str(element) + "\""
        # Delete the final " from the end of the string
        newLine = newLine[:-1:]

        # Update the existing CSV variable data with the updated newLine
        xmlDataNew[index] = newLine

        rowOrderCounter += 1
        calFactorCounter += 1
        hiddenIndexCounter += 1

# Write the XML data to a file
with open(xmlFilePath, 'w') as filehandle:
    for listItem in xmlDataNew:
        # print(listItem)
        filehandle.write(listItem)
writeLog("Wrote interpolated data to XML file at: {}".format(xmlFilePath), logFile)


writeLog("CF Interpolation process completed.", logFile)
print("> Interpolation check completed...")
print("")

# ===================================================================================================================
#                                  Attempt to export results into Excel template file
# ===================================================================================================================
writeLog("Starting Export of Data to Excel Template", logFile)
writeLog("Using Excel Spreadsheet: {}".format(excelTemplateFile), logFile)
print("> Attempting to export data to excel template...")



try:
    exportXmlToExcel(xmlDataNew, configFile, excelTemplateFile, selectedStandards)
    print("")
    print("> Successfully exported data to Excel.")
    print("")
    print("==============================================================")
    print("          - - Use Excel to save data as PDF - -")
    print("==============================================================")
    writeLog("Successfully exported data to Excel Template", logFile)
except:
    writeLog("Failed to write excel data to Excel template", logFile)
    print("")
    print("> Failed to write data to Excel template!")
    print("==============================================================")
    print("- - Open the XML file in PS-Cal to verify and save as PDF - -")
    print("==============================================================")



writeLog("Program output saved to: {}.".format(xmlFilePath), logFile)


print("")
print("XML Output file saved at: {}".format(xmlFilePath))
print("")

print("")
print("This program will close automatically in 5 seconds...")
writeLog("Program ended.", logFile)
time.sleep(5)
sys.exit()