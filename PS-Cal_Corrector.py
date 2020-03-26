# PS-Cal Corrector
# By Micah Hurd
programName = "PS-Cal Corrector"
version = 1.1

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
            print(2)

            if pRangePresent == True:
                tempList = rangeList[index].split(">")
                startPow = float(tempList[0])
                stopPow = float(tempList[1])
                print(3)

                if (uncMsmt >= startPow) and (uncMsmt <= stopPow):
                    tempUncListVal = uncList[index]
                    print(4)
                    if uncVal < tempUncListVal:
                        print(5)
                        uncVal = tempUncListVal
                        return uncVal
            else:
                tempUncListVal = uncList[index]
                if uncVal < tempUncListVal:
                    uncVal = tempUncListVal
                    return uncVal


    return uncVal

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

print("__________  _________         _________     _____  .____ ")
print("\______   \/   _____/         \_   ___ \   /  _  \ |    |")
print(" |     ___/\_____  \   ______ /    \  \/  /  /_\  \|    |")
print(" |    |    /        \ /_____/ \     \____/    |    \    |___")
print(" |____|   /_______  /          \______  /\____|__  /_______ \\")
print("                  \/                  \/         \/        \/")
print("   _____                                 .___.__  __")
print("  /  _  \   ____  ___________   ____   __| _/|__|/  |_  ___________")
print(" /  /_\  \_/ ___\/ ___\_  __ \_/ __ \ / __ | |  \   __\/  _ \_  __ \\")
print("/    |    \  \__\  \___|  | \/\  ___// /_/ | |  ||  | (  <_> )  | \/")
print("\____|__  /\___  >___  >__|    \___  >____ | |__||__|  \____/|__| ")
print("=======================================================================")


# Pull in settings from the config file ------------
configFile = "accreditor.cfg"
a = 1

debug = readConfigFile(configFile, "debug", "int")
PS_CalResultsFolder = readConfigFile(configFile, "PS_CalResultsFolder")
archivePath = readConfigFile(configFile, "archivePath")
standardsDataFolder = readConfigFile(configFile, "standardsDataFolder")
interpReferenceMethod = readConfigFile(configFile, "interpReferenceMethod", "int")
numberSigDigits = readConfigFile(configFile, "numberSigDigits", "int")
rhoBudgetTxtFile = readConfigFile(configFile, "rhoBudgetTxtFile")
cfBudgetTxtFile = readConfigFile(configFile, "cfBudgetTxtFile")
linBudgetTxtFile = readConfigFile(configFile, "linBudgetTxtFile")


# Set debug flag
if debug == 1:
    debugBool = True
else:
    debugBool = False

# Load XML file for interpolation ---------------------------------------------
if debugBool == True:
    xmlFile = "test3.XML"
    xmlFilePath = cwd + xmlFile
else:
    print("Use the file dialogue window to select the PS-Cal XML to be corrected...")
    extensionType = "*.XML"
    xmlFilePath = getFilePath(extensionType,initialDir=PS_CalResultsFolder,extensionDescription="PSCAL XML")

    # Split out the xmlFilePath to obtain the xmlFile name itself
    tempList = xmlFilePath.split("/")
    xmlFile = tempList[-1]

# Setup to allow for interpolation to occur via the alternate reference method (using the Standard's data as the ref)
# The alternate method was introduced later, so the code below only massages the xmlFile data into a state where
# it can be interpolated by the normal method. Thereafter the normal method is used.
if interpReferenceMethod == 2:
    if debugBool == True:
        standardDataFile = "3538.XML"
        xmlFilePath = cwd + xmlFile
    else:
        print("Use the file dialogue window to select the XML data file of the standard used for the sensor cal...")
        extensionType = "*.XML"
        standardDataFile = getFilePath(extensionType, initialDir=PS_CalResultsFolder, extensionDescription="PSCAL XML")

    stdXMLData = readXMLFile(standardDataFile)

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
dest = shutil.copyfile(xmlFilePath, archiveFilePath)

# ===================================================================================================================
#                                    Check Against Uncertainty Budget Files
# ===================================================================================================================
# Read-in the XML file data to a list
xmlData = readTxtFile(xmlFilePath)

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
            if (firstWrapper in line2) and (secondWrapper in line2):
                value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line2)
                rhoValue = float(value)

            searchTerm = "Rho_Uncertainty"
            firstWrapper = "<" + searchTerm + ">"
            secondWrapper = "</" + searchTerm + ">"
            if (firstWrapper in line2) and (secondWrapper in line2):
                value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line2)
                rhoUnc = float(value)

            if (searchHeaderEnd in line2) and not (searchHeaderStart in line2):
                break



        # Perform the uncertainty value lookup
        newUncValue = checkUncBudget(rhoBudgetTxtFile, rhoUnc, freqValue, rhoValue)
        print("newUncValue = {}".format(newUncValue))

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

            searchTerm = "Uncertainty"
            firstWrapper = "<" + searchTerm + ">"
            secondWrapper = "</" + searchTerm + ">"
            if (firstWrapper in line2) and (secondWrapper in line2):
                value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line2)
                cfUnc = float(value)

            if (searchHeaderEnd in line2) and not (searchHeaderStart in line2):
                break

        # Perform the uncertainty value lookup
        newUncValue = checkUncBudget(cfBudgetTxtFile, cfUnc, freqValue)
        print("newUncValue = {}".format(newUncValue))

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

    # Do uncertainty lookup for all CF data
    searchHeaderStart = "<Linearity"
    searchHeaderEnd = "</Linearity>"
    if searchHeaderStart in line:

        # Obtain all the parameters necessary to perform the uncertainty lookup
        counter = 0
        while counter <= 8:
            counter += 1
            index2 = index + counter
            line2 = xmlData[index2]

            searchTerm = "Measured_Power"
            firstWrapper = "<" + searchTerm + ">"
            secondWrapper = "</" + searchTerm + ">"
            if (firstWrapper in line2) and (secondWrapper in line2):
                line2 = line2.replace(",", "")
                # print("line2: {}".format(line2))
                value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line2)
                linValue = float(value.replace(",", ""))

            searchTerm = "Uncertainty"
            firstWrapper = "<" + searchTerm + ">"
            secondWrapper = "</" + searchTerm + ">"
            if (firstWrapper in line2) and (secondWrapper in line2):
                value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line2)
                linUnc = float(value)

            if (searchHeaderEnd in line2) and not (searchHeaderStart in line2):
                break

        # Perform the uncertainty value lookup
        freqValue = 50_000_000                          # Hard-coded because linearity is always at 50 MHz
        newUncValue = checkUncBudget(linBudgetTxtFile, linUnc, freqValue, linValue)
        print("newUncValue = {}".format(newUncValue))

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



# Write the XML data to a file
with open(xmlFilePath, 'w') as filehandle:
    for listItem in xmlData:
        # print(listItem)
        filehandle.write(listItem)

print("")
print("* * Uncertainty Lookup Completed * *")
print("")



# ===================================================================================================================
#                                         Check Significant Figures
# ===================================================================================================================

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
        print("Prepped Line: {} and Value: {}".format(outputXMLstring, value))

        value = "PS-Cal Corrected"

        outputXMLstring = outputXMLstring.replace("val", value)
        # print(outputXMLstring)
        line = outputXMLstring + "\n"

    searchTerm = "Rho_Uncertainty"
    firstWrapper = "<" + searchTerm + ">"
    secondWrapper = "</" + searchTerm + ">"
    if (firstWrapper in line) and (secondWrapper in line):
        value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line)
        print("Prepped Line: {} and Value: {}".format(outputXMLstring,value))

        value = str(setSigDigits(value, numberSigDigits))

        outputXMLstring = outputXMLstring.replace("val", value)
        print(outputXMLstring)
        line = outputXMLstring + "\n"

    searchTerm = "Uncertainty"
    firstWrapper = "<" + searchTerm + ">"
    secondWrapper = "</" + searchTerm + ">"
    if (firstWrapper in line) and (secondWrapper in line):
        value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line)
        print("Prepped Line: {} and Value: {}".format(outputXMLstring, value))

        value = str(setSigDigits(value, numberSigDigits))

        outputXMLstring = outputXMLstring.replace("val", value)
        print(outputXMLstring)
        line = outputXMLstring + "\n"

    xmlDataNew.append(line)

# Write the XML data to a file
with open(xmlFilePath, 'w') as filehandle:
    for listItem in xmlDataNew:
        # print(listItem)
        filehandle.write(listItem)

print("")
print("* * Quantity of significant digits set to two * *")
print("")


# ===================================================================================================================
#                                  Perform Interpolation of Data (if required)
# ===================================================================================================================

# Read-in the XML file data to a list
xmlDataNew = []
xmlData = readTxtFile(xmlFilePath)

rhoFreqList = []
cfFreqList = []
cfList = []
uncList = []
dbList = []

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

        cf = xmlData[index + 2]
        cf = re.sub("[^0-9.]", "", cf)
        # print("cf: {}".format(cf))
        cf = float(cf)
        cfList.append(cf)

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


# Perform interpolation process for all missing points
performInterpolation()

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


print("")
print("* * INTERPOLATION COMPLETED * *")
print("")
print("Output file saved at: {}".format(xmlFilePath))
print("")
print("- - Open the XML file in PS-Cal to verify and save as PDF - -")
print("")
print("This program will close automatically in 5 seconds...")
time.sleep(5)
sys.exit()