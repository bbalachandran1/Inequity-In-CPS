import requests
import pandas as pd
import html
import time
import socket
import errno

def read_data():
    filename = "gridview"

    df = pd.read_excel(filename + ".xlsx"   )
    newDf = df.to_numpy()
    bigList = []
    for i in range(0,len(newDf)):
        miniList=[]
        for p in range(len(newDf[0])):
            miniList.append(newDf[i][p])

        bigList.append(miniList)

    return bigList

def scrapeWeb(dataList,rowList):
    waittimer = 0
    failList = []
    for i in rowList:
        
        print("Currently on "+ str((i-rowList[0])) + " out of " + str(len(rowList)) +" (" + str(100*((i-rowList[0])/(len(rowList))))+"%)")
        requestString = dataList[i][0]
        
        splitString = requestString.split("00")
        splitString = splitString[1]
        requestString = splitString.split("-")
        requestString = requestString[0]

        
                
        URL = "https://cps.mycusthelp.com/webapp/_rs/(S(xsqtvcgoedb042oyzh1mctqm))/RequestArchiveDetails.aspx?rid=" + requestString
        try:
            page = requests.get(URL)

        except:
            waittimer += 0
            print("Encountered an error. Waiting a minute before resuming")
            print("Starting Cat Animation...")
            time.sleep(5)
            for i in range(352):
                catAnimation(i)
                time.sleep(0.16)
            print("Done Waiting")
            page = requests.get(URL)

        newPage = page.text
        startString = "<p style=\"font-weight: 400; max-width: 75%; font-size: 0.875rem\" tabindex=\"0\">"
        endString = "</p>"

        stringList = newPage.split(startString)
        breakLoop = "false"
        for p in range(1,6):
            if breakLoop == "true":
                break
            else:
                try:                
                    newString = stringSplit(stringList[p], endString, "first")
                    print(len(dataList[i]))
                    if len(dataList[i]) < 9:
                        dataList[i].append(newString)
                    else:
                        dataList[i][p+3] = newString
                        
                except IndexError:
                    print("Failed  on row " + str(i+2) + ". Program continuing...")
                    failList.append(i+2)
                    breakLoop = "true"
                    lastFailed = i
    #                raise SystemExit(0)
    
        if i%50==0:
            write2Excel(dataList)
            print("Progress Saved.")

    print("Failed on rows:")
    for i in failList:
        print(i)
        
    return dataList

def write2Excel(dataList):
    df = pd.DataFrame(dataList,columns=['Request Number','Create Date','Summary','Request Status','Date Received','Name of Requester','Record Description','Status','Date Complete'])
    df.to_excel("gridView.xlsx", index=False)
    
def stringSplit(string2split, splitString, placeString):
    subString = string2split.split(splitString)
    if placeString == "first":
        return subString[0]
    elif placeString == "last":
        return subString[1]
    else:
        print("SPLIT STRING ERROR")

def getRows():
    start = input("What row would you like to begin your request at?: ")
    end = input("What row would you like to end your request at?: ")

    start = int(start)-2
    end = int(end) - 1

    rowList = []
    for i in range(start, end):
        rowList.append(i)

    return rowList

def checkRows(dataList):
    failList = []
    for i in range(len(dataList)):
        wasFailed = 0
        for p in range(5):
            if str(dataList[i][4+p])== "nan":
                wasFailed += 1
        if wasFailed > 4:
            failList.append(i+2)

    return failList

def main():
    rowList = getRows()
    dataList = read_data()
    dataList = scrapeWeb(dataList,rowList) 
    write2Excel(dataList)

def failedRows():
    dataList = read_data()
    print(dataList[10][5])
    failList = checkRows(dataList)
    failList = checkConsecutive(failList)
    print("A total of "+str(len(failList))+" requests failed:")
    string = "Row "
    for i in range(len(failList)):
        string += str(failList[i])+", "
    print(string)
def catAnimation(i):
    print("""

























""")
    default = "       "
    if i % 16 == 0:
        catAngelLeft()
    if i % 16 == 1:
        catAngelLeft()

    if i % 16 == 2:
        spaces = default
        rightWalkUp(spaces)
    if i % 16 == 3:
        spaces = default*2
        rightWalkDown(spaces)
    if i % 16 == 4:
        spaces = default*3
        rightWalkUp(spaces)
    if i % 16 == 5:
        spaces = default*4
        rightWalkDown(spaces)
    if i % 16 == 6:
        spaces = default*5
        rightWalkUp(spaces)
    if i % 16 == 7:
        spaces = default*6
        rightWalkDown(spaces)

    if i % 16 == 8:
        catAngelRight()
    if i % 16 == 9:
        catAngelRight()

    if i % 16 == 10:
        spaces = default*6
        leftWalkUp(spaces)
    if i % 16 == 11:
        spaces = default*5
        leftWalkDown(spaces)
    if i % 16 == 12:
        spaces = default*4
        leftWalkUp(spaces)
    if i % 16 == 13:
        spaces = default*3
        leftWalkDown(spaces)
    if i % 16 == 14:
        spaces = default*2
        leftWalkUp(spaces)
    if i % 16 == 15:
        spaces = default*1
        leftWalkDown(spaces)

    print("-------------------------------------------------------------------")
    
def catAngelRight():
    print("""
                                                             (___) 
                                                      ____
                                                    _\___ \  |\_/|
                                                   \     \ \/ , , \\
                                                    \__   \ \ ="= /
                                                     |===  \/____)_)
                                                     \______|    |
                                                         _/_|  | |
                                                        (_/  \_)_)
     """)

def catAngelLeft():
    print("""
  (___)           
          ____
  |\_/|  / ___/_
 / , ,  \/ /     /
 \ ="= / /   __/
(_(____\/  ===|
  |    |______/
  | |  |_\_
  (_(_/  \_)  
     """)

def leftWalkUp(spaces):
    print("""
"""+spaces+"""
"""+spaces+"""
"""+spaces+"""   |\__/,|   (`\\
"""+spaces+"""   |, ,  |__ _)|
"""+spaces+""" __( T   )  `  /
"""+spaces+"""((_ `^--' /_<  \\
"""+spaces+"""   `-'(((/  (((/
""")

def rightWalkUp(spaces):
    print("""
"""+spaces+"""
"""+spaces+"""/')   |,\__/|
"""+spaces+"""|(_ __|  , ,|
"""+spaces+"""\  `  (   T )__
"""+spaces+"""/  >_\ '--^` _))
"""+spaces+"""\)))  \)))'-`
""")

def leftWalkDown(spaces):
    print("""
"""+spaces+"""
"""+spaces+"""
"""+spaces+"""   |\__/,|   (`\\
"""+spaces+"""   |, ,  |__ _)|
"""+spaces+"""   ( T __)  `  /
"""+spaces+"""  / `((__`_<  \\
"""+spaces+""" ((_/      (((/
""")

def rightWalkDown(spaces):
    print("""
"""+spaces+"""
"""+spaces+"""/')   |,\__/|
"""+spaces+"""|(_ __|  , ,|
"""+spaces+"""\  `  (__ T )
"""+spaces+""" /  >_`__))` \\
"""+spaces+""" \)))      \_))
""")



#for i in range(1000):
#    catAnimation(i)
#    time.sleep(0.16)

#failedRows()
main()
