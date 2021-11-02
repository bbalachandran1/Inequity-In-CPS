import requests
import pandas as pd
import html
import time
import socket
import errno


#Gets user input, finds data online, and saves it to 
def main():
    #asks the user which rows of the datasheet they'd like to start and end on
    rowList = getRows()

    dataList = read_data()

    dataList = scrapeWeb(dataList,rowList)

    write2Excel(dataList)


#Reads the data of the Excel Sheet using the pandas library and inserts it into a list
def read_data():
    filename = "[FOIA]FOIA Requests"

    #Reads the data in the Excel Sheet as a pandas DataFrame
    df = pd.read_excel(filename + ".xlsx"   )
    newDf = df.to_numpy()

    #Transforms the dataframe into a 2 dimensional list
    bigList = []
    for i in range(0,len(newDf)):
        miniList=[]
        for p in range(len(newDf[0])):
            miniList.append(newDf[i][p])

        bigList.append(miniList)

    return bigList


#Navigates the CPS Public Archive of FOIA requests to find necessary data
def scrapeWeb(dataList,rowList):
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
            #If the program is shut out of the connection, the cat animation starts (because who doesn't like cats :D)
            print("Encountered an error. Waiting a minute before resuming")
            print("Starting Cat Animation...")

            #waits 5 seconds before beginning, and then plays for 55 seconds
            time.sleep(5)
            
            for i in range(344):
                catAnimation(i)
                time.sleep(0.16)

            #Notifies the user that the program is ready to resume and tries the request again
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
                    if len(dataList[i]) < 9:
                        dataList[i].append(newString)
                    else:
                        dataList[i][p+3] = newString
                        
                except IndexError:
                    print("Failed  on row " + str(i+2) + ". Program continuing...")
                    failList.append(i+2)
                    breakLoop = "true"
                    lastFailed = i
    
        if i%50==0:
            write2Excel(dataList)
            print("Progress Saved.")

    print("Failed on rows:")
    for i in failList:
        print(i)
        
    return dataList

#Writes the data to Excel
def write2Excel(dataList):
    df = pd.DataFrame(dataList,columns=['Request Number','Create Date','Summary','Request Status','Date Received','Name of Requester','Record Description','Status','Date Complete'])
    df.to_excel("[FOIA]FOIA Requests.xlsx", index=False)
    
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


# finds rows that 
def failedRows():
    dataList = read_data()
    failList = checkRows(dataList)
    failList = checkConsecutive(failList)
    print("A total of "+str(len(failList))+" requests failed:")
    string = "Row "
    for i in range(len(failList)):
        string += str(failList[i])+", "
    print(string)

#Cat Animation!!! (This is not technically necessary,
#but definitely makes the program better (goes up until )
def catAnimation(i):
    #arbritary amount of space to "push" all other text up in the console for a cleaner, animation
    print("\n" *30)

    default = " " * 7
    index = i % 16
    if index < 8:
        spaces = default * (index-1)
    elif index > 9:
        spaces = default * (16-index)
    else:
        spaces = ""

    functionList = [catAngelLeft,catAngelLeft,rightWalkUp,rightWalkDown,rightWalkUp,rightWalkDown,rightWalkUp,rightWalkDown,catAngelRight,catAngelRight,leftWalkUp,leftWalkDown,leftWalkUp,leftWalkDown,leftWalkUp,leftWalkDown]
    functionList[index](spaces)
    print("-------------------------------------------------------------------")

def catAngelRight(spaces):
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

def catAngelLeft(spaces):
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

for i in range(1000):
    catAnimation(i)
    time.sleep(0.16)
    
#failedRows()
main()
